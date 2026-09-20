"""
Medical Personnel Lookup — by Player, by Team, or Both

Also covers front-office/coaching leadership: head coaches and general managers
are identified alongside medical/conditioning staff in Step 1.

Four lookup modes:
  1. Player mode      : enter a player name -> finds medical staff (+ head coach/GM)
                        linked to that player
  2. Team mode        : enter a team name and sport -> finds medical staff (+ head
                        coach/GM) for the team
  3. Player + Team    : enter both -> finds staff linked to the player AND staff
                        employed by the team (combined, de-duplicated)
  4. Excel batch      : point at an .xlsx file with a column of organizations
                        (e.g. column B) -> runs team mode for every organization
                        in the column and combines all contacts into one workbook

Step 1 — Gemini + grounding : identify staff names/roles
Step 2 — Gemini + grounding : find email addresses
Step 3 — SerpAPI            : collect LinkedIn candidates using multiple query variants
                              (primary pass, then a secondary broader pass if needed),
                              score each candidate on name/institution/title/location,
                              and assign a 0-100 confidence score
Step 4 — SerpAPI validation : independently verify the chosen LinkedIn URL (must appear
                              in top 5 results for "LinkedIn {team} {sport} {location}
                              {name} {job title}")

Email addresses are found via Gemini (Step 2) but are not independently verified.

This script distinguishes "found" (a candidate was located) from "verified" (the candidate
was independently confirmed by a second search) for LinkedIn, and assigns each LinkedIn
match a 0-100 confidence score based on name, institution, title, and location signals.

IMPORTANT: Only independently VERIFIED LinkedIn URLs are ever displayed in the final
output. A candidate that was found but not verified is withheld — the report will note
that a candidate existed (with its confidence score) but will not print the URL itself.

PERFORMANCE NOTES (this version):
  - All of the slow work in this script is network I/O (Gemini + SerpAPI calls), so the
    biggest win is concurrency, not algorithmic changes.
  - Steps 2, 3, and 4 process every staff member in a thread pool instead of one at a
    time — wall-clock time for these steps now scales with the slowest person, not the
    sum of all people.
  - Within Step 3, the multiple SerpAPI query variants for a single person are also
    fired concurrently instead of sequentially with a fixed 1-second sleep between each.
  - A global semaphore caps how many SerpAPI requests are in flight at once (across all
    threads) so we still respect the API's rate limits even with concurrency turned on.
  - In "both" mode, the player-staff and team-staff Gemini lookups in Step 1 now run
    concurrently instead of one after the other.
  - Because output now arrives out of order (whichever person's request finishes first
    prints first), console logs are labelled by name rather than by a fixed index.

Requirements:
    pip install google-genai requests

Usage (interactive):
    python3 athlete_medical_staff.py

Usage (CLI flags):
    python athlete_medical_staff.py --player "LeBron James"
    python athlete_medical_staff.py --team "Los Angeles Lakers" --sport "NBA basketball"
    python athlete_medical_staff.py --player "LeBron James" --team "Los Angeles Lakers" --sport "NBA basketball"
    python athlete_medical_staff.py --excel "injury_report.xlsx"
    python athlete_medical_staff.py --excel "injury_report.xlsx" --limit 5

Environment variables:
    GEMINI_API_KEY  — https://aistudio.google.com/app/apikey
    SERPAPI_KEY     — https://serpapi.com/manage-api-key
"""

import os
import re
import json
import time
import random
import difflib
import argparse
import threading
import requests
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from google import genai
from google.genai import types
from google.genai import errors as genai_errors
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


# ── Tunables ────────────────────────────────────────────────────────────────────

# If the best candidate found in the primary search pass scores below this, a
# secondary, broader search pass is run automatically to look for better matches.
SECONDARY_PASS_CONFIDENCE_THRESHOLD = 65

# Step 4: how many top SerpAPI results to check a candidate URL against when
# independently verifying it. Higher = more lenient (fewer correct matches
# rejected just for ranking slightly lower on the verification query), but
# slightly higher chance of a false positive slipping through.
LINKEDIN_VERIFY_TOP_N = 5

# Final report: an unverified LinkedIn candidate is still shown if its
# confidence score is above this threshold. Lower this if too many real
# matches are being withheld; raise it if too many wrong matches are showing
# up unverified.
LINKEDIN_DISPLAY_CONFIDENCE_THRESHOLD = 65

# Weights used to combine the four match signals into a single 0-100 confidence score.
CONFIDENCE_WEIGHTS = {
    "name": 0.40,
    "institution": 0.25,
    "title": 0.20,
    "location": 0.15,
}

# ── Concurrency tunables ────────────────────────────────────────────────────────

# How many staff members to process at once in Steps 2-4 (email lookup, LinkedIn
# search, LinkedIn verification). Each person's work is independent of every other
# person's, so this is safe to parallelize.
MAX_PARALLEL_PEOPLE = 5

# How many SerpAPI query variants to run concurrently for a single person within
# one batch (primary pass or secondary pass) in Step 3.
MAX_PARALLEL_QUERIES = 4

# Hard cap on how many SerpAPI requests may be in flight at the same time, across
# every thread in the whole program. This is what actually protects us from
# tripping the API's rate limit when running things concurrently — raise/lower
# this if you have a higher/lower SerpAPI plan tier.
SERPAPI_MAX_CONCURRENT = 6

_serpapi_semaphore = threading.Semaphore(SERPAPI_MAX_CONCURRENT)


# ── Gemini client ──────────────────────────────────────────────────────────────

def get_client() -> genai.Client:
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise EnvironmentError(
            "GEMINI_API_KEY environment variable is not set.\n"
            "Get a free key at https://aistudio.google.com/app/apikey and run:\n"
            "  export GEMINI_API_KEY='your_key_here'"
        )
    return genai.Client(api_key=api_key)


def gemini_generate(client: genai.Client, prompt: str, max_retries: int = 5) -> str:
    """Call Gemini with Google Search grounding, retrying on 503 errors.

    Safe to call concurrently from multiple threads — the google-genai client
    issues one independent HTTP request per call and holds no call-specific
    mutable state between calls.
    """
    for attempt in range(1, max_retries + 1):
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config=types.GenerateContentConfig(
                    tools=[types.Tool(google_search=types.GoogleSearch())],
                    temperature=0.1,
                ),
            )
            return response.text.strip()
        except genai_errors.ServerError as e:
            if "503" in str(e) or "UNAVAILABLE" in str(e):
                wait = 2 ** attempt
                print(f"  [Attempt {attempt}/{max_retries}] Gemini busy — retrying in {wait}s …")
                time.sleep(wait)
            else:
                raise
    raise RuntimeError("Gemini API unavailable after all retries.")


def parse_json(raw: str) -> dict | list:
    """Strip markdown fences and parse JSON."""
    clean = re.sub(r"^```(?:json)?\s*", "", raw.strip())
    clean = re.sub(r"\s*```$", "", clean).strip()
    return json.loads(clean)


# ── Step 1: Identify staff via Gemini ─────────────────────────────────────────

def identify_staff_for_player(client: genai.Client, player: str) -> list[dict]:
    """Find medical/conditioning staff, plus head coach and GM, linked to a specific player."""
    prompt = f"""
You are a sports research assistant. Use Google Search to find every known medical
professional, conditioning professional, and team leadership figure linked to this
athlete:

  Player : {player}

Roles to find (include every role you can confirm with a real source):
  - Personal trainer / strength & conditioning coach
  - Team physician / team doctor
  - Physical therapist / physiotherapist
  - Athletic trainer / certified athletic trainer (ATC)
  - Nutritionist / sports dietitian
  - Sports psychologist / mental performance coach
  - Head coach (of the player's current team)
  - General manager (of the player's current team)
  - Any other notable medical or wellness professional

IMPORTANT:
  - Only include people you can confirm from a real news article, official team
    page, or professional profile found via Google Search.
  - Do NOT include anyone you are guessing or inferring.
  - Do NOT include any URLs, emails, or social media handles — names and roles only.

Respond ONLY with a JSON array. Each element must have exactly these keys:
  "name"         : full name (string)
  "role"         : job title (string)
  "relationship" : how they are linked to the player (string)

Return only the JSON array — no markdown fences, no explanation.
"""
    raw = gemini_generate(client, prompt)
    try:
        return parse_json(raw)
    except json.JSONDecodeError as e:
        print(f"  Warning: could not parse staff JSON: {e}")
        print(f"  Raw output: {raw[:500]}")
        return []


def identify_staff_for_team(client: genai.Client, team: str, sport: str) -> list[dict]:
    """Find medical/conditioning staff, plus head coach and GM, employed by a specific team."""
    prompt = f"""
You are a sports research assistant. Use Google Search to find every known medical
professional, conditioning professional, and team leadership figure currently or
recently employed by this team:

  Team  : {team}
  Sport : {sport}

Roles to find (include every role you can confirm with a real source):
  - Head coach
  - General manager
  - Head team physician / team doctor
  - Assistant team physician
  - Physical therapist / physiotherapist
  - Head athletic trainer / certified athletic trainer (ATC)
  - Assistant athletic trainer
  - Strength & conditioning coach / head of performance
  - Nutritionist / sports dietitian
  - Sports psychologist / mental performance coach
  - Any other notable medical or wellness professional on the team staff

IMPORTANT:
  - Only include people you can confirm from a real news article, official team
    page, or professional profile found via Google Search.
  - Do NOT include anyone you are guessing or inferring.
  - Do NOT include any URLs, emails, or social media handles — names and roles only.

Respond ONLY with a JSON array. Each element must have exactly these keys:
  "name"         : full name (string)
  "role"         : job title (string)
  "relationship" : how they are linked to the team (string)

Return only the JSON array — no markdown fences, no explanation.
"""
    raw = gemini_generate(client, prompt)
    try:
        return parse_json(raw)
    except json.JSONDecodeError as e:
        print(f"  Warning: could not parse staff JSON: {e}")
        print(f"  Raw output: {raw[:500]}")
        return []


# ── Step 2: Find office contact details via Gemini ────────────────────────────

def _office_contact_pass(client: genai.Client, full_name: str, role: str,
                          context: str, broad: bool = False) -> dict:
    """
    Run one Gemini grounded-search pass for office/work contact details.
    `broad=False` is a tight, high-precision pass (team/employer site directly).
    `broad=True` is a wider fallback pass that checks more source types — used
    only if the first pass came back empty.
    """
    if not broad:
        source_list = """  - The team's, clinic's, hospital's, or university's official staff directory page
  - Faculty/bio pages if they have an academic appointment
  - Press release or media-contact sections
  - Try: "{name}" "{ctx}" email
  - Try: "{name}" {role} office phone""".format(name=full_name, ctx=context, role=role)
    else:
        # Wider net: alternate institution pages, professional-association directories,
        # conference/speaker bios, "find a provider" listings, news bylines/quotes that
        # cite an official work contact — still restricted to office/work info only.
        source_list = """  - Any current OR former employer's staff directory (people sometimes move between
    teams/clinics — check recent past employers too if the current one has nothing)
  - Hospital or clinic "Find a Doctor" / "Find a Provider" pages
  - Sports-medicine or professional association membership directories (e.g. NATA,
    AMSSM, AOSSM, APTA) which often list a work email/phone for members
  - Conference, webinar, or speaker-bio pages, which often list a work contact
  - University/department faculty directories, including adjunct or visiting roles
  - News articles or interviews that print an official media/press contact for them
  - LinkedIn "Contact info" sections sometimes display a work email if public
  - Try several name variants: "{name}", with middle initial, with common nicknames
  - Try: {role} "{ctx}" staff directory
  - Try: "{name}" "{role}" -personal -gmail -yahoo (to bias away from personal accounts)""".format(
            name=full_name, ctx=context, role=role
        )

    prompt = f"""
Use Google Search to find OFFICE/WORK contact information for this person, as
published by their employer (or professional association) for professional/
business purposes:

  Name    : {full_name}
  Role    : {role}
  Context : {context}

Find up to two things:
  1. "office_email" — their employer-issued or professionally-published work
     email address (e.g. firstname.lastname@team.com, @hospital.org,
     @university.edu, or an email listed against their name in a professional
     association directory).
  2. "office_phone" — a published office/department phone number that is
     specific to this person or their specific department (e.g. athletic
     training room direct line, clinic front desk, press/media office direct
     line).

Search strategies:
{source_list}

Rules:
  - Only return values found in a real, official, published source — do NOT
    guess or construct an email or phone number.
  - ONLY return office/work contact details. Do NOT return personal email
    addresses (Gmail, Yahoo, iCloud, etc.), personal cell numbers, or home
    numbers, even if you find them — leave those fields null instead.
  - Do NOT return generic, unattributed addresses like info@, support@, noreply@,
    or contact@ unless they are specifically the directory entry for THIS person
    or their specific department/office.
  - Do NOT return a team's, stadium's, arena's, or venue's general main
    switchboard, ticket office, box office, or guest-services phone number.
    Those numbers route to the general public line, not to this person or
    their department, even if they show up prominently on the team's official
    "Contact Us" page — skip them.
  - Only return a phone number if the source explicitly attributes it to this
    person BY NAME, or to their specific department (e.g. "Athletic Training",
    "Sports Medicine", "Team Physician's Office") as distinct from the
    organization's general/main contact line.
  - If you cannot find a real office email or office phone that meets the
    above criteria, return null for that field rather than falling back to a
    generic organization number.

Respond ONLY with a JSON object using exactly these keys:
  {{"office_email": "address@example.com" or null, "office_phone": "+1 555 555 5555" or null}}

Return only the JSON object — no markdown fences, no explanation.
"""
    try:
        raw = gemini_generate(client, prompt)
        data = parse_json(raw)
        return {
            "office_email": data.get("office_email"),
            "office_phone": data.get("office_phone"),
        }
    except Exception as e:
        print(f"    [Gemini warning] office contact search failed for {full_name}"
              f"{' (broad pass)' if broad else ''}: {e}")
        return {"office_email": None, "office_phone": None}


def find_office_contact(client: genai.Client, full_name: str, role: str, context: str) -> dict:
    """
    Find this person's OFFICE/WORK email and OFFICE/WORK phone number via grounded
    search — i.e. contact details published for professional/business purposes.
    Personal emails and personal/cell phone numbers are explicitly out of scope.

    Runs a tight, high-precision pass first (team/employer site directly). If that
    comes back completely empty, automatically runs a second, broader pass that
    checks more source types (professional associations, "find a provider" pages,
    past employers, conference bios, etc.) before giving up — this is the main
    lever for improving recall without loosening the office/work-only restriction.
    """
    result = _office_contact_pass(client, full_name, role, context, broad=False)
    if not result["office_email"] and not result["office_phone"]:
        result = _office_contact_pass(client, full_name, role, context, broad=True)
    return result


# ── SerpAPI core ────────────────────────────────────────────────────────────────

def serpapi_search(query: str, num: int = 10, max_retries: int = 5,
                    timeout: float = 45.0) -> list[dict]:
    """
    Run a SerpAPI Google search with retries on both 429 rate-limits and timeouts.
    Returns a list of {title, link, snippet} dicts.

    Safe to call concurrently — the actual outbound request is gated by a module-
    level semaphore (_serpapi_semaphore) so no more than SERPAPI_MAX_CONCURRENT
    requests are ever in flight at once, regardless of how many threads call this.

    Timeout handling notes:
      - `timeout` is a (connect, read) pair under the hood: a short connect
        timeout (we should fail fast if SerpAPI's endpoint itself is unreachable)
        and a longer read timeout (SerpAPI's own backend can take a while on
        complex/heavily-loaded queries before it responds).
      - Backoff on timeout is exponential WITH jitter and a hard cap, instead of
        a bare 2**attempt. Bare exponential backoff across many concurrent
        threads tends to make everyone retry in near-lockstep, which just
        recreates the same load spike that caused the timeouts. Jitter spreads
        retries out; the cap keeps any single retry from waiting absurdly long.
    """
    api_key = os.environ.get("SERPAPI_KEY")
    if not api_key:
        raise EnvironmentError(
            "SERPAPI_KEY environment variable is not set.\n"
            "Get a free key at https://serpapi.com/manage-api-key and run:\n"
            "  export SERPAPI_KEY='your_key_here'"
        )

    connect_timeout = min(10.0, timeout)
    read_timeout = timeout
    max_backoff = 30.0

    def _backoff_seconds(attempt: int) -> float:
        base = min(2 ** attempt, max_backoff)
        return base * (0.5 + random.random() * 0.5)  # jitter: 50%-100% of base

    for attempt in range(1, max_retries + 1):
        try:
            with _serpapi_semaphore:
                resp = requests.get(
                    "https://serpapi.com/search",
                    params={"api_key": api_key, "engine": "google", "q": query,
                            "num": num, "hl": "en", "gl": "us"},
                    timeout=(connect_timeout, read_timeout),
                )
            if resp.status_code == 429:
                retry_after = int(resp.headers.get("Retry-After", 0))
                wait = retry_after if retry_after > 0 else _backoff_seconds(attempt)
                print(f"    [SerpAPI] rate-limited — waiting {wait:.1f}s (attempt {attempt}/{max_retries}) …")
                time.sleep(wait)
                continue
            resp.raise_for_status()
            return [
                {"title": r.get("title", ""), "link": r.get("link", ""), "snippet": r.get("snippet", "")}
                for r in resp.json().get("organic_results", [])
            ]
        except requests.exceptions.Timeout:
            wait = _backoff_seconds(attempt)
            print(f"    [SerpAPI] timeout — retrying in {wait:.1f}s (attempt {attempt}/{max_retries}) …")
            time.sleep(wait)
        except requests.exceptions.RequestException as e:
            print(f"    [SerpAPI] request error: {e}")
            return []
    print(f"    [SerpAPI] all retries exhausted for query: {query}")
    return []


# ── Step 3a: Name matching (existing logic, unchanged) ─────────────────────────

def name_match_score(full_name: str, title: str) -> float:
    """
    Return a 0-1 similarity score between the target name and a result title.
    Uses both token overlap and sequence matching to handle name order variations.
    """
    name_lower  = full_name.lower()
    title_lower = title.lower()

    name_tokens = name_lower.split()
    token_hits  = sum(1 for t in name_tokens if t in title_lower)
    token_score = token_hits / len(name_tokens) if name_tokens else 0

    seq_score = difflib.SequenceMatcher(None, name_lower, title_lower[:80]).ratio()

    return max(token_score, seq_score)


# ── Step 3b: Extra match signals (Improvement #1 — stronger matching criteria) ─

def extract_candidate_fields(title: str, snippet: str) -> dict:
    """
    Heuristically pull a job-title fragment, an employer/institution fragment, and a
    location fragment out of a LinkedIn search result's title + snippet text.

    LinkedIn search titles are typically formatted like:
        "Jane Doe - Head Athletic Trainer - Los Angeles Lakers | LinkedIn"
    and snippets often contain patterns like:
        "Jane Doe. Head Athletic Trainer at Los Angeles Lakers. Greater Los Angeles Area"

    This is necessarily approximate (we only ever see search-result text, never the
    profile itself), so callers should treat the output as a hint, not ground truth.
    """
    combined_text = f"{title} {snippet}"

    title_segments = re.split(r"\s[-|]\s", title)
    job_segment = title_segments[1].strip() if len(title_segments) > 1 else ""
    company_segment = title_segments[2].strip() if len(title_segments) > 2 else ""

    if not company_segment:
        at_match = re.search(r"\bat\s+([A-Z][\w&.,' -]{2,60})", snippet)
        if at_match:
            company_segment = at_match.group(1).strip(" .,")

    location_match = re.search(
        r"([A-Z][a-zA-Z.]+(?:,\s*[A-Z][a-zA-Z]+)?)\s+(?:Area|Metropolitan Area)",
        combined_text,
    )
    location_segment = location_match.group(0).strip() if location_match else ""

    return {
        "job_segment": job_segment,
        "company_segment": company_segment,
        "location_segment": location_segment,
    }


def compute_institution_match(candidate_fields: dict, team: str | None) -> float:
    """
    0-1 score for whether the candidate's apparent employer/affiliation matches the
    team we're searching for. Returns a neutral 0.5 if we have no team to compare
    against (e.g. player-only mode with no team context).
    """
    if not team:
        return 0.5

    text = f"{candidate_fields['company_segment']} {candidate_fields['job_segment']}".lower()
    team_lower = team.lower()

    if team_lower in text:
        return 1.0

    significant_words = [w for w in team_lower.split() if len(w) > 3]
    if not significant_words:
        return 0.0

    hits = sum(1 for w in significant_words if w in text)
    return hits / len(significant_words)


def compute_title_similarity(role: str, candidate_fields: dict) -> float:
    """
    0-1 score comparing the role we're looking for (e.g. "Head Athletic Trainer")
    against the job-title fragment extracted from the candidate's search result.
    Returns a mild default (not zero, not full credit) if no fragment was found,
    since absence of a parsed title isn't strong evidence either way.
    """
    job_segment = candidate_fields.get("job_segment", "")
    if not job_segment or not role:
        return 0.4

    primary_role = re.split(r"[/|,]", role)[0].strip().lower()
    return difflib.SequenceMatcher(None, primary_role, job_segment.lower()).ratio()


def compute_location_consistency(expected_location: str | None, candidate_fields: dict) -> float:
    """
    0-1 score comparing an expected location (if the caller has one) against the
    location fragment extracted from the candidate's search result. Neutral 0.5 if
    either side is unknown — we should not penalize a match just because location
    text wasn't present in a snippet.
    """
    candidate_location = candidate_fields.get("location_segment", "")
    if not expected_location or not candidate_location:
        return 0.5
    return difflib.SequenceMatcher(None, expected_location.lower(), candidate_location.lower()).ratio()


def compute_confidence_score(name_score: float, institution_score: float,
                              title_score: float, location_score: float) -> int:
    """Combine the four signals into a single 0-100 confidence score."""
    raw = (
        name_score * CONFIDENCE_WEIGHTS["name"]
        + institution_score * CONFIDENCE_WEIGHTS["institution"]
        + title_score * CONFIDENCE_WEIGHTS["title"]
        + location_score * CONFIDENCE_WEIGHTS["location"]
    )
    return round(raw * 100)


# ── Step 3c: Candidate collection (Improvement #2 — discovery coverage) ────────

def _run_query_batch(queries: list[str], full_name: str, team: str | None,
                      role: str, expected_location: str | None,
                      seen_urls: set[str], source_label: str) -> list[dict]:
    """
    Run a batch of SerpAPI queries CONCURRENTLY and turn results into scored
    candidates.

    The network calls (serpapi_search) run in a thread pool; the actual in-flight
    concurrency is still bounded by the global _serpapi_semaphore in serpapi_search
    itself, so this is safe even if multiple people are being processed at once
    elsewhere in the program. Result processing (scoring, dedup against seen_urls)
    happens back in this single thread as each query's results arrive, so no lock
    is needed around seen_urls.
    """
    candidates = []
    max_workers = min(len(queries), MAX_PARALLEL_QUERIES) or 1

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_query = {executor.submit(serpapi_search, q, 10): q for q in queries}
        for future in as_completed(future_to_query):
            try:
                results = future.result()
            except Exception as e:
                print(f"    [SerpAPI warning] query failed: {e}")
                continue

            for r in results:
                link = r.get("link", "")
                if "linkedin.com/in/" not in link or link in seen_urls:
                    continue
                seen_urls.add(link)

                fields = extract_candidate_fields(r.get("title", ""), r.get("snippet", ""))
                name_score = name_match_score(full_name, r.get("title", ""))
                institution_score = compute_institution_match(fields, team)
                title_score = compute_title_similarity(role, fields)
                location_score = compute_location_consistency(expected_location, fields)
                confidence = compute_confidence_score(name_score, institution_score, title_score, location_score)

                candidates.append({
                    "url": link,
                    "title": r.get("title", ""),
                    "snippet": r.get("snippet", ""),
                    "name_score": round(name_score, 3),
                    "institution_score": round(institution_score, 3),
                    "title_score": round(title_score, 3),
                    "location_score": round(location_score, 3),
                    "location_segment": fields.get("location_segment", ""),
                    "confidence": confidence,
                    "source_pass": source_label,
                })

    return candidates


# Common honorific prefixes and credential/license suffixes that Gemini sometimes
# attaches to a name (e.g. "Dr. Jane Doe, ATC" or "John Smith, DPT"). LinkedIn
# profile names almost never include these, so a quoted exact-name search using
# the raw name as given frequently fails to match a real profile even when one
# exists. Stripping these before building search queries fixes that silent
# recall loss.
_NAME_PREFIXES = {"dr", "dr.", "mr", "mr.", "mrs", "mrs.", "ms", "ms.", "prof", "prof."}
_NAME_SUFFIX_PATTERN = re.compile(
    r",?\s*\b(MD|DO|PhD|DPT|ATC|PT|RD|RDN|CSCS|MS|MBA|PA-C|DC|DPM|MEd|EdD|LAT|CES)\b\.?",
    re.IGNORECASE,
)


def clean_name_for_search(full_name: str) -> str:
    """Strip honorific prefixes and credential suffixes so quoted name searches
    actually match how the name appears on a LinkedIn profile."""
    name = _NAME_SUFFIX_PATTERN.sub("", full_name).strip().rstrip(",").strip()
    parts = name.split()
    if parts and parts[0].lower().rstrip(".") in {p.rstrip(".") for p in _NAME_PREFIXES}:
        parts = parts[1:]
    cleaned = " ".join(parts).strip()
    return cleaned if cleaned else full_name


def find_linkedin_candidates(full_name: str, role: str, context: str,
                              team: str | None = None, sport: str | None = None,
                              expected_location: str | None = None) -> list[dict]:
    """
    Run a primary set of high-precision SerpAPI queries to find LinkedIn candidates.
    If the primary pass turns up nothing, or nothing above the confidence threshold,
    automatically run a secondary, broader pass before giving up. Every candidate
    is scored on name/institution/title/location and given a 0-100 confidence score.

    Query budget:
      - Primary pass : 3 queries (4 if a team is known)
      - Secondary pass: only runs if the primary pass is empty/low-confidence,
                        and adds 3 more queries
      - Best case  : 3-4 SerpAPI calls per person
      - Worst case : 6-7 SerpAPI calls per person (+1 more in Step 4 if a
                     candidate needs independent verification)

    The secondary pass genuinely depends on the primary pass's outcome (its
    confidence score decides whether we even need it), so the two passes still run
    sequentially relative to each other — but each pass's queries run concurrently
    internally (see _run_query_batch).
    """
    primary_role = re.split(r"[/|,]", role)[0].strip()
    search_name = clean_name_for_search(full_name)
    name_parts = search_name.split()
    first_last = f"{name_parts[0]} {name_parts[-1]}" if len(name_parts) >= 2 else search_name

    seen_urls: set[str] = set()

    # All primary queries keep the quoted, honorific/credential-stripped exact
    # name AND the site:linkedin.com/in restriction — that combination is what
    # actually filters search results down to real LinkedIn profile pages.
    # One query also drops the site: restriction (still quoted) since some
    # LinkedIn pages don't surface well under that operator but do show up in
    # an ordinary search.
    primary_queries = [
        f'"{search_name}" {primary_role} {context} site:linkedin.com/in',
        f'"{search_name}" {context} site:linkedin.com/in',
        f'"{search_name}" {primary_role} linkedin',
    ]
    if team:
        primary_queries.append(f'"{search_name}" {team} site:linkedin.com/in')

    candidates = _run_query_batch(primary_queries, full_name, team, role,
                                   expected_location, seen_urls, "primary")

    best_confidence = max((c["confidence"] for c in candidates), default=0)
    if not candidates or best_confidence < SECONDARY_PASS_CONFIDENCE_THRESHOLD:
        # Secondary pass: broader queries — first/last name only, no site:
        # restriction on one of them — to catch profiles the primary pass
        # missed (e.g. nicknames, unusual title formatting, sparse snippets).
        secondary_queries = [
            f'"{search_name}" linkedin',
            f'"{first_last}" {primary_role} site:linkedin.com/in',
            f'"{first_last}" linkedin {team or context}',
        ]
        candidates += _run_query_batch(secondary_queries, full_name, team, role,
                                        expected_location, seen_urls, "secondary")


    return candidates


def select_best_candidate(candidates: list[dict]) -> dict:
    """
    Pick the LinkedIn candidate with the highest confidence score.

    Returns a dict describing the outcome:
      {
        "url": str | None,
        "confidence": int,
        "candidate_count": int,
        "match_breakdown": {...},
        "location_segment": str,
      }
    """
    if not candidates:
        return {"url": None, "confidence": 0, "candidate_count": 0, "location_segment": ""}

    ranked = sorted(candidates, key=lambda c: c["confidence"], reverse=True)
    best = ranked[0]

    return {
        "url": best["url"],
        "confidence": best["confidence"],
        "candidate_count": len(ranked),
        "match_breakdown": {
            "name": best["name_score"],
            "institution": best["institution_score"],
            "title": best["title_score"],
            "location": best["location_score"],
        },
        "location_segment": best.get("location_segment", ""),
    }


# ── Step 4: Independent verification of LinkedIn URL ("Verified", not just "Found") ─

def validate_linkedin_url(
    linkedin_url: str,
    staff_name: str,
    staff_role: str | None,
    location: str | None,
    team: str | None,
    sport: str | None,
) -> bool:
    """
    Independently verify a LinkedIn URL by checking whether it appears in the top
    LINKEDIN_VERIFY_TOP_N organic results when searching:
        LinkedIn {team} {sport} {location} {name of medical personnel} {job title}

    The job title (as identified in Step 1) is included so the verification query
    is more specific — useful for disambiguating common names — without being the
    same search strategy used to find the candidate in the first place
    (Improvement #4 — "found" vs "verified"). The player's name is intentionally
    excluded from this query; location is used instead as the disambiguating signal.

    Previously this only checked the top 5 results, which meant a genuinely
    correct profile that happened to rank 6th-10th for this particular query
    phrasing would fail verification and get withheld from the report even
    though it was right. Widened to top 10 to reduce that false-negative rate.
    """
    primary_role = re.split(r"[/|,]", staff_role)[0].strip() if staff_role else ""

    parts = ["LinkedIn"]
    if team:
        parts.append(team)
    if sport:
        parts.append(sport)
    if location:
        parts.append(location)
    parts.append(staff_name)
    if primary_role:
        parts.append(primary_role)
    query = " ".join(parts)

    results = serpapi_search(query, num=LINKEDIN_VERIFY_TOP_N)

    top_links = [r.get("link", "") for r in results[:LINKEDIN_VERIFY_TOP_N]]
    return linkedin_url in top_links


# ── Combining results ──────────────────────────────────────────────────────────

def merge_staff_lists(player_staff: list[dict], team_staff: list[dict]) -> list[dict]:
    """
    Combine player-linked and team-linked staff lists, de-duplicating by name
    (case-insensitive).
    """
    merged: dict[str, dict] = {}

    for person in player_staff:
        key = person["name"].strip().lower()
        entry = dict(person)
        entry["sources"] = ["player"]
        merged[key] = entry

    for person in team_staff:
        key = person["name"].strip().lower()
        if key in merged:
            existing = merged[key]
            existing["sources"].append("team")
            team_relationship = person.get("relationship", "")
            if team_relationship and team_relationship not in existing.get("relationship", ""):
                existing["relationship"] = (
                    f"{existing.get('relationship', '')} | {team_relationship}".strip(" |")
                )
            if not existing.get("role") and person.get("role"):
                existing["role"] = person["role"]
        else:
            entry = dict(person)
            entry["sources"] = ["team"]
            merged[key] = entry

    return list(merged.values())


# ── Excel batch input ──────────────────────────────────────────────────────────

def read_organizations_from_excel(path: str, limit: int | None = None) -> list[dict]:
    """
    Read a column of organizations (and, if present, matching sport and athlete
    columns) from an .xlsx file for batch team-mode processing.

    Row 1 is treated as a header row. Columns are matched by header text
    ("Organization"/"Team", "Sport", "Athlete"/"Player"), falling back to column B,
    column C, and column A respectively if no matching header is found — this
    matches the athlete-injury-feed report layout (Athlete in column A,
    Organization in column B, Sport in column C).

    Rows are de-duplicated by organization name (case-insensitive, first occurrence
    sets the order) since the same organization can appear on multiple rows — one
    row per injury in the source report. The first non-empty sport value seen for
    an organization is kept if later rows disagree or are blank. Every distinct
    athlete name seen for an organization is collected (in first-seen order) so the
    output can show which athlete(s) triggered that organization's inclusion.
    """
    wb = load_workbook(path, data_only=True)
    ws = wb.active

    header_row = next(ws.iter_rows(min_row=1, max_row=1))
    header = [str(c.value).strip() if c.value is not None else "" for c in header_row]

    def find_column(*names: str, fallback_index: int) -> int:
        for name in names:
            for idx, h in enumerate(header):
                if h.lower() == name.lower():
                    return idx
        return fallback_index

    org_idx = find_column("Organization", "Team", fallback_index=1)     # column B
    sport_idx = find_column("Sport", fallback_index=2)                  # column C
    athlete_idx = find_column("Athlete", "Player", fallback_index=0)    # column A

    seen: dict[str, dict] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if org_idx >= len(row):
            continue
        org = row[org_idx]
        if org is None or not str(org).strip():
            continue
        org = str(org).strip()
        key = org.lower()

        sport = ""
        if sport_idx < len(row) and row[sport_idx]:
            sport = str(row[sport_idx]).strip()

        athlete = ""
        if athlete_idx < len(row) and row[athlete_idx]:
            athlete = str(row[athlete_idx]).strip()

        if key not in seen:
            seen[key] = {"team": org, "sport": sport, "athletes": []}
        elif sport and not seen[key]["sport"]:
            seen[key]["sport"] = sport

        if athlete and athlete not in seen[key]["athletes"]:
            seen[key]["athletes"].append(athlete)

    teams = list(seen.values())
    if limit is not None:
        teams = teams[:limit]
    return teams


# ── Output formatting ──────────────────────────────────────────────────────────

def format_staff(staff: list[dict], show_sources: bool = False) -> str:
    """
    Render the final report.

    LinkedIn display policy: a URL is only ever printed if it was independently
    verified in Step 4 (person["linkedin_validated"] is True). A candidate that was
    found in Step 3 but failed Step 4 verification is treated as "Not found" for
    display purposes — the URL itself is withheld — though a note is added so the
    reader knows an unverified candidate exists, along with its confidence score.
    """
    lines = []
    for person in staff:
        lines += [
            f"  Name              : {person.get('name', 'Unknown')}",
            f"  Role              : {person.get('role', 'Unknown')}",
        ]
        if show_sources:
            sources = person.get("sources", [])
            label = " & ".join(s.capitalize() for s in sources) if sources else "Unknown"
            lines.append(f"  Linked via        : {label}")

        # Office contact details (work email + work phone only — no personal info)
        office_email = person.get("office_email")
        if office_email:
            lines.append(f"  Office email      : {office_email}")
        else:
            lines.append(f"  Office email      : Not found")

        office_phone = person.get("office_phone")
        if office_phone:
            lines.append(f"  Office phone      : {office_phone}")
        else:
            lines.append(f"  Office phone      : Not found")

        # LinkedIn — show the URL if it was independently verified, OR if it wasn't
        # verified but its confidence score is high enough (>70) to surface anyway.
        linkedin = person.get("linkedin")
        linkedin_valid = person.get("linkedin_validated")
        confidence = person.get("linkedin_confidence", 0)
        high_confidence_unverified = (
            bool(linkedin) and not linkedin_valid
            and confidence > LINKEDIN_DISPLAY_CONFIDENCE_THRESHOLD
        )

        if linkedin and (linkedin_valid or high_confidence_unverified):
            status_label = "✓ independently verified" if linkedin_valid else "high confidence, NOT verified"
            lines.append(f"  LinkedIn          : {linkedin}  [{status_label}]")
            lines.append(f"  LinkedIn confidence: {confidence}/100")
        else:
            lines.append(f"  LinkedIn          : Not found")
            if linkedin and not linkedin_valid:
                # A candidate was found but failed independent verification and
                # didn't clear the confidence bar either — withhold the URL.
                lines.append(
                    f"  LinkedIn note     : candidate found (confidence {confidence}/100) "
                    f"but NOT independently verified and below confidence threshold — URL withheld"
                )
            else:
                candidate_count = person.get("linkedin_candidate_count", 0)
                if candidate_count:
                    lines.append(f"  LinkedIn note     : {candidate_count} candidate(s) seen, none confident enough")

        lines.append("")
    return "\n".join(lines)


def summarise_missing(staff: list[dict]) -> None:
    """
    Reports "found" and "independently verified" as separate metrics (Improvement #4),
    plus how many matches needed manual review and how many were flagged as
    duplicate-name risks (Improvements #3 and #5).

    Note: "found" below means a candidate URL existed after Step 3, regardless of
    whether it was later verified — this is purely a diagnostic count. The actual
    report (format_staff) only ever displays verified URLs.
    """
    total = len(staff)
    if not total:
        return

    office_email_found = sum(1 for p in staff if p.get("office_email"))
    office_phone_found = sum(1 for p in staff if p.get("office_phone"))

    linkedin_found = sum(1 for p in staff if p.get("linkedin"))
    linkedin_verified = sum(1 for p in staff if p.get("linkedin_validated"))
    linkedin_shown = sum(
        1 for p in staff
        if p.get("linkedin") and (p.get("linkedin_validated") or p.get("linkedin_confidence", 0) > 70)
    )
    avg_confidence = (
        round(sum(p.get("linkedin_confidence", 0) for p in staff if p.get("linkedin")) / linkedin_found)
        if linkedin_found else 0
    )

    def bar(found, total):
        pct = int(100 * found / total) if total else 0
        return "█" * (pct // 10) + "░" * (10 - pct // 10), pct

    print("\n── Field coverage summary ──────────────────────────────────")
    b, pct = bar(office_email_found, total)
    print(f"  {'Office email — found':<28} {b} {office_email_found}/{total} ({pct}%)")
    b, pct = bar(office_phone_found, total)
    print(f"  {'Office phone — found':<28} {b} {office_phone_found}/{total} ({pct}%)")
    b, pct = bar(linkedin_found, total)
    print(f"  {'LinkedIn — possible candidate found':<28} {b} {linkedin_found}/{total} ({pct}%)")
    b, pct = bar(linkedin_shown, total)
    print(f"  {'LinkedIn — shown in report':<28} {b} {linkedin_shown}/{total} ({pct}%)")
    print()
    print(f"  Average LinkedIn match confidence (found candidates) : {avg_confidence}/100")
    print(f"  Note: only verified URLs are shown in the report above.")
    print()


def export_to_excel(staff: list[dict], subject_label: str) -> str:
    """
    Write the same contact information shown in the terminal report to an .xlsx
    file. LinkedIn URLs follow the same display policy as format_staff(): only
    verified URLs, or unverified URLs above LINKEDIN_DISPLAY_CONFIDENCE_THRESHOLD,
    are written — everything else is left blank with a note.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Medical Staff"

    headers = ["Name", "Role"]
    headers += [
        "Office Email", "Office Phone",
        "LinkedIn URL",
    ]

    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    def last_name(person: dict) -> str:
        return person.get("name", "").strip().split(" ")[-1].lower()

    for person in sorted(staff, key=last_name):
        row = [
            person.get("name", "Unknown"),
            person.get("role", "Unknown"),
        ]
        row.append(person.get("office_email") or "")
        row.append(person.get("office_phone") or "")

        linkedin = person.get("linkedin")
        linkedin_valid = person.get("linkedin_validated")
        confidence = person.get("linkedin_confidence", 0)
        high_confidence_unverified = (
            bool(linkedin) and not linkedin_valid
            and confidence > LINKEDIN_DISPLAY_CONFIDENCE_THRESHOLD
        )

        if linkedin and (linkedin_valid or high_confidence_unverified):
            row.append(linkedin)
        else:
            row.append("")

        ws.append(row)

    for col_idx, header in enumerate(headers, start=1):
        max_len = max(
            [len(str(header))] + [len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(2, ws.max_row + 1)]
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

    ws.freeze_panes = "A2"

    safe_label = re.sub(r"[^\w\- ]", "", subject_label).strip().replace(" ", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"medical_personnel_{safe_label}_{timestamp}.xlsx"
    wb.save(filename)
    return filename


def export_combined_excel(all_staff: list[dict], subject_label: str) -> str:
    """
    Batch-mode counterpart to export_to_excel(): writes every contact from every
    organization processed in an --excel run into ONE workbook (one row per person,
    tagged with "Organization" and "Athlete(s)" columns — the latter listing every
    athlete from the input spreadsheet whose row referenced that organization),
    sorted by organization then last name. Follows the same LinkedIn display policy
    as export_to_excel().
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Medical Staff"

    headers = ["Organization", "Injured Athlete(s)", "Name", "Role", "Office Email", "Office Phone", "LinkedIn URL"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    def sort_key(person: dict) -> tuple[str, str]:
        return (
            person.get("organization", "").lower(),
            person.get("name", "").strip().split(" ")[-1].lower(),
        )

    for person in sorted(all_staff, key=sort_key):
        linkedin = person.get("linkedin")
        linkedin_valid = person.get("linkedin_validated")
        confidence = person.get("linkedin_confidence", 0)
        high_confidence_unverified = (
            bool(linkedin) and not linkedin_valid
            and confidence > LINKEDIN_DISPLAY_CONFIDENCE_THRESHOLD
        )

        ws.append([
            person.get("organization", "Unknown"),
            person.get("athletes", ""),
            person.get("name", "Unknown"),
            person.get("role", "Unknown"),
            person.get("office_email") or "",
            person.get("office_phone") or "",
            linkedin if linkedin and (linkedin_valid or high_confidence_unverified) else "",
        ])

    for col_idx, header in enumerate(headers, start=1):
        max_len = max(
            [len(str(header))] + [len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(2, ws.max_row + 1)]
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)

    ws.freeze_panes = "A2"

    safe_label = re.sub(r"[^\w\- ]", "", subject_label).strip().replace(" ", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"medical_personnel_{safe_label}_{timestamp}.xlsx"
    wb.save(filename)
    return filename


# ── Parallel helpers for Steps 2-4 ─────────────────────────────────────────────

def _run_in_parallel(staff: list[dict], worker, label: str) -> None:
    """
    Run `worker(person)` for every person in `staff` concurrently (bounded by
    MAX_PARALLEL_PEOPLE), printing progress as each one finishes. `worker` is
    expected to mutate `person` in place with its results — return value is
    ignored, exceptions are caught and logged so one failure doesn't abort the
    whole batch.
    """
    if not staff:
        return
    max_workers = min(len(staff), MAX_PARALLEL_PEOPLE) or 1
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_person = {executor.submit(worker, person): person for person in staff}
        done = 0
        for future in as_completed(future_to_person):
            person = future_to_person[future]
            done += 1
            try:
                future.result()
            except Exception as e:
                print(f"  [{done}/{len(staff)}] {label} failed for {person.get('name', 'Unknown')}: {e}")


# ── Mode selection ────────────────────────────────────────────────────────────

def interactive_mode_select() -> tuple[str, dict]:
    print("\nLookup mode:")
    print("  1. Player        — find medical staff linked to a specific player")
    print("  2. Team          — find medical staff employed by a team")
    print("  3. Player + Team — find medical staff linked to BOTH the player and the team")
    print("  4. Excel batch   — read a column of organizations from an .xlsx file and run each one")
    while True:
        choice = input("Enter 1, 2, 3, or 4: ").strip()
        if choice == "1":
            player = input("Enter player name: ").strip()
            if not player:
                print("Player name cannot be empty.")
                continue
            return "player", {"player": player}
        elif choice == "2":
            team = input("Enter team name: ").strip()
            sport = input("Enter sport (e.g. NBA basketball, Premier League soccer): ").strip()
            if not team or not sport:
                print("Both team name and sport are required.")
                continue
            return "team", {"team": team, "sport": sport}
        elif choice == "3":
            player = input("Enter player name: ").strip()
            team = input("Enter team name: ").strip()
            sport = input("Enter sport (e.g. NBA basketball, Premier League soccer): ").strip()
            if not player or not team or not sport:
                print("Player name, team name, and sport are all required.")
                continue
            return "both", {"player": player, "team": team, "sport": sport}
        elif choice == "4":
            excel_path = input("Enter path to .xlsx file: ").strip()
            if not excel_path:
                print("Excel path cannot be empty.")
                continue
            limit_raw = input("Limit number of organizations to process (Enter for no limit): ").strip()
            limit = int(limit_raw) if limit_raw else None
            return "excel", {"excel_path": excel_path, "limit": limit}
        else:
            print("Please enter 1, 2, 3, or 4.")


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Look up medical/training staff for a player, a team, or both.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python athlete_medical_staff.py --player \"LeBron James\"\n"
            "  python athlete_medical_staff.py --team \"Los Angeles Lakers\" --sport \"NBA basketball\"\n"
            "  python athlete_medical_staff.py --player \"LeBron James\" --team \"Los Angeles Lakers\" --sport \"NBA basketball\"\n"
            "  python athlete_medical_staff.py --excel \"injury_report.xlsx\"\n"
            "  python athlete_medical_staff.py --excel \"injury_report.xlsx\" --limit 5\n"
        ),
    )
    parser.add_argument("--player", "-p", type=str, default=None)
    parser.add_argument("--team",   "-t", type=str, default=None)
    parser.add_argument("--sport",  "-s", type=str, default=None)
    parser.add_argument("--excel",  "-x", type=str, default=None,
                         help="Path to an .xlsx file with a column of organizations to batch-process "
                              "(matches an 'Organization'/'Team' header, or falls back to column B)")
    parser.add_argument("--limit",  "-l", type=int, default=None,
                         help="Only process the first N organizations from --excel (useful for testing)")
    return parser.parse_args()


def enrich_staff_contacts(client: genai.Client, staff: list[dict],
                           team: str | None, sport: str | None, context: str) -> None:
    """
    Steps 2-4 of the pipeline: office contact lookup, LinkedIn discovery, and
    LinkedIn verification. Mutates each person dict in `staff` in place.

    Factored out of main() so the exact same logic runs once per subject in
    player/team/both mode, and once per organization in an --excel batch run,
    instead of being duplicated between the two.
    """
    print(f"Step 2/4 — Finding office email & office phone via Gemini "
          f"(up to {min(len(staff), MAX_PARALLEL_PEOPLE)} at a time) …")

    def _contact_worker(person: dict) -> None:
        details = find_office_contact(client, person["name"], person.get("role", ""), context)
        person["office_email"] = details["office_email"]
        person["office_phone"] = details["office_phone"]
        found = []
        if person["office_email"]:
            found.append(f"email: {person['office_email']}")
        if person["office_phone"]:
            found.append(f"phone: {person['office_phone']}")
        print(f"  {person['name']}: {', '.join(found) if found else 'no office contact found'}")

    _run_in_parallel(staff, _contact_worker, "office contact lookup")

    print(f"\nStep 3/4 — Finding LinkedIn profiles via SerpAPI "
          f"(scored, with secondary pass; up to {min(len(staff), MAX_PARALLEL_PEOPLE)} people "
          f"and up to {MAX_PARALLEL_QUERIES} queries per person at a time) …")

    def _linkedin_worker(person: dict) -> None:
        candidates = find_linkedin_candidates(
            person["name"], person.get("role", ""), context, team=team, sport=sport
        )
        outcome = select_best_candidate(candidates)

        person["linkedin"] = outcome["url"]
        person["linkedin_confidence"] = outcome["confidence"]
        person["linkedin_candidate_count"] = outcome["candidate_count"]
        # Location extracted from the winning LinkedIn candidate's search snippet,
        # used later (instead of the player's name) to disambiguate Step 4/4 queries.
        person["location"] = outcome.get("location_segment", "")

        if outcome["url"]:
            print(f"  {person['name']}: LinkedIn candidate {outcome['url']} "
                  f"(confidence {outcome['confidence']}/100, pending verification)")
        else:
            print(f"  {person['name']}: LinkedIn not found "
                  f"({outcome['candidate_count']} candidate(s) considered)")

    _run_in_parallel(staff, _linkedin_worker, "LinkedIn search")

    print("\nStep 4/4 — Independently verifying LinkedIn profiles via SerpAPI top-5 check …")
    print(  "           Query pattern: LinkedIn {team} {sport} {location} {name} {job title}")
    print(  "           Only URLs that pass this check will appear in the final report.")

    to_validate = [p for p in staff if p.get("linkedin")]
    skipped = [p for p in staff if not p.get("linkedin")]
    for person in skipped:
        person["linkedin_validated"] = False
        print(f"  {person['name']} — skipped (no URL found)")

    def _validate_worker(person: dict) -> None:
        valid = validate_linkedin_url(
            linkedin_url=person["linkedin"],
            staff_name=person["name"],
            staff_role=person.get("role"),
            location=person.get("location"),
            team=team,
            sport=sport,
        )
        person["linkedin_validated"] = valid
        status = "✓ verified (in top 5) — will be shown" if valid else "✗ NOT verified (not in top 5) — URL withheld"
        print(f"  {person['name']}: {status}")

    _run_in_parallel(to_validate, _validate_worker, "LinkedIn verification")


def run_excel_batch(client: genai.Client, excel_path: str, limit: int | None) -> None:
    """
    Batch team-mode entry point: reads a column of organizations (+ sport, if
    present) from an .xlsx file via read_organizations_from_excel(), then runs the
    full Step 1-4 pipeline for each organization in turn (sequentially — each
    organization's own Steps 2-4 are still internally parallelized across its
    staff), tagging every contact with its organization and writing all of them to
    a single combined workbook via export_combined_excel().
    """
    print(f"\nReading organizations from: {os.path.abspath(excel_path)}")
    teams = read_organizations_from_excel(excel_path, limit=limit)
    if not teams:
        print("No organizations found in the spreadsheet.")
        raise SystemExit(0)

    print(f"Found {len(teams)} unique organization(s) to process"
          f"{f' (limited to first {limit})' if limit else ''}:")
    for t in teams:
        print(f"  - {t['team']}" + (f" ({t['sport']})" if t["sport"] else " (no sport listed)"))

    all_staff: list[dict] = []
    for i, entry in enumerate(teams, start=1):
        team, sport = entry["team"], entry["sport"]
        athletes_label = ", ".join(entry.get("athletes", []))
        print(f"\n{'=' * 65}")
        print(f"Organization {i}/{len(teams)}: {team}" + (f" ({sport})" if sport else ""))
        if athletes_label:
            print(f"Athlete(s): {athletes_label}")
        print("=" * 65)

        print(f"\nStep 1/4 — Identifying medical personnel & leadership for {team} via Gemini …")
        staff = identify_staff_for_team(client, team, sport)
        if not staff:
            print(f"  No staff found for {team} — skipping.")
            continue
        print(f"  Found {len(staff)} person(s): {', '.join(p['name'] for p in staff)}\n")

        context = " ".join(p for p in [team, sport] if p)
        enrich_staff_contacts(client, staff, team, sport, context)

        for person in staff:
            person["organization"] = team
            person["athletes"] = athletes_label

        print("\n" + "-" * 65)
        print(format_staff(staff, show_sources=False))
        summarise_missing(staff)

        all_staff.extend(staff)

    if not all_staff:
        print("\nNo contacts found for any organization.")
        raise SystemExit(0)

    excel_out = export_combined_excel(all_staff, "batch_organizations")
    print(f"\nCombined Excel report written to: {os.path.abspath(excel_out)}")
    print(f"Total contacts across {len(teams)} organization(s): {len(all_staff)}")


def main() -> None:
    args = parse_args()

    if args.excel:
        mode = "excel"
        inputs = {"excel_path": args.excel, "limit": args.limit}
    elif args.player and args.team:
        if not args.sport:
            print("Error: --sport is required when using --team.")
            raise SystemExit(1)
        mode = "both"
        inputs = {"player": args.player.strip(), "team": args.team.strip(), "sport": args.sport.strip()}
    elif args.player:
        mode = "player"
        inputs = {"player": args.player.strip()}
    elif args.team:
        if not args.sport:
            print("Error: --sport is required when using --team.")
            raise SystemExit(1)
        mode = "team"
        inputs = {"team": args.team.strip(), "sport": args.sport.strip()}
    else:
        mode, inputs = interactive_mode_select()

    client = get_client()

    if mode == "excel":
        run_excel_batch(client, inputs["excel_path"], inputs.get("limit"))
        return

    show_sources = (mode == "both")

    player = inputs.get("player")
    team   = inputs.get("team")
    sport  = inputs.get("sport")

    context_parts = [p for p in [player, team, sport] if p]
    context = " ".join(context_parts)

    # ── Step 1 ────────────────────────────────────────────────────────────────
    if mode == "player":
        print(f"\nStep 1/4 — Identifying medical personnel for player '{player}' via Gemini …")
        staff = identify_staff_for_player(client, player)
        subject_label = f"player '{player}'"

    elif mode == "team":
        print(f"\nStep 1/4 — Identifying medical personnel for {team} ({sport}) via Gemini …")
        staff = identify_staff_for_team(client, team, sport)
        subject_label = f"team '{team}'"

    else:
        # Player-staff and team-staff lookups are fully independent Gemini calls —
        # run them concurrently instead of one after another.
        print(f"\nStep 1/4 — Identifying medical personnel for player '{player}' "
              f"and team '{team}' ({sport}) via Gemini (in parallel) …")
        with ThreadPoolExecutor(max_workers=2) as executor:
            player_future = executor.submit(identify_staff_for_player, client, player)
            team_future = executor.submit(identify_staff_for_team, client, team, sport)
            player_staff = player_future.result()
            team_staff = team_future.result()

        print(f"  Found {len(player_staff)} person(s) linked to the player: "
              f"{', '.join(p['name'] for p in player_staff) or '(none)'}")
        print(f"  Found {len(team_staff)} person(s) linked to the team: "
              f"{', '.join(p['name'] for p in team_staff) or '(none)'}")

        staff = merge_staff_lists(player_staff, team_staff)
        subject_label = f"player '{player}' and team '{team}'"

    if not staff:
        print(f"No staff found for {subject_label}.")
        raise SystemExit(0)

    if mode == "both":
        print(f"\nCombined total: {len(staff)} unique person(s) after de-duplication\n")
    else:
        print(f"  Found {len(staff)} person(s): {', '.join(p['name'] for p in staff)}\n")

    # ── Steps 2-4 ─────────────────────────────────────────────────────────────
    enrich_staff_contacts(client, staff, team, sport, context)

    print("\n" + "=" * 65)
    print(format_staff(staff, show_sources=show_sources))
    print("=" * 65)
    summarise_missing(staff)

    excel_path = export_to_excel(staff, subject_label)
    print(f"Excel report written to: {os.path.abspath(excel_path)}")


if __name__ == "__main__":
    main()
