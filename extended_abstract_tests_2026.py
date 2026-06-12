import openreview
import config as c
import pprint
import math
import re
import pandas as pd
from collections import defaultdict

# The 9 chairs' canonical OpenReview ids are curated here (same people).
from Lists.IWAI2026_Exclude import EXCLUDE_REVIEWERS

# DEFINITIONS
BASE_URL = 'https://api2.openreview.net'
YEAR='2026'
VENUE = f'IWAI-{YEAR}'
venue_id    = f'IWAI/{YEAR}/Workshop'
SUBMISSION_INVITATION = f'{venue_id}/-/Submission'
BL_SUBMISSION_INVITATION = f'{venue_id}/-/Blind_Submission'
ASSIGNMENT_INVITATION = f'{venue_id}/Reviewers/-/Assignment'
MSG_INVITATION=f'{venue_id}/-/Edit'
REVIEWER_GROUP = f'{venue_id}/Reviewers'
AUTHORS_GROUP = f'{venue_id}/Authors'
venue_id = f'IWAI/{YEAR}/Workshop'

# CLIENT (single login -- OpenReview rate-limits logins to 3 / 30s)
client = openreview.api.OpenReviewClient(baseurl=BASE_URL, username=c.usr, password=c.pas)
chair_group_id = f'{venue_id}/Program_Chairs'
chair_group = client.get_group(chair_group_id)
venue_group = client.get_group(venue_id)
submission_str   = venue_group.content['submission_name']['value'] # this query results in text "Submission"

# SUBMISSION TYPES
streams = ['1-comp','2-cogn','3-appl']
types=['1-paper','2-abstr']

submissions = client.get_all_notes(invitation=SUBMISSION_INVITATION, sort='number:asc', details='replies')

# ----- ABSTRACT REVIEW ASSIGNMENT CONFIG -----
# The 9 IWAI Chairs review the extended abstracts (type 2-abstr) themselves.
# Their expertise / CoI is curated in IWAI_reviewers.xlsx (sheet "Reviewers").
REVIEWERS_XLSX = "IWAI_chairs_2026.xlsx"
REVIEWERS_PER_ABSTRACT = 2
ABSTRACT_TYPE = 2  # types[1] == '2-abstr'
ASSIGNMENT_XLSX = f"{VENUE}_abstract_assignments.xlsx"

# ----- UPLOAD CONFIG -----
# Safety first: nothing is written to OpenReview while DRY_RUN is True.
# Flip to False only after a dry-run looks correct. Then optionally restrict to
# one or a few abstracts (by submission number) via LIMIT_TO_PAPERS to test that
# the assignment becomes visible in the Program Chairs console before deploying all.
DRY_RUN = True
LIMIT_TO_PAPERS = []          # e.g. [12] -> deploy only abstract #12; [] -> all
DELETE_EXISTING_EDGES = True  # remove prior Assignment edges on a paper before re-posting

# ----- NOTIFICATION CONFIG -----
# Reply-To shown on the assignment emails (OpenReview's native email uses the
# program chairs' contact address). DRY_RUN above also guards notify_reviewers().
NOTIFY_REPLYTO = "ivilinpeev.stoianov@cnr.it"
REVIEWERS_MESSAGE_INVITATION = f"{venue_id}/Reviewers/-/Message"


def resolve_profile(identifier):
    """Resolve an OpenReview profile by profile id or email."""
    try:
        return client.get_profile(identifier)
    except Exception:
        matches = client.search_profiles(term=identifier)
        if not matches:
            return None

        identifier_lc = identifier.strip().lower()
        for profile in matches:
            content = getattr(profile, 'content', {})
            preferred_email = str(content.get('preferredEmail', '')).strip().lower()
            emails = [str(email).strip().lower() for email in content.get('emails', [])]
            if identifier_lc == preferred_email or identifier_lc in emails:
                return profile

        return matches[0]


# ======================================================================
#  Extended-abstract reviewer assignment (the 9 IWAI Chairs)
# ======================================================================

def extract_field(content, key):
    """OpenReview V2 stores values as content[key]["value"] or content[key]."""
    if key not in content:
        return None
    value = content[key]
    if isinstance(value, dict) and "value" in value:
        return value["value"]
    return value


def _norm(text):
    """Lowercase + collapse whitespace."""
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text).strip().lower())


def _stream_index(cell):
    """Parse the leading digit (1/2/3) of a sub-category cell -> int, or None."""
    if cell is None or (isinstance(cell, float) and math.isnan(cell)):
        return None
    m = re.match(r"\s*(\d)", str(cell))
    return int(m.group(1)) if m else None


def _name_tokens(cell):
    """Surname-ish tokens from a free-text name/collaborator cell, lowercased.

    Splits on commas/semicolons/'and', then keeps alphabetic word tokens of
    length >= 3 (skips initials like 'M.'). Used for fuzzy CoI name matching.
    """
    if cell is None or (isinstance(cell, float) and math.isnan(cell)):
        return set()
    tokens = set()
    for chunk in re.split(r"[;,]| and ", str(cell)):
        for word in re.findall(r"[A-Za-z]+", chunk):
            if len(word) >= 3:
                tokens.add(word.lower())
    return tokens


def load_reviewers(path=REVIEWERS_XLSX):
    """Load the 9 chairs from IWAI_reviewers.xlsx -> list of reviewer dicts."""
    df = pd.read_excel(path, sheet_name="Reviewers")
    reviewers = []
    for _, row in df.iterrows():
        name = str(row.get("Name", "")).strip()
        if not name or name.lower() == "nan":
            continue
        emails = [e.strip() for e in re.split(r"[;,]", str(row.get("Email(s)", ""))) if e.strip()]
        keywords = {
            _norm(k) for k in re.split(r"[;,]", str(row.get("Keywords / expertise", "")))
            if _norm(k)
        }
        coi_names = (
            _name_tokens(row.get("Key collaborators (COI input)"))
            | _name_tokens(row.get("Do-not-review (COI)"))
        )
        reviewers.append({
            "name": name,
            "emails": emails,
            "emails_lc": {e.lower() for e in emails},
            "primary_stream": _stream_index(row.get("Primary sub-category")),
            "secondary_stream": _stream_index(row.get("Secondary sub-category")),
            "keywords": keywords,
            "coi_names": coi_names,
            "or_ids": set(),       # all OpenReview profile ids, filled by map_reviewers_to_or
            "rev_id": None,        # preferred ~Tilde_Id1 used for the Assignment edge
            "load": 0,
        })
    print(f"Loaded {len(reviewers)} chair-reviewers from {path}.")
    return reviewers


def get_abstracts():
    """Collect all type 2-abstr submissions as plain dicts."""
    abstracts = []
    for note in submissions:
        ptype = extract_field(note.content, "type")
        if not ptype or int(str(ptype)[0]) != ABSTRACT_TYPE:
            continue
        stream = extract_field(note.content, "stream")
        kwds = extract_field(note.content, "keywords") or []
        abstracts.append({
            "paper_id": note.id,
            "number": note.number,
            "stream": int(str(stream)[0]) if stream else None,
            "stream_name": streams[int(str(stream)[0]) - 1] if stream else "",
            "title": extract_field(note.content, "title") or "",
            "keywords": {_norm(k) for k in kwds if _norm(k)},
            "abstract": extract_field(note.content, "abstract") or "",
            "authors": extract_field(note.content, "authors") or [],
            "authorids": extract_field(note.content, "authorids") or [],
        })
    print(f"Found {len(abstracts)} extended abstracts (2-abstr).")
    return abstracts


def _name_key(text):
    """Normalize a person name for matching: lowercase, drop punctuation/initials.

    Strips parenthetical aliases and periods so 'Christopher L. Buckley' matches
    the profile's 'Christopher L Buckley'.
    """
    text = re.sub(r"\(.*?\)", "", text or "")
    text = re.sub(r"[^\w\s]", " ", text)  # drop '.', ',', etc.
    return _norm(text)


def _profile_names(profile):
    """All name keys a profile exposes (fullname + first/last), punctuation-free."""
    out = set()
    for n in profile.content.get("names", []) or []:
        full = n.get("fullname")
        if full:
            out.add(_name_key(full))
        fl = " ".join(p for p in [n.get("first"), n.get("last")] if p)
        if fl:
            out.add(_name_key(fl))
    return out


def map_reviewers_to_or(reviewers):
    """Resolve each chair's canonical OpenReview ~id and conflict ids.

    Email lookup is unreliable here (chair emails are masked/unconfirmed and the
    confirmedEmails search endpoint is broken in this client version), so we
    resolve against the curated tilde-ids in EXCLUDE_REVIEWERS -- these ARE the 9
    chairs. Each curated id is fetched with get_profile (which works), and matched
    to an Excel row by normalized profile name. Email overlap is used as a
    tie-break / fallback. Sets r["rev_id"] (the Assignment edge 'tail') and
    r["or_ids"] (profile id + emails, used for conflict detection).
    """
    # Build a lookup of curated profiles by name and by email.
    by_name, profile_by_id = {}, {}
    for tid in EXCLUDE_REVIEWERS:
        try:
            profile = client.get_profile(tid)
        except Exception as ex:
            print(f"  Could not fetch {tid}: {ex}")
            continue
        profile_by_id[tid] = profile
        for nm in _profile_names(profile):
            by_name[nm] = tid

    for r in reviewers:
        match = None
        # 1. Name match (Excel "Name" vs profile names; punctuation-insensitive).
        match = by_name.get(_name_key(r["name"]))
        # 2. Fallback: email overlap with a curated profile's emails.
        if match is None:
            for tid, profile in profile_by_id.items():
                pemails = {str(e).lower() for e in profile.content.get("emails", []) or []}
                # masked emails ("****@verses.ai") still share the domain tail
                if r["emails_lc"] & pemails:
                    match = tid
                    break
        if match is None:
            print(f"  WARNING: no OpenReview ~id resolved for {r['name']} "
                  f"-- cannot be uploaded.")
            continue
        r["rev_id"] = match
        r["or_ids"].add(match)
        r["or_ids"].update(profile_by_id[match].content.get("emails", []) or [])


def list_chair_or_ids(path=REVIEWERS_XLSX):
    """Print each chair's resolved OpenReview ~id (answers 'how to extract them').

    Also cross-checks membership in the venue Reviewers group, since a chair
    must be in {venue_id}/Reviewers for an Assignment edge to be valid.
    """
    reviewers = load_reviewers(path)
    map_reviewers_to_or(reviewers)
    try:
        members = set(client.get_group(REVIEWER_GROUP).members)
    except Exception:
        members = set()
    print(f"\nChair OpenReview ids (Reviewers group has {len(members)} members):")
    for r in reviewers:
        in_group = "in Reviewers" if r["rev_id"] in members else "NOT in Reviewers group"
        print(f"  {r['rev_id'] or '??? UNRESOLVED'}  <-  {r['name']}  [{in_group}]")
    return reviewers


def ensure_chairs_in_reviewers_group(path=REVIEWERS_XLSX):
    """Add any chair missing from the venue Reviewers group.

    An Assignment edge is only valid if the reviewer is a member of
    {venue_id}/Reviewers, so run this BEFORE upload_abstract_assignments().
    Guarded by DRY_RUN: prints who would be added and changes nothing.
    """
    reviewers = load_reviewers(path)
    map_reviewers_to_or(reviewers)
    members = set(client.get_group(REVIEWER_GROUP).members)

    to_add = sorted({
        r["rev_id"] for r in reviewers
        if r["rev_id"] and r["rev_id"] not in members
    })
    if not to_add:
        print("All chairs already in the Reviewers group.")
        return

    mode = "DRY-RUN (no change)" if DRY_RUN else "ADDING"
    print(f"{mode}: {len(to_add)} chair(s) missing from {REVIEWER_GROUP}: {to_add}")
    if DRY_RUN:
        return
    client.add_members_to_group(group=REVIEWER_GROUP, members=to_add)
    print(f"Added {len(to_add)} chair(s) to the Reviewers group.")


def get_CoI_from_OpenReview():
    """paper_id -> {conflicted reviewer ids} from OpenReview's computed edges."""
    paper_conflicts = {}
    conflict_invitation = f"{venue_id}/Reviewers/-/Conflict"
    try:
        grouped_edges = client.get_grouped_edges(invitation=conflict_invitation, groupby='head')
    except Exception as e:
        print(f"  (OpenReview CoI edges unavailable: {e})")
        return paper_conflicts
    for group in grouped_edges:
        paper_id = group['id']['head']
        tails = {edge['tail'] if isinstance(edge, dict) else edge.tail for edge in group['values']}
        paper_conflicts[paper_id] = tails
    print(f"  Loaded OpenReview CoI edges for {len(paper_conflicts)} papers.")
    return paper_conflicts


def has_conflict(abstract, reviewer, paper_conflicts):
    """True if the chair must not review this abstract."""
    # 1. Self-authorship (by OpenReview id or by email present in authorids).
    author_ids_lc = {str(a).lower() for a in abstract["authorids"]}
    if reviewer["or_ids"] & set(abstract["authorids"]):
        return True
    if reviewer["emails_lc"] & author_ids_lc:
        return True

    # 2. Excel collaborator / do-not-review surnames vs abstract author surnames.
    author_tokens = set()
    for a in abstract["authors"]:
        author_tokens |= _name_tokens(a)
    if reviewer["coi_names"] & author_tokens:
        return True

    # 3. OpenReview computed conflict edges.
    conflicted = paper_conflicts.get(abstract["paper_id"], set())
    if reviewer["or_ids"] & conflicted:
        return True

    return False


def score(abstract, reviewer):
    """Interpretable expertise score: stream match + keyword overlap (no ML)."""
    s = 0.0
    if abstract["stream"] is not None:
        if abstract["stream"] == reviewer["primary_stream"]:
            s += 1.0
        elif abstract["stream"] == reviewer["secondary_stream"]:
            s += 0.5

    # Exact keyword-set overlap.
    s += 0.1 * len(abstract["keywords"] & reviewer["keywords"])

    # Soft substring match of reviewer keywords against title + abstract text.
    haystack = _norm(abstract["title"]) + " " + _norm(abstract["abstract"])
    soft = sum(1 for kw in reviewer["keywords"] if kw and kw in haystack)
    s += 0.05 * min(soft, 6)  # capped to avoid keyword-stuffing dominance
    return s


def assign(abstracts, reviewers, paper_conflicts):
    """Greedy, load-balanced assignment of 2 chairs per abstract."""
    n_slots = REVIEWERS_PER_ABSTRACT * len(abstracts)
    base_cap = math.ceil(n_slots / len(reviewers)) if reviewers else 0

    selected = defaultdict(list)   # paper_id -> [reviewer dict, ...]

    def fill_pass(cap):
        for abs in abstracts:
            while len(selected[abs["paper_id"]]) < REVIEWERS_PER_ABSTRACT:
                chosen = selected[abs["paper_id"]]
                eligible = [
                    r for r in reviewers
                    if r not in chosen
                    and r["load"] < cap
                    and not has_conflict(abs, r, paper_conflicts)
                ]
                if not eligible:
                    break
                # Highest score, then lowest current load.
                eligible.sort(key=lambda r: (-score(abs, r), r["load"]))
                pick = eligible[0]
                pick["load"] += 1
                chosen.append(pick)

    fill_pass(base_cap)          # fair pass
    fill_pass(base_cap + 1)      # relaxed fallback for any still-short abstracts

    # Build the output table (analogous to reviewer_selector.assignment3).
    rows = []
    unfilled = []
    for abs in abstracts:
        chosen = selected[abs["paper_id"]]
        if len(chosen) < REVIEWERS_PER_ABSTRACT:
            unfilled.append(abs["number"])
        for i, r in enumerate(chosen):
            if abs["stream"] == r["primary_stream"]:
                role = "primary"
            elif abs["stream"] == r["secondary_stream"]:
                role = "secondary"
            else:
                role = "fallback"
            rows.append({
                "pap_stream": abs["stream_name"],
                "pap_number": abs["number"],
                "pap_id": abs["paper_id"],
                "pap_title": abs["title"],
                "pap_authors": "; ".join(abs["authors"]),
                "rev_id": r["rev_id"],   # ~Tilde_Id1 -> Assignment edge 'tail'
                "reviewer_name": r["name"],
                "reviewer_email": r["emails"][0] if r["emails"] else "",
                "score": round(score(abs, r), 3),
                "rev_role": role,
                "rev_load": r["load"],
                "rev_primary_stream": r["primary_stream"],
                "rev_secondary_stream": r["secondary_stream"],
            })

    df = pd.DataFrame(rows).sort_values(by=["pap_stream", "pap_number", "rev_role"])
    df.to_excel(ASSIGNMENT_XLSX, index=False)
    print(f"\nSaved {len(df)} assignment rows to {ASSIGNMENT_XLSX}")

    loads = [r["load"] for r in reviewers]
    print("\nPer-chair load:")
    for r in sorted(reviewers, key=lambda r: -r["load"]):
        print(f"  {r['load']:>2}  {r['name']}")
    if loads:
        print(f"Load min/mean/max: {min(loads)} / {sum(loads) / len(loads):.2f} / {max(loads)}")
    print(f"Expected rows (2 x {len(abstracts)} abstracts): {REVIEWERS_PER_ABSTRACT * len(abstracts)}")
    if unfilled:
        print(f"WARNING: abstracts with < {REVIEWERS_PER_ABSTRACT} reviewers after fallback: {unfilled}")
    return df


def assign_abstract_reviewers():
    reviewers = load_reviewers()
    abstracts = get_abstracts()
    print("\nResolving chair OpenReview profiles...")
    map_reviewers_to_or(reviewers)
    print("Loading conflicts from OpenReview...")
    paper_conflicts = get_CoI_from_OpenReview()
    return assign(abstracts, reviewers, paper_conflicts)


def upload_abstract_assignments():
    """Deploy the abstract assignments to OpenReview as Assignment edges.

    Reads ASSIGNMENT_XLSX, groups by pap_id, and for each abstract writes an
    Assignment edge (head=paper note id, tail=reviewer ~id) plus per-submission
    Reviewers-group membership -- the same edge-based model the full-paper
    upload_assignments() uses in review_assignments.py.

    Guarded by DRY_RUN (default True): prints what it would post and writes
    nothing. Set DRY_RUN=False to deploy; optionally restrict to a few abstracts
    via LIMIT_TO_PAPERS to test visibility in the Program Chairs console first.
    """
    df = pd.read_excel(ASSIGNMENT_XLSX)

    if LIMIT_TO_PAPERS:
        df = df[df["pap_number"].isin(LIMIT_TO_PAPERS)]
        print(f"Restricted to abstracts {LIMIT_TO_PAPERS}: {len(df)} rows.")

    missing = df[df["rev_id"].isna()]
    if len(missing):
        print(f"ERROR: {len(missing)} rows have no rev_id (unresolved chairs); "
              f"fix these before uploading:\n{missing[['pap_number','reviewer_name']]}")
        return

    mode = "DRY-RUN (nothing will be written)" if DRY_RUN else "LIVE UPLOAD"
    print(f"\n=== {mode} : {len(df)} edges over {df['pap_id'].nunique()} abstracts ===")

    for paper_id, group in df.groupby("pap_id"):
        paper_number = int(group["pap_number"].iloc[0])
        paper_reviewers_group_id = f"{venue_id}/Submission{paper_number}/Reviewers"
        rev_ids = list(group["rev_id"])

        print(f"\nAbstract #{paper_number} ({paper_id}) -> {rev_ids}")

        if DRY_RUN:
            continue

        # Replace any prior assignment edges for this abstract.
        if DELETE_EXISTING_EDGES:
            existing = client.get_edges(invitation=ASSIGNMENT_INVITATION, head=paper_id)
            if existing:
                client.delete_edges(invitation=ASSIGNMENT_INVITATION, head=paper_id,
                                    soft_delete=False)

        edges = [
            openreview.api.Edge(
                invitation=ASSIGNMENT_INVITATION,
                head=paper_id, tail=rev_id, weight=1,
                readers=[venue_id, paper_reviewers_group_id, rev_id],
                writers=[venue_id],
                signatures=[venue_id],
            )
            for rev_id in rev_ids
        ]
        client.add_members_to_group(group=paper_reviewers_group_id, members=rev_ids)
        openreview.tools.post_bulk_edges(client=client, edges=edges)

    if DRY_RUN:
        print("\nDRY-RUN complete. Review the above, then set DRY_RUN=False "
              "(optionally LIMIT_TO_PAPERS=[<one number>]) and re-run.")
    else:
        print(f"\nUploaded {len(df)} assignments over {df['pap_id'].nunique()} abstracts.")


NOTIFY_SUBJECT = f"[{VENUE}] You have been assigned a paper to review"

NOTIFY_BODY = (
    "This is to inform you that you have been assigned as a Reviewer for paper "
    "number {paper_number} for {VENUE_PLAIN}.\n\n"
    'Title: "{paper_title}"\n\n'
    "To review this new assignment, please login to OpenReview and go to "
    "https://openreview.net/forum?id={forum_id}.\n\n"
    "To check all of your assigned papers, go to "
    "https://openreview.net/group?id={reviewer_group}.\n\n"
    "Thank you,\n\n"
    "{VENUE_PLAIN} Workshop Program Chairs"
)


def notify_reviewers():
    """Email each chair their assigned abstracts, mirroring OpenReview's native
    assignment notification (one message per paper-reviewer pair).

    Uses the Reviewers Message invitation and sets replyTo to the program chairs'
    contact (NOTIFY_REPLYTO) -- the same shape OpenReview's manual-assignment email
    has. Guarded by DRY_RUN: prints the messages it would send and sends nothing.
    """
    df = pd.read_excel(ASSIGNMENT_XLSX)
    if LIMIT_TO_PAPERS:
        df = df[df["pap_number"].isin(LIMIT_TO_PAPERS)]
        print(f"Restricted to abstracts {LIMIT_TO_PAPERS}: {len(df)} rows.")

    missing = df[df["rev_id"].isna()]
    if len(missing):
        print(f"ERROR: {len(missing)} rows have no rev_id; cannot notify:\n"
              f"{missing[['pap_number','reviewer_name']]}")
        return

    venue_plain = VENUE.replace("-", " ")  # "IWAI 2026"
    mode = "DRY-RUN (no email sent)" if DRY_RUN else "SENDING EMAILS"
    print(f"\n=== {mode} : {len(df)} notifications (one per paper-reviewer pair) ===")

    sent = 0
    for _, row in df.iterrows():
        paper_number = int(row["pap_number"])
        forum_id = row["pap_id"]
        paper_title = row["pap_title"]
        rev_id = row["rev_id"]
        body = NOTIFY_BODY.format(
            paper_number=paper_number, paper_title=paper_title, forum_id=forum_id,
            reviewer_group=REVIEWER_GROUP, VENUE_PLAIN=venue_plain,
        )
        print(f"  -> {rev_id}  (abstract #{paper_number})")
        if DRY_RUN:
            continue
        try:
            client.post_message(
                invitation=REVIEWERS_MESSAGE_INVITATION,
                recipients=[rev_id],
                subject=NOTIFY_SUBJECT,
                message=body,
                replyTo=NOTIFY_REPLYTO,
                signature=venue_id,
            )
            sent += 1
        except Exception as ex:
            print(f"     FAILED for {rev_id}: {ex}")

    if DRY_RUN:
        print("\nDRY-RUN complete. Set DRY_RUN=False (optionally LIMIT_TO_PAPERS=[..]) "
              "to actually send.")
    else:
        print(f"\nSent {sent} assignment notifications.")


if __name__ == '__main__':
    # 1. Compute and export the assignment table.
    assign_abstract_reviewers()
    # 2. (helper) Print each chair's resolved OpenReview ~id + group membership.
    list_chair_or_ids()
    # 3. Ensure all chairs are in the venue Reviewers group (required for valid
    #    assignment edges). Run BEFORE upload. DRY_RUN-guarded.
    ensure_chairs_in_reviewers_group()
    # 4. Deploy to OpenReview (DRY_RUN=True by default -- prints only).
    upload_abstract_assignments()
    # 5. Email reviewers their assignments (DRY_RUN=True by default -- prints only).
    #    Run this AFTER upload, so reviewers can actually see the papers.
    notify_reviewers()
