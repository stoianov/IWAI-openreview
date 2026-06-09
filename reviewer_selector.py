import openreview
import pandas as pd
import numpy as np
from tqdm import tqdm
from collections import defaultdict
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import itertools
import re
import pprint

import config_ivo as c
from collections import defaultdict
import time

# SUBMISSION TYPES
streams = ['1-comp','2-cogn','3-appl']
types=['1-paper','2-abstr']
TECHNICAL_STREAM = streams[0] # Computational

BASEURL = 'https://api2.openreview.net'
YEAR = '2026'
VENUE = f'IWAI-{YEAR}'
venue_id = f'IWAI/{YEAR}/Workshop'
SUBMISSION_INVITATION = f'{venue_id}/-/Submission'

TOP_K_CANDIDATES = 10
DEFAULT_REVIEW_LOAD = 2
SENIOR_REVIEW_LOAD = 2
TECHNICAL_REVIEWERS_PER_PAPER = 1
STREAM_REVIEWERS_PER_PAPER = 2
MODEL_NAME = "all-MiniLM-L6-v2"
RANDOM_SEED = 42
np.random.seed(RANDOM_SEED)


USE_INSTITUTION_CONFLICTS = True
MAX_COAUTHOR_YEARS = 5
SLEEP_BETWEEN_PROFILE_CALLS = 0.05

client=openreview.api.OpenReviewClient(baseurl=BASEURL,username=c.usr,password=c.pas)
submissions = client.get_all_notes(invitation=SUBMISSION_INVITATION)
print(f"Downloaded {len(submissions)} submissions.")


# Empty CoI database
papers, reviewers, reviewers_id = [], {}, []
profile_cache, reviewer_coauthors, reviewer_institutions, paper_author_institutions = {},{},{}, {}

# HELPERS
def normalize_text(x):
    if x is None:
        return ""
    return re.sub(r"\s+", " ", x.strip().lower())

def extract_field(content, key):
    """  OpenReview V2 stores values as:  content[key]["value"] or content[key]  """
    if key not in content:  return None
    value = content[key]
    if isinstance(value, dict) and "value" in value:
        return value["value"]
    return value

def get_submissions():
    global papers
    papers = []
    for note in submissions:
        title = extract_field(note.content, "title")
        abstract = extract_field(note.content, "abstract")
        authors = extract_field(note.content, "authors")
        authorids = extract_field(note.content, "authorids")
        kwds = extract_field(note.content, "keywords")
        stream = extract_field(note.content, "stream")
        Stream = streams[int(stream[0])-1] if stream else ""
        ptype = extract_field(note.content, "type")
        Ptype = types[int(ptype[0])-1]
        paper = {
            "paper_id": note.id,
            "number": note.number,
            "title": title,
            "keywords": kwds,
            "abstract": abstract,
            "authors": authors,
            "authorids": authorids,
            "stream": Stream,
            "paper_type": Ptype,
        }
        papers.append(paper)

    papers_df = pd.DataFrame(papers)
    print(f"\nStreams:{papers_df['stream'].value_counts()}")
    print(f"\nTypes: {papers_df['paper_type'].value_counts()}")

def extract_reviewers():
    global reviewers, reviewers_id
    print("\nBuilding reviewer pool...")
    reviewers = {}
    for paper in papers:
        authors = paper["authors"]
        authorids = paper["authorids"]
        if not authors:  continue
        positions = []
        if len(authors) >= 1: positions.append((authorids[0],authors[0], "first"))
        if len(authors) >= 2: positions.append((authorids[1],authors[1], "second"))
        if len(authors) >= 3: positions.append((authorids[-1],authors[-1], "senior"))

        for authorid, author, role in positions:
            if author not in reviewers:
                max_load = SENIOR_REVIEW_LOAD if role == "senior" else DEFAULT_REVIEW_LOAD
                reviewers[authorid] = {
                    "reviewer_id": authorid,
                    "reviewer": author,
                    "roles": set([role]),
                    "streams": set([paper["stream"]]),
                    "paper_types": set([paper["paper_type"]]),
                    "max_load": max_load,
                    "current_load": 0,
                    "papers_authored": set([paper["paper_id"]]),
                    "expertise_texts": [],
                }
            else:
                reviewers[authorid]["roles"].add(role)
                reviewers[authorid]["streams"].add(paper["stream"])
                reviewers[authorid]["paper_types"].add(paper["paper_type"])
                reviewers[authorid]["papers_authored"].add(paper["paper_id"])

            kwd = " ".join(paper["keywords"])
            expertise_text = (normalize_text(paper["title"])+" "+normalize_text(paper["abstract"])+" "+normalize_text(kwd))
            reviewers[authorid]["expertise_texts"].append(expertise_text)

    reviewers_id = []
    for rid, data in reviewers.items(): reviewers_id.append(rid)

    print(f"Total reviewers: {len(reviewers)}")

def paper_to_reviewer_match():
    # REVIEWER EXPERTISE vs PAPER TEXT
    reviewer_texts = []
    for rid, data in reviewers.items():
        combined = " ".join(data["expertise_texts"])
        reviewer_texts.append(combined)
    paper_texts = [
        normalize_text(p["title"]) + " " + normalize_text(p["abstract"]) + " " + normalize_text(" ".join(p["keywords"]))
        for p in papers  ]
    print("\nLoading embedding model...")
    model = SentenceTransformer(MODEL_NAME)
    print("Encoding papers...")
    paper_embeddings = model.encode(paper_texts, show_progress_bar=True)
    print("Encoding reviewers...")
    reviewer_embeddings = model.encode(reviewer_texts, show_progress_bar=True)
    similarity_matrix = cosine_similarity(paper_embeddings, reviewer_embeddings)
    return similarity_matrix

def has_conflict0(paper, reviewer_id):
    if reviewer_id in paper["authors"]: return True   # self-review
    return False

def has_conflict(paper, reviewer_id):
    if reviewer_id in paper["authors"]: return True  # self-review
    reviewer_cos = reviewer_coauthors.get( reviewer_id, set())
    for author in paper["authors"]:
        if author in reviewer_cos: return True
    if USE_INSTITUTION_CONFLICTS:
        reviewer_inst = reviewer_institutions.get( reviewer_id, set())
        author_inst = paper_author_institutions.get( paper["paper_id"], set())
        overlap = reviewer_inst.intersection(author_inst)
        if len(overlap) > 0: return True
    return False

def reviewer_eligible(paper, reviewer_id):
    reviewer = reviewers[reviewer_id]
    if reviewer["current_load"] >= reviewer["max_load"]:  return False
    if has_conflict(paper, reviewer_id): return False
    if paper["stream"] not in reviewer["streams"]: return False
    if paper["paper_type"] not in reviewer["paper_types"]: return False
    return True

def reviewer_candidates(similarity_matrix):
    print("\nGenerating candidate reviewers...")
    candidate_rows = []
    paper_to_candidates = {}

    for paper_idx, paper in tqdm(list(enumerate(papers))):
        scores = []
        for reviewer_idx, reviewer_id in enumerate(reviewers_id):
            if not reviewer_eligible(paper, reviewer_id):  continue
            sim = similarity_matrix[paper_idx, reviewer_idx]
            reviewer = reviewers[reviewer_id]
            load_penalty = reviewer["current_load"] / reviewer["max_load"]
            score = sim - 0.15 * load_penalty
            scores.append( ( reviewer_id, score, sim ) )
        scores = sorted(scores, key=lambda x: x[1], reverse=True)
        top_candidates = scores[:TOP_K_CANDIDATES]
        paper_to_candidates[paper["paper_id"]] = top_candidates
        row = {"paper_number": paper["number"], "paper_title": paper["title"]}

        for i, (rid, score, sim) in enumerate(top_candidates):
            row[f"candidate_{i+1}"] = rid
            row[f"score_{i+1}"] = round(float(score), 4)
        candidate_rows.append(row)

    candidate_df = pd.DataFrame(candidate_rows)
    fname1=f"{VENUE}_candidate_reviewers.xlsx"
    candidate_df.to_excel(fname1 , index=False)
    print(f"\nSaved to {fname1}")
    return paper_to_candidates

def assignment(paper_to_candidates):

    print("\nComputing final assignments...")
    TARGET_REVIEWS = (TECHNICAL_REVIEWERS_PER_PAPER + STREAM_REVIEWERS_PER_PAPER)
    assignments = []
    paper_assignment_count = defaultdict(int)
    reviewer_assignment_titles = defaultdict(list)

    for paper in papers:
        selected = []
        candidates = paper_to_candidates[paper["paper_id"]]
        # sort dynamically by current load
        candidates = sorted( candidates, key=lambda x: ( reviewers[x[0]]["current_load"], -x[1] ) )

        for rid, score, sim in candidates:
            reviewer = reviewers[rid]
            if reviewer["current_load"] >= reviewer["max_load"]:  continue
            if rid in [x["reviewer"] for x in selected]:          continue # avoid duplicate reviewer
            role = "stream"
            if len(selected) == 0: role = "technical"
            # Add reviewer
            selected.append({ "reviewer": rid, "score": score, "role": role})
            reviewer["current_load"] += 1
            reviewer_assignment_titles[rid].append( paper["title"])
            paper_assignment_count[paper["paper_id"]] += 1
            if len(selected) >=TARGET_REVIEWS: break

        # Fallback relaxed criteria
        if len(selected) < TARGET_REVIEWS:
            for reviewer_id in reviewers_id:
                reviewer = reviewers[reviewer_id]
                if reviewer["current_load"] >= reviewer["max_load"] + 1:  continue
                if has_conflict(paper, reviewer_id):                      continue
                if paper["paper_type"] not in reviewer["paper_types"]:    continue
                if reviewer_id in [x["reviewer"] for x in selected]:      continue # avoid duplicates
                # Add reviewer
                selected.append({"reviewer": reviewer_id, "score": -1,"role": "fallback", })
                reviewer["current_load"] += 1
                reviewer_assignment_titles[reviewer_id].append(paper["title"])
                paper_assignment_count[paper["paper_id"]] += 1
                if len(selected) >= TARGET_REVIEWS: break

        # Export reviewers x paper
        for item in selected:
            rid = item["reviewer"]
            reviewer_submissions = [ p["title"] for p in papers if rid in p["authors"] ]
            assignments.append({
                "paper_number":     paper["number"],
                "paper_title":      paper["title"],
                "paper_authors":    "; ".join(paper["authors"]),
                "reviewer":         rid,
                "reviewer_role":    item["role"],
                "score":            round(float(item["score"]), 2),
                "reviewer_tasks":   reviewers[rid]["current_load"],
                "reviewer_nsubm":   len(reviewers[rid]["papers_authored"]),
                "reviewer_submis":  " | ".join(reviewer_submissions),
                "stream":           paper["stream"],
                "paper_type":       paper["paper_type"],
            })

    assignments_df = pd.DataFrame(assignments)
    fname2=f"{VENUE}_final_assignment.xlsx"
    assignments_df.to_excel(fname2, index=False)
    print(f"\nSaved to {fname2}")

    loads = []
    for rid, data in reviewers.items(): loads.append(data["current_load"])
    loads = np.array(loads)
    print(f"Reviewer Load: Min/Mean/Max load: {loads.min()} / {loads.mean():.2f} / {loads.max()}")


def assignment2(paper_to_candidates):
    print("\nComputing final assignments ...")
    assignments = []
    for paper_idx, paper in enumerate(papers):
        selected_reviewers = set()
        selected_items = []
        candidates = paper_to_candidates[paper["paper_id"]]

        # Computational reviewer
        technical_candidates = []
        for rid, score, sim in candidates:
            reviewer = reviewers[rid]
            if reviewer["current_load"] >= reviewer["max_load"]:  continue
            if rid in selected_reviewers:  continue
            if TECHNICAL_STREAM not in reviewer["streams"]:  continue
            technical_candidates.append((rid, score, sim))
        technical_candidates = sorted( technical_candidates, key=lambda x: (reviewers[x[0]]["current_load"], -x[1]))
        if len(technical_candidates) > 0:
            rid, score, sim = technical_candidates[0]
            reviewers[rid]["current_load"] += 1
            selected_reviewers.add(rid)
            selected_items.append({ "reviewer": rid, "role": "technical", "score": score, })

        # Same-stream reviewer
        stream_candidates = []
        for rid, score, sim in candidates:
            reviewer = reviewers[rid]
            if reviewer["current_load"] >= reviewer["max_load"]:  continue
            if rid in selected_reviewers: continue
            if paper["stream"] not in reviewer["streams"]: continue
            stream_candidates.append((rid, score, sim))
        stream_candidates = sorted( stream_candidates, key=lambda x: ( reviewers[x[0]]["current_load"], -x[1] ))
        for rid, score, sim in stream_candidates:
            reviewers[rid]["current_load"] += 1
            selected_reviewers.add(rid)
            selected_items.append({ "reviewer": rid, "role": "stream", "score": score, })
            if len([ x for x in selected_items  if x["role"] == "same_stream"]) >= STREAM_REVIEWERS_PER_PAPER: break

        # Export reviewers x paper
        for item in selected_items:
            rid = item["reviewer"]
            reviewer_submissions = [ p["title"]  for p in papers if rid in p["authorids"] ]
            assignments.append({
                "paper_number":     paper["number"],
                "paper_title":      paper["title"],
                "paper_authors":    "; ".join(paper["authors"]),
                "reviewer":         rid,
                "reviewer_role":    item["role"],
                "score":            round(float(item["score"]), 2),
                "reviewer_tasks":   reviewers[rid]["current_load"],
                "reviewer_nsubm":   len(reviewers[rid]["papers_authored"]),
                "reviewer_submis":  " | ".join(reviewer_submissions),
                "stream":           paper["stream"],
                "paper_type":       paper["paper_type"],
            })

    assignments_df = pd.DataFrame(assignments)
    fname2=f"{VENUE}_final_assignment.xlsx"
    assignments_df.to_excel(fname2, index=False)
    print(f"\nSaved to {fname2}")

    loads = []
    for rid, data in reviewers.items(): loads.append(data["current_load"])
    loads = np.array(loads)
    print(f"Reviewer Load: Min/Mean/Max load: {loads.min()} / {loads.mean():.2f} / {loads.max()}")



def get_profiles():
    global profile_cache
    print("\nDownloading reviewer profiles...")
    profile_cache = {}
    for rid in tqdm(reviewers_id):
        try:
            profile = client.get_profile(rid)
            profile_cache[rid] = profile
        except Exception as e:
            print(f"Could not retrieve profile for {rid}")
            profile_cache[rid] = None
        time.sleep(SLEEP_BETWEEN_PROFILE_CALLS)

def extract_institutions(profile):
    institutions = set()
    if profile is None:  return institutions
    content = profile.content
    history = content.get("history", [])
    for item in history:
        institution = item.get("institution", {})
        if isinstance(institution, dict):
            name = institution.get("name")
            if name: institutions.add(normalize_text(name))
    return institutions

def extract_coauthors(profile_id):
    coauthors = set()
    try:
        notes = client.get_all_notes( content={"authorids": profile_id})
    except Exception as e:
        print(f"Could not retrieve notes for {profile_id}")
        return coauthors
    for note in notes:
        content = note.content
        authorids = extract_field( content,"authorids")
        if not authorids: continue
        for aid in authorids:
            if aid != profile_id: coauthors.add(aid)
    return coauthors

def extract_coauthors0(profile):
    coauthors = set()
    if profile is None: return coauthors
    content = profile.content
    publications = content.get("publications", [])
    current_year = YEAR
    for pub in publications:
        year = pub.get("year")
        if year is not None:
            try:
                year = int(year)
                if current_year - year > MAX_COAUTHOR_YEARS: continue
            except: pass
        authors = pub.get("authors", [])
        for a in authors:
            if isinstance(a, dict):
                author_id = a.get("id")
                if author_id: coauthors.add(author_id)
            elif isinstance(a, str):
                coauthors.add(a)
    return coauthors

def build_CoI_database():
    global reviewer_coauthors, reviewer_institutions, paper_author_institutions
    print("\nBuilding CoI databases...")
    reviewer_coauthors, reviewer_institutions, paper_author_institutions  = {}, {},{}
    for rid in reviewers_id:
        profile = profile_cache[rid]
        #reviewer_coauthors[rid] = extract_coauthors(profile)
        reviewer_coauthors[rid] = extract_coauthors(rid)
        reviewer_institutions[rid] = extract_institutions(profile)
    for paper in papers:
        institutions = set()
        for author in paper["authors"]:
            if author in reviewer_institutions:
                institutions.update(reviewer_institutions[author])
        paper_author_institutions[paper["paper_id"]] = institutions


if __name__ == '__main__':
    get_submissions()
    extract_reviewers()
    get_profiles()
    build_CoI_database()
    similarity = paper_to_reviewer_match()
    paper_to_candidates = reviewer_candidates(similarity)
    assignment(paper_to_candidates)
    print("\nDone.")


