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
from textwrap import dedent

import config_ivo as c
from collections import defaultdict
import time
from openpyxl import Workbook


from utils import extract_field, normalize_text, normalize_dict

from Lists.IWAI2026_Exclude import EXCLUDE_REVIEWERS
from Lists.IWAI2026_Exclude import EXCLUDE_PAPERS

# SUBMISSION TYPES
streams = ['1-comp','2-cogn','3-appl']
types=['1-paper','2-abstr']
TECHNICAL_STREAM = streams[0] # Computational
only_papers=True

BASEURL = 'https://api2.openreview.net'
YEAR = '2026'
VENUE = f'IWAI-{YEAR}'
VENUE2 = f'IWAI {YEAR}'
venue_id = f'IWAI/{YEAR}/Workshop'
SUBMISSION_INVITATION = f'{venue_id}/-/Submission'
ASSIGNMENT_INVITATION = f'{venue_id}/Reviewers/-/Assignment'
COI_INVITATION = f"{venue_id}/Reviewers/-/Conflict"
MSG_INVITATION=f'{venue_id}/-/Edit'

REVIEWER_GROUP = f'{venue_id}/Reviewers'


TOP_K_CANDIDATES = 35
DEFAULT_REVIEW_LOAD = 2
SENIOR_REVIEW_LOAD = 2
TECHNICAL_REVIEWERS_PER_PAPER = 1
STREAM_REVIEWERS_PER_PAPER = 2
MODEL_NAME = "all-MiniLM-L6-v2"
RANDOM_SEED = 42
np.random.seed(RANDOM_SEED)

USE_OpenReview_CoI = True
BUILD_AND_USE_INSTITUTION_CONFLICTS = False
MAX_COAUTHOR_YEARS = 5
SLEEP_BETWEEN_PROFILE_CALLS = 0.03
USE_OpenReview_Affinity = True

client=openreview.api.OpenReviewClient(baseurl=BASEURL,username=c.usr,password=c.pas)
submissions = client.get_all_notes(invitation=SUBMISSION_INVITATION)
print(f"Downloaded {len(submissions)} submissions.")

# INVITATIONS
venue_group = client.get_group(venue_id)
CoI_invitation = extract_field(venue_group.content, 'reviewers_conflict_id')
Affinity_invitation = extract_field(venue_group.content, 'reviewers_affinity_score_id')

venue = openreview.helpers.get_venue(client, venue_group.id)

# Empty CoI database
papers, reviewers, reviewers_id = [], {}, []
profile_cache, reviewer_coauthors, reviewer_institutions, paper_author_institutions = {},{},{}, {}
paper_affinity,reviewer_affinity,paper_conflicts = {},{},{} # CoI and Affinity from OpenReview. Contains a list per paper.

def set_all_authors_as_reviewers():
    authors = set()
    for paper in submissions:
        ids = extract_field(paper.content, 'authorids')
        for a_id in ids:
            if a_id not in (None, "", 'None'):
                authors.add(a_id)
    group = client.get_group(REVIEWER_GROUP)
    old_members = set(group.members)
    group.members = sorted(authors)
    client.post_group_edit( invitation=f"{venue_id}/-/Edit", signatures=[venue_id], group=group)
    print(f"Reviewers group updated: {len(old_members)} -> {len(authors)} members.")

def get_CoI_from_OpenReview():
    global paper_conflicts; paper_conflicts = {}
    """https://docs.openreview.net/how-to-guides/data-retrieval-and-modification/how-to-get-edges-for-conflicts-assignments-custom-max-papers-and-more"""
    conflict_invitation = f"{venue_id}/Reviewers/-/Conflict"
    grouped_edges = client.get_grouped_edges(invitation=conflict_invitation, groupby='head')
    for group in grouped_edges:
        paper_id = group['id']['head']
        reviewers = { edge['tail'] if isinstance(edge,dict) else edge.tail for edge in group['values']}
        paper_conflicts[paper_id]=reviewers
    pass

def get_Affinity_from_OpenReview():
    global paper_affinity, reviewer_affinity; paper_affinity,reviewer_affinity = {},{}
    """https://docs.openreview.net/how-to-guides/data-retrieval-and-modification/how-to-get-edges-for-conflicts-assignments-custom-max-papers-and-more"""
    grouped_edges = client.get_grouped_edges(invitation=Affinity_invitation, groupby='head')

    for group in grouped_edges:
        paper_id = group['id']['head']
        #reviewers = { edge['tail'] if isinstance(edge,dict) else edge.tail for edge in group['values']}
        affinities = {
            (edge["tail"] if isinstance(edge, dict) else edge.tail):
            (edge["weight"] if isinstance(edge, dict) else edge.weight)
            for edge in group["values"]}
        paper_affinity[paper_id]=affinities

    # Searchable by Reviewer
    reviewer_affinity = {}
    for paper_id, affs in paper_affinity.items():
        for reviewer, score in affs.items():
            reviewer_affinity.setdefault(reviewer, {})[paper_id] = score

    pass

def export_openreview_CoI_and_Affinity():
    get_CoI_from_OpenReview()
    get_Affinity_from_OpenReview()

    all_reviewers = sorted({
        reviewer
        for affs in paper_affinity.values()
        for reviewer in affs
    })

    rowscoi,rowsaff=[],[]
    for note in submissions:
        title = extract_field(note.content, "title")
        authors = extract_field(note.content, "authors")
        pap_num = note.number
        pap_id  = note.id
        cois   = paper_conflicts.get(pap_id,set())
        rowscoi.append(
                {"pap_id": pap_id, "pap_num": pap_num, "title": title,
                 "Authors": ", ".join(authors),
                 "CoI":     ", ".join(sorted(cois)) if cois else "",
                 })

        rowaff={"pap_id": pap_id, "pap_num": pap_num, "title": title, "Authors": ", ".join(authors)}
        affs = paper_affinity.get(pap_id, {})
        for reviewer in all_reviewers:
                rowaff[reviewer] = affs.get(reviewer, "")
        rowsaff.append(rowaff)

    fnamecoi=f"{VENUE}_COI_from_OpenReview.xlsx";
    dfcoi = pd.DataFrame(rowscoi)
    dfcoi.to_excel(fnamecoi, index=False)
    print(f"COI: Exported {len(dfcoi)} papers to '{fnamecoi}'.")
    fnameaff=f"{VENUE}_AFF_from_OpenReview.xlsx";
    dfaff = pd.DataFrame(rowsaff)
    dfaff.to_excel(fnameaff, index=False)
    print(f"AFF: Exported {len(dfaff)} papers to '{fnameaff}'.")
    pass

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
        ptype_int=int(ptype[0])-1
        Ptype = types[ptype_int]
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
            "paper_itype": ptype_int,
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
        if len(authors) >= 1: positions.append((authorids[ 0],authors[ 0], "first"))
        if len(authors) >  2: positions.append((authorids[ 1],authors[ 1], "second"))
        #if len(authors) >  3: positions.append((authorids[ 2],authors[ 2], "third"))
        if len(authors) >  3: positions.append((authorids[-2],authors[-2], "senior"))
        if len(authors) >  1: positions.append((authorids[-1],authors[-1], "senior"))

        for authorid, author, role in positions:
            if authorid not in reviewers:
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
                    "n_submissions": 1,
                    "expertise_texts": [],
                }
            else:
                reviewers[authorid]["roles"].add(role)
                reviewers[authorid]["streams"].add(paper["stream"])
                reviewers[authorid]["paper_types"].add(paper["paper_type"])
                reviewers[authorid]["papers_authored"].add(paper["paper_id"])
                reviewers[authorid]["n_submissions"]+=1

            kwd = " ".join(paper["keywords"])
            expertise = (normalize_text(paper["title"])+" "+normalize_text(paper["abstract"])+" "+normalize_text(kwd))
            reviewers[authorid]["expertise_texts"].append(expertise)

    reviewers_id = []
    for rid, data in reviewers.items(): reviewers_id.append(rid)

    reviewers_df = pd.DataFrame.from_dict(reviewers, orient="index")
    fname=f"{VENUE}_All_Reviewers.xlsx"
    reviewers_df.to_excel(fname , index=False)
    print(f"\nSaved to {fname}")
    #print(reviewers_id)
    print(f"Total reviewers: {len(reviewers)}")
    pass

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

def has_conflict(paper, reviewer_id):
    if reviewer_id in paper["authorids"]: return True  # self-review

    reviewer_cos = reviewer_coauthors.get( reviewer_id, set())
    for author_id in paper["authorids"]:
        if author_id in reviewer_cos: return True

    if BUILD_AND_USE_INSTITUTION_CONFLICTS:
        reviewer_inst = reviewer_institutions.get( reviewer_id, set())
        author_inst = paper_author_institutions.get( paper["paper_id"], set())
        overlap = reviewer_inst.intersection(author_inst)
        if len(overlap) > 0: return True
    if USE_OpenReview_CoI:
        if reviewer_id in paper_conflicts.get(paper["paper_id"],set()): return True
    return False

def reviewer_eligible(paper, reviewer_id):
    reviewer = reviewers[reviewer_id]
    if reviewer["current_load"] >= reviewer["max_load"]:  return False
    if has_conflict(paper, reviewer_id): return False
    # if paper["stream"] not in reviewer["streams"]: return False
    # if paper["paper_type"] not in reviewer["paper_types"]: return False
    if reviewer_id in EXCLUDE_REVIEWERS: return False # IWAI Reviewers etc
    return True

def reviewer_candidates(similarity_matrix):
    print("\nGenerating candidate reviewers...")
    candidate_rows = []
    paper_to_candidates = {}

    for paper_idx, paper in tqdm(list(enumerate(papers))):
        paper_id = paper["paper_id"]
        scores = []
        for reviewer_idx, reviewer_id in enumerate(reviewers_id):
            if not reviewer_eligible(paper, reviewer_id):  continue

            # ---- OPEN REVIEW AFFINITY; set 0.5 as default score for reviewers without proper openreview profile ----
            aff_oprev=0.0
            if USE_OpenReview_Affinity:
                aff_oprev = normalize_dict(paper_affinity[paper_id]).get(reviewer_id, 0.5)

            # ---- SUBMISSION-BASED SIMILARITIES ----
            sim = float(similarity_matrix[paper_idx, reviewer_idx])
            reviewer = reviewers[reviewer_id]

            # ---- PENALTIES ----
            load_penalty = reviewer["current_load"] / max(reviewer["max_load"],1)
            type_penalty = float(paper["paper_type"] not in reviewer["paper_types"])
            stream_penalty = float(paper["stream"] not in reviewer["streams"])

            # ---- FINAL SCORE ---- #+ np.random.normal(0, 0.01) \
            score = (sim *(1.0 + 0.3 * aff_oprev) - 0.15 * load_penalty - 0.15 * type_penalty - 0.15 * stream_penalty )
            scores.append( ( reviewer_id, score, sim ) )

        # ---- RANK REVIEWERS ----
        scores.sort(key=lambda x: x[1], reverse=True)
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
        if only_papers and paper["paper_itype"]==1: continue # Exclude Posters
        if paper["number"] in EXCLUDE_PAPERS: continue  # Exclude some special cases (test subm)
        selected_reviewers = []
        selected_reviewers[paper_idx] = set()
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
            if rid in selected_reviewers: continue
            reviewer = reviewers[rid]
            if reviewer["current_load"] >= reviewer["max_load"]:  continue
            #if paper["stream"] not in reviewer["streams"]: continue
            stream_candidates.append((rid, score, sim))
        stream_candidates = sorted( stream_candidates, key=lambda x: ( reviewers[x[0]]["current_load"], -x[1] ))
        for rid, score, sim in stream_candidates:
            reviewers[rid]["current_load"] += 1
            selected_reviewers.add(rid)
            selected_items.append({ "reviewer": rid, "role": "stream", "score": score, })
            if len([ x for x in selected_items  if x["role"] == "stream"]) >= STREAM_REVIEWERS_PER_PAPER: break

        # Export reviewers x paper
        for item in selected_items:
            rid = item["reviewer"]
            reviewer_submissions = [ p["title"]  for p in papers if rid in p["authorids"] ]
            assignments.append({
                "pap_type":         paper["paper_type"],
                "pap_stream":       paper["stream"],
                "pap_number":       paper["number"],
                "pap_title":        paper["title"],
                "pap_authors":      "; ".join(paper["authors"]),
                "reviewer":         rid,
                "score":            round(float(item["score"]), 2),
                "rev_role":         item["role"],
                "rev_tasks":        reviewers[rid]["current_load"],
                "rev_nsubm":        len(reviewers[rid]["papers_authored"]),
                "rev_streams":      reviewers[rid]["streams"],
                "rev_types":        reviewers[rid]["paper_types"],
                "rev_submis":       " | ".join(reviewer_submissions),
            })

    assignments_df = pd.DataFrame(assignments)
    fname2=f"{VENUE}_final_assignment2orcoi.xlsx"
    assignments_df.to_excel(fname2, index=False)
    print(f"\nSaved to {fname2}")

    loads = []
    for rid, data in reviewers.items(): loads.append(data["current_load"])
    loads = np.array(loads)
    print(f"Reviewer Load: Min/Mean/Max load: {loads.min()} / {loads.mean():.2f} / {loads.max()}")

def assignment3(paper_to_candidates):
    print("\nComputing final assignments ...")
    assignments = []
    selected_reviewers = {}
    selected_papers= {}
    selected_items = []

    # Computational reviewer
    for paper in papers:
        if only_papers and paper["paper_itype"]==1: continue # Exclude Posters
        if paper["number"] in EXCLUDE_PAPERS: continue  # Exclude some special cases (test subm)
        paper_id=paper["paper_id"]
        selected_reviewers[paper_id] = set()           # Init the set of reviewers per paper with an empty set
        selected_papers[paper_id] = set()           # Init the set of papers whose authors revise the current paper
        candidates = paper_to_candidates[paper_id]
        technical_candidates = []
        for rid, score, sim in candidates:
            reviewer = reviewers[rid]
            if reviewer["current_load"] >= reviewer["max_load"]:  continue
            if TECHNICAL_STREAM not in reviewer["streams"]:  continue
            technical_candidates.append((rid, score, sim))
        technical_candidates.sort(key=lambda x: (reviewers[x[0]]["current_load"], -float(x[1])))
        if technical_candidates:
            rid, score, sim = technical_candidates[0]
            reviewers[rid]["current_load"] += 1
            selected_reviewers[paper_id].add(rid)
            selected_papers[paper_id].update(reviewers[rid].get("papers_authored",set()))
            selected_items.append({ "paper_id":paper_id,  "reviewer": rid, "role": "technical", "score": score, })

    # 2 same-stream reviewers
    for _ in range(2):
        for paper in papers:
            if only_papers and paper["paper_itype"] == 1: continue  # Exclude Posters
            if paper["number"] in EXCLUDE_PAPERS: continue  # Exclude some special cases (test subm)
            paper_id = paper["paper_id"]
            candidates = paper_to_candidates[paper_id]
            stream_candidates = []
            for rid, score, sim in candidates:
                if rid in selected_reviewers[paper_id]: continue
                reviewer = reviewers[rid]
                if reviewer["current_load"] >= reviewer["max_load"]:  continue
                if selected_papers[paper_id].intersection(reviewer.get("papers_authored",set())): continue
                #if paper["stream"] not in reviewer["streams"]: continue # This is through penalty
                stream_candidates.append((rid, score, sim))
            stream_candidates.sort(key=lambda x: ( reviewers[x[0]]["current_load"], -x[1] ))
            if stream_candidates:
                rid, score, sim = stream_candidates[0]
                reviewers[rid]["current_load"] += 1
                selected_reviewers[paper_id].add(rid)
                selected_papers[paper_id].update(reviewers[rid].get("papers_authored", set()))
                selected_items.append({ "paper_id":paper_id,  "reviewer": rid, "role": "stream", "score": score, })

    # Export reviewers x paper
    paper_dic = {p["paper_id"]: p for p in papers}
    for item in selected_items:
        paper_id = item["paper_id"]
        rid = item["reviewer"]
        paper=paper_dic[paper_id]
        reviewer_submissions = [ p["title"]  for p in papers if rid in p["authorids"] ]
        reviewer_subm_numbs = [ p["number"]  for p in papers if rid in p["authorids"] ]
        assignments.append({
            "pap_number":       paper["number"],
            "pap_id":           paper_id,
            "pap_title":        paper["title"],
            "pap_authors":      "; ".join(paper["authors"]),
            "pap_type":         paper["paper_type"],
            "pap_stream":       paper["stream"],
            "rev_id":           rid,
            "score":            round(float(item["score"]), 2),
            "rev_role":         item["role"],
            "rev_tasks":        reviewers[rid]["current_load"],
            "rev_nsubm":        len(reviewers[rid]["papers_authored"]),
            "rev_streams":      reviewers[rid]["streams"],
            "rev_types":        reviewers[rid]["paper_types"],
            "rev_subm_n":       " ".join(map(str, reviewer_subm_numbs)),
            "rev_submis":       " | ".join(reviewer_submissions),
        })

    assignments_df = pd.DataFrame(assignments)
    assignments_df = assignments_df.sort_values(by=["pap_number", "rev_role", "score"], ascending=[True, False, False])
    fname3=f"{VENUE}_Final3_assignment.xlsx"
    assignments_df.to_excel(fname3, index=False)
    print(f"\nSaved to {fname3}")


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

def compute_assignments():
    get_submissions()
    extract_reviewers()
    get_profiles()
    if BUILD_AND_USE_INSTITUTION_CONFLICTS:  build_CoI_database()
    if USE_OpenReview_CoI: get_CoI_from_OpenReview()
    if USE_OpenReview_Affinity: get_Affinity_from_OpenReview()
    similarity = paper_to_reviewer_match()
    paper_to_candidates = reviewer_candidates(similarity)
    assignment3(paper_to_candidates)
    print("\nDone.")

def get_anon_id(paper_number, reviewer_id):
    #             anon_group_id = get_anon_id(paper_number, reviewer_id)
    #           Use anon group ID as reader if available, else fall back to reviewer_id
    #              reviewer_reader = anon_group_id if anon_group_id else reviewer_id

    """Returns the anonymous group ID for a reviewer on a given paper."""
    anon_groups = client.get_groups( prefix=f"{venue_id}/Submission{paper_number}/Reviewer_",   member=reviewer_id )
    if anon_groups: return anon_groups[0].id
    return None  # fallback: reviewer has no anon group yet

def upload_assignments():
    fname=f"{VENUE}-assignments.xlsx"
    df = pd.read_excel(fname)
    grouped = df.groupby("pap_id")

    for paper_id, group in grouped:
        paper_number = group["pap_number"].iloc[0]
        paper_reviewers_group_id = f"{venue_id}/Submission{paper_number}/Reviewers"
        paper_authors_group_id = f"{venue_id}/Submission{paper_number}/Authors"

        # Rebuild assignments per paper_id
        edges = []
        paper_reviewers = set()

        for _, row in group.iterrows():
            reviewer_id=row["rev_id"]
            paper_reviewers.add(reviewer_id)
            edges.append(
                openreview.api.Edge(invitation=ASSIGNMENT_INVITATION,
                    head=paper_id,
                    tail=reviewer_id,   #label="Reviewer", # ???
                    weight=1,
                    readers=[venue_id, reviewer_id],
                    nonreaders=[paper_authors_group_id],
                    writers=[venue_id],  signatures=[venue_id] )
            )

        try:
            # Current and New set of reviewers
            paper_reviewer_group = client.get_group(paper_reviewers_group_id) # Current reviewers
            # Handle reviewer's groupd
            current_reviewers = set(paper_reviewer_group.members or [])
            desired_reviewers = paper_reviewers
            # Make list of reviewers to remove and add for the current submission
            reviewers_to_add = list(desired_reviewers - current_reviewers)
            reviewers_to_remove = list(current_reviewers - desired_reviewers)
            # Deploy the reviewers
            if reviewers_to_remove:
                client.remove_members_from_group(group=paper_reviewers_group_id, members=reviewers_to_remove)
            if reviewers_to_add:
                client.add_members_to_group(group=paper_reviewers_group_id, members=reviewers_to_add)

            # Post assignments
            existing = client.get_edges(invitation=ASSIGNMENT_INVITATION, head=paper_id)
            if existing:    # first, remove all old assignments
                client.delete_edges(invitation=ASSIGNMENT_INVITATION, head=paper_id, soft_delete=False)
            if edges:       # then, add all new assignments
                openreview.tools.post_bulk_edges(client=client, edges=edges)
            print(f"Paper {paper_number}:  {len(edges)} assignments uploaded.")

            # Update paper reviewer's group to allow access to reviewers IDs only to Admin
            # client.post_group_edit(
            #     invitation=f"{venue_id}/-/Edit", signatures=[venue_id],
            #     group=openreview.api.Group(
            #         id=paper_reviewer_group.id,
            #         readers=[venue_id],
            #         writers=paper_reviewer_group.writers,
            #         signatures=paper_reviewer_group.signatures,
            #         members=paper_reviewer_group.members,
            #         nonreaders=paper_reviewer_group.nonreaders,
            #         anonids=paper_reviewer_group.anonids,
            #         deanonymizers=paper_reviewer_group.deanonymizers
            #     )
            # )

        except Exception as e:
            print(f"ERROR while processing paper {paper_number} ({paper_id}): {e}")

        # ---------- NOTIFICATION -----------
        try:
            BODY = dedent(""" Dear {VENUE} Contributor,
            
            You have been assigned as a Reviewer for Submission number {p_number}.
            
            Title: {p_title}

            To view the assignment, please visit: https://openreview.net/forum?id={p_id}
            
            or click on Tasks: https://openreview.net/tasks
            
            Please submit your review by 5 July.
            
            We thank you for your objective, constructive, and timely reviews.
            
            The {VENUE} Technical Program Chairs
            """)
            REPLYTO = "ivilinpeev.stoianov@cnr.it"
            REVIEWERS_MESSAGE_INVITATION = f"{venue_id}/Reviewers/-/Message"

            for _, row in group.iterrows():
                pap_number = int(row["pap_number"])
                pap_id = row["pap_id"]
                pap_title = row["pap_title"]
                rev_id = row["rev_id"]
                body = BODY.format(p_number=pap_number, p_title=pap_title, p_id=pap_id, VENUE=VENUE2)
                SUBJ = f"{VENUE}: You have been assigned as a Reviewer for Submission n.{pap_number}"
                client.post_message( invitation=REVIEWERS_MESSAGE_INVITATION,
                    recipients=[rev_id], subject=SUBJ, message=body, replyTo=REPLYTO, signature=venue_id)

        except Exception as e:
            print(f" NOTIFICATION FAILED for {rev_id}: {e}")

    print(f"Successfully deleted previous assignments and uploaded {len(df)} assignments of {len(grouped)} papers")

def after_upload():
    reviewer_committee_id = venue.get_reviewers_id()  # f'{venue_id}/Reviewers'  o 'IWAI/2026/Workshop/Reviewers'
    venue.set_assignments(
        assignment_title=ASSIGNMENT_INVITATION,
        committee_id=reviewer_committee_id,      # ('IWAI/2026/Workshop/Reviewers')
        overwrite=False
    )
    print(f" Deployment of {ASSIGNMENT_INVITATION} complete.")

def check_papers():
    reviewers_group_id = venue.get_reviewers_id()  # f'{venue_id}/Reviewers'  o 'IWAI/2026/Workshop/Reviewers'
    print(reviewers_group_id)
    g=client.get_group(reviewers_group_id)
    print(g.members)


    #submissions = client.get_notes(invitation=venue.get_submissions_id())
    print(f"Checking reader permissions across {len(submissions)} submissions...")
    print("-" * 60)
    issues_found = 0

    for submission in submissions:
        # API V2 Note objects store reader permissions directly in note.readers
        paper_id = submission.id
        paper_n  = submission.number
        paper_title = submission.content.get('title', {}).get('value', 'No Title')
        readers_list = submission.readers

        # Check if the global Reviewers group is allowed to read this paper
        # NOTE: Depending on your workflow, this might look for the global group
        # OR specific anonymous reviewer IDs like 'VENUE_ID/Submission1/Reviewers'
        is_group_present = reviewers_group_id in readers_list

        # Check if paper-specific assigned reviewer groups are present
        paper_reviewers_id = f"{venue_id}/Submission{submission.number}/Reviewers"
        is_paper_group_present = paper_reviewers_id in readers_list

        #print(f"Submission{paper_n} readers = {readers_list}")

    print(f"------- ASSIGNMENTS -------")

    fname=f"{VENUE}-assignments.xlsx"
    df = pd.read_excel(fname)
    grouped = df.groupby("pap_id")


    for paper_id, group in grouped:
        paper_number = group["pap_number"].iloc[0]
        paper_reviewers_group_id = f"{venue_id}/Submission{paper_number}/Reviewers"
        paper_reviewer_group = client.get_group(paper_reviewers_group_id)  # Current reviewers
        paper_reviewers = set(paper_reviewer_group.members or [])
        print(paper_reviewers_group_id)
        print(paper_reviewers)
        assignments = client.get_edges(invitation=ASSIGNMENT_INVITATION, head=paper_id)
        if assignments:
            print(assignments)
        pass

def check5_submission_groups():
    print("CHECK5");   missing_groups, empty_groups = [],[]
    #for s in submissions:
    fname=f"{VENUE}-assignments.xlsx"
    df = pd.read_excel(fname)
    grouped = df.groupby("pap_id")
    for paper_id, group in grouped:
        number = group["pap_number"].iloc[0]
        #number   = s.number
        group_id = f"{venue_id}/Submission{number}/Reviewers"
        try:
            group = client.get_group(group_id)
            if not group.members:
                empty_groups.append(group_id)
                print(f"Submission{number}/Reviewers exists but has NO members → reviewer cannot be matched")
            else:
                print(f"Submission{number}/Reviewers — {len(group.members)} member(s)")
        except openreview.OpenReviewException:
            missing_groups.append((number, group_id))
            print(f"Submission{number}/Reviewers — group does NOT exist")

def check14_invitations():
    # Name used to build review invitation ids (usually "Official_Review")
    REVIEW_INVITATION_NAME = "Official_Review"
    fname=f"{VENUE}-assignments.xlsx"
    df = pd.read_excel(fname)
    grouped = df.groupby("pap_id")
    for paper_id, group in grouped:
        number = group["pap_number"].iloc[0]
        inv_id     = f"{venue_id}/Submission{number}/-/{REVIEW_INVITATION_NAME}"
        per_paper_group = f"{venue_id}/Submission{number}/Reviewers"
        top_level_group = f"{venue_id}/Reviewers"
        print(f"\n  Paper #{number}  ({inv_id})")

        try:
            inv = client.get_invitation(inv_id)
        except openreview.OpenReviewException:
            print(f"Invitation not found — skipping checks 1–4 for this paper")
            continue

def Post_Submission_Ready_to_Review():
    """ Set reviewers readers of the submissions. """
    venue.post_submission_stage(submission_readers=['areachairs', 'reviewers'])

# -------------------------------------------------------
def Assignments_1_Setup():
    # 1. In OpenReview GUI create an assignment configuration
    # Status now is "Initialized"

    # 2. Change the status to "Complete"
    assignment_title = f'{VENUE}-PythonAssignments'
    AssignmentConfig_Invitation = f'{venue_id}/Reviewers/-/Assignment_Configuration'
    reviewer_committee_id = venue.get_reviewers_id()  # f'{venue_id}/Reviewers'  o 'IWAI/2026/Workshop/Reviewers'

    assignment_config_note = client.get_all_notes(
        invitation = AssignmentConfig_Invitation,
        content={'title': assignment_title})[0]

    content = assignment_config_note.content.copy()
    content["status"] = {"value": "Complete"}

    client.post_note_edit( invitation= AssignmentConfig_Invitation, signatures=[venue_id],
        note = openreview.api.Note( id=assignment_config_note.id, content = content ) )

def Assignments_2_Upload():
    assignment_title = f'{VENUE}-PythonAssignments'
    reviewer_committee_id = venue.get_reviewers_id()  # f'{venue_id}/Reviewers'  o 'IWAI/2026/Workshop/Reviewers'
    proposed_assignment_invitation_id = venue.get_assignment_id(committee_id = reviewer_committee_id, deployed = False) # 'IWAI/2026/Workshop/Reviewers/-/Proposed_Assignment'

    fname=f"{VENUE}-assignments.xlsx"
    df = pd.read_excel(fname)
    grouped = df.groupby("pap_id")

    for paper_id, group in grouped:
        paper_number = group["pap_number"].iloc[0]
        paper_reviewers_group_id = f"{venue_id}/Submission{paper_number}/Reviewers"
        paper_authors_group_id = f"{venue_id}/Submission{paper_number}/Authors"

        edges = []
        # Get group IDs for each role for this paper
        paper_reviewer_id = venue.get_reviewers_id(number=paper_number)  # E.g. "venue_id/Submission5/Reviewers",
        #paper_ac_id = venue.get_area_chairs_id(number=paper_number)
        #paper_sac_id = venue.get_senior_area_chairs_id(number=paper_number)
        paper_author_id = venue.get_authors_id(number=paper_number)

        # Build Edges for each reviewer
        for _, row in group.iterrows():
            reviewer_id=row["rev_id"]
            edges.append(
                openreview.api.Edge(invitation=proposed_assignment_invitation_id,
                    head=paper_id,
                    tail=reviewer_id,
                    weight=1,
                    label=assignment_title,
                    readers=[venue_id, reviewer_id], # paper_sac_id, paper_ac_id,
                    nonreaders=[paper_author_id],
                    writers=[venue_id], # paper_sac_id, paper_ac_id,
                    signatures=[venue.get_program_chairs_id()]  # E.g. venue_id/Program_Chairs
            ))
        try:
            openreview.tools.post_bulk_edges(client=client, edges=edges)
        except Exception as e:
            print(f"ERROR while posting assignments for paper {paper_number} ({paper_id}): {e}")
        print(f"Assignments for Paper {paper_number} (n={len(edges)}) uploaded.")

    pass

def Assignments_3_Deploy():
    assignment_title = f'{VENUE}-PythonAssignments'
    reviewer_committee_id = venue.get_reviewers_id()  # f'{venue_id}/Reviewers'  o 'IWAI/2026/Workshop/Reviewers'
    # 3. Deploy the finalized configuration to the Reviewers committee
    venue.set_assignments(
        assignment_title=assignment_title,      # The exact title of your draft match
        committee_id=reviewer_committee_id,     # The reviewer group ID ('IWAI/2026/Workshop/Reviewers')
        overwrite=True                          # True replaces old deployments; False appends them
    )
    print(f" Deployment of {assignment_title} complete.")

def Assignments_X_Unset():
    assignment_title = f'{VENUE}-PythonAssignments'
    reviewer_committee_id = venue.get_reviewers_id()  # f'{venue_id}/Reviewers'  o 'IWAI/2026/Workshop/Reviewers'
    # 3. Deploy the finalized configuration to the Reviewers committee
    venue.unset_assignments(
        assignment_title=assignment_title,      # The exact title of your draft match
        committee_id=reviewer_committee_id,     # The reviewer group ID ('IWAI/2026/Workshop/Reviewers')
    )
    print(f" Deployment of {assignment_title} complete.")

# -----------------------------------------------------

def skeleton_for_submission_readers_change():
    # Fetch the specific submission note
    note_id = "YOUR_SUBMISSION_FORUM_ID"
    submission = client.get_note(note_id)
    # Construct the exact assigned groups (replace with your specific venue path)
    assigned_reviewers = f"{venue_id}/Submission{submission.number}/Reviewers"
    # Add the Program Chairs & Assigned Groups as the only entities allowed to read
    submission.readers = [ f"{venue_id}/Program_Chairs", assigned_reviewers  ]
    # Update the submission in API v2
    client.post_note(submission)



# IWAI/2026/Workshop/-/Submission&content.venueid=IWAI/2026/Workshop/Submission


def check_groupedit():
    # invitations = client.get_invitations(prefix='IWAI/2026/Workshop/Submission31')
    # for inv in invitations:
    #     print(inv.id, inv.invitees)

    gr_id='IWAI/2026/Workshop/Submission31/Reviewers'

    print("--REVIEWERS Subm31--")
    group = client.get_group(gr_id)
    print(group)
    client.post_group_edit(
        invitation=f"{venue_id}/-/Edit", signatures=[venue_id],
        group=openreview.api.Group(
            id=group.id,
            readers=[venue_id],
            writers=group.writers,
            signatures=group.signatures,
            members=group.members,
            nonreaders=group.nonreaders,
            anonids=group.anonids,
            deanonymizers=group.deanonymizers
        )
    )

    print("-- POST EDIT --")
    group = client.get_group('IWAI/2026/Workshop/Submission31/Reviewers')
    print(group)

    pass

def check2():
    edges = client.get_all_edges(invitation = ASSIGNMENT_INVITATION, tail = " ~Ivilin_Peev_Stoianov1") #head = 'zqV2uduW5y'
    print(edges)

# https://api2.openreview.net/edges/count?invitation=IWAI/2026/Workshop/Reviewers/-/Assignment

def check3():
    proposed_assignment_invitation_id = client.get_assignment_id(committee_id = REVIEWER_GROUP, deployed = False)
    print(proposed_assignment_invitation_id)

def check_assignments():
    fname = f"{VENUE}-assignments.xlsx"
    df = pd.read_excel(fname)
    grouped = df.groupby("pap_id")

    for paper_id, group in grouped:
        existing = client.get_edges(invitation=ASSIGNMENT_INVITATION, head=paper_id)
        if existing:
            print(existing)

def debug_reviewer_state():
    reviewer_id='~Ivilin_Stoianov1'
    paper_id='31'
    groups = [f"{venue_id}/Submission{paper_id}/Reviewers"]

    # 1. Reviewer membership
    print("1. Reviewer group membership check")
    for g in groups:
        try:
            group = client.get_group(g)
            print(f"\nGroup: {g}")
            print(f"Members includes reviewer {reviewer_id}:", reviewer_id in (group.members or []))
            print("Readers:", group.readers)
        except Exception as e:
            print(f"ERROR reading group {g}: {e}")

    # 2. Assignment edge
    print("\n2. Assignment edges")

    edges = client.get_edges(invitation=ASSIGNMENT_INVITATION, tail=reviewer_id)
    print(f"Edges found for reviewer: {len(edges)}")
    for e in edges[:3]:
        print("\nEdge sample:")
        print("head (paper):", e.head)
        print("tail (reviewer):", e.tail)
        print("readers:", e.readers)

    # 3. Edge visibility
    print("\n3. Edge visibility (simulate reviewer access)")

    try:
        visible_edges = client.get_edges( invitation=ASSIGNMENT_INVITATION, tail=reviewer_id )
        print("Reviewer CAN see their own edges:", len(visible_edges) > 0)
    except Exception as e:
        print("Reviewer CANNOT query edges:", e)

    # 4. Invitation inspect

    try:
        inv = client.get_invitation(ASSIGNMENT_INVITATION)
        print("Invitation ID:", inv.id)
        print(inv)

        if hasattr(inv, "reply"):
            print("\nReply fields:")
            try:    print("reply.readers:", inv.reply.get("readers"))
            except: print("reply.readers: not found")

            try:    print("reply.signatures:", inv.reply.get("signatures"))
            except: print("reply.signatures: not found")

        print("\nInvitees:", getattr(inv, "invitees", None))
        print("Readers:", getattr(inv, "readers", None))

    except Exception as e: print("ERROR reading invitation:", e)

def export_reviewers_and_authors_per_submission():
    ass1,ass2 = [],[]
    maxr1, maxr2, maxa, maxr = 0,0,0,0
    reviewer_to_subs = defaultdict(list)
    pap={}
    for s in submissions:
        paper_id=s.id
        number = getattr(s, "number", 0)
        typ=s.content.get("type")
        typ = int(typ["value"][0]) if typ else 1 # Fallback for previous IWAI editions without Type.
        strm = extract_field(s.content, "stream")
        stream = streams[int(strm[0])-1] if strm else ""
        title= s.content['title']['value'] if isinstance(s.content['title'], dict) else s.content['title']
        authors = s.content['authorids']['value']
        maxa = max(maxa,len(authors))
        reviewer_edges=client.get_edges(invitation=f'{venue_id}/Reviewers/-/Assignment', head=paper_id)
        reviewers = [edge.tail for edge in reviewer_edges]
        for r in reviewers: reviewer_to_subs[r].append(number)
        maxr = max(maxr, len(reviewers))
        entry = {'ID':paper_id, 'number':number, 'stream':stream, 'title':title, 'reviewers':reviewers, 'authors':authors}
        if typ==1: ass1.append(entry);  maxr1=max(maxr1,len(reviewers))
        else:      ass2.append(entry);  maxr2=max(maxr2,len(reviewers))

    ass1.sort(key=lambda x: (x['stream'], x['number']))
    ass2.sort(key=lambda x: (x['stream'], x['number']))

    # Create workbook
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Full Papers"
    ws2 = wb.create_sheet("Extended Abstracts")

    def format_reviewer(rid):
        ass_pap = reviewer_to_subs[rid]
        return f"{rid} - {len(ass_pap)}\n[{','.join(str(p) for p in ass_pap)}]" # "" for new line

    def format_author(aid):
        #return f"{aid}"
        ass_pap = reviewer_to_subs.get(aid, [])
        if ass_pap:
            return f"{aid} - {len(ass_pap)}\n[{','.join(str(p) for p in ass_pap)}]"
        else:
            return aid

    def write_sheet(ws, assignment_list):
        header = ( ["ID", "stream", "title"]
            + [f"REV_{i+1}" for i in range(4)] + [f"AUT_{i+1}" for i in range(10)])
        ws.append(header)

        for a in assignment_list:
            revs = [format_reviewer(r) for r in a['reviewers'][:4]]
            revs += [""] * (4 - len(revs))

            auts = [format_author(auth) for auth in a['authors'][:10]]
            auts += [""] * (10 - len(auts))

            row = [a['number'], a['stream'], a['title']] + revs + auts
            ws.append(row)

    write_sheet(ws1, ass1)
    write_sheet(ws2, ass2)
    filename=f"{VENUE}-submissions-reviewers-authors.xlsx"
    wb.save(filename)
    print(f"Exported reviewer assignments and authors to {filename}")

def communicate_review_guidelines():
    from messages.review_guidelines import MSG,SBJ
    REPLYTO = "ivilinpeev.stoianov@cnr.it"
    AUTHORS = set()
    for s in submissions: AUTHORS.update(s.content['authorids']['value'])
    try:client.post_message(invitation=MSG_INVITATION,
            recipients=list(AUTHORS), subject=SBJ,  message=MSG.format(VENUE=VENUE2),
            replyTo=REPLYTO, signature=venue_id)
    except Exception as e: print(f"NOTIFICATION FAILED: {e}")
    print(f"Successfully notified {len(AUTHORS)} authors")


if __name__ == '__main__':
    # 1. Set reviewers
    #set_all_authors_as_reviewers()
    # 2. GUI -> Compute Paper Matching (with Comprehensive Conflict computation and Specter2+SciIncl; takes 10-15 min)
    # 2a (optional): Export Conflict-of-Interest
    # get_CoI()
    # 3. Compute assignments (takes 5 min)
    #compute_assignments()

    # 4. Upload assignments
    #upload_assignments()
    #after_upload()

    #check_groupedit()
    #check_assignments()
    #export_openreview_CoI_and_Affinity()
    #debug_reviewer_state()

    #Assignments_1_Setup()
    #Assignments_2_Upload()
    #Assignments_3_Deploy()
    #Assignments_X_Unset()

    # check_papers()

    #check5_submission_groups()
    #check14_invitations()

    # export_reviewers_and_authors_per_submission()

    communicate_review_guidelines()

    pass


# TO DO
# + add all authors as reviewers
# x verify institution
# + verify CoI from openreview
# ADD Assignments to openreview
# - do it first for a test list and check when assignments become visible and active
