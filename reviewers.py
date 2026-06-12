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
from openpyxl import Workbook

import config_ivo as c
from collections import defaultdict
import time

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
venue_id = f'IWAI/{YEAR}/Workshop'
SUBMISSION_INVITATION = f'{venue_id}/-/Submission'
ASSIGNMENT_INVITATION = f'{venue_id}/Reviewers/-/Assignment'
COI_INVITATION = f"{venue_id}/Reviewers/-/Conflict"
REVIEWER_GROUP = f'{venue_id}/Reviewers'

client=openreview.api.OpenReviewClient(baseurl=BASEURL,username=c.usr,password=c.pas)
submissions = client.get_all_notes(invitation=SUBMISSION_INVITATION)
print(f"Downloaded {len(submissions)} submissions.")

def extract_field(content, key):
    """  OpenReview V2 stores values as:  content[key]["value"] or content[key]  """
    if key not in content:  return None
    value = content[key]
    if isinstance(value, dict) and "value" in value:
        return value["value"]
    return value

def reviewers_and_authors():
    # paper_reviewers = client.get_group(f"{VENUE}/Submission31/Reviewers")
    # print(paper_reviewers.members)

    ass1,ass2 = [],[]
    maxr1, maxr2, maxa, maxr = 0,0,0,0
    reviewer_to_subs = defaultdict(list)
    pap={}
    for s in submissions:
        paper_id=s.id
        number = getattr(s, "number", 0)
        #if not number == 31: continue
        typ=s.content.get("type")
        typ = int(typ["value"][0]) if typ else 1 # Fallback for previous IWAI editions without Type.
        strm=s.content.get("stream")["value"]
        stream = streams[int(strm[0])-1]
        title= s.content['title']['value'] if isinstance(s.content['title'], dict) else s.content['title']
        authors = s.content['authorids']['value']
        maxa = max(maxa,len(authors))

        rev_edges2=client.get_edges(invitation=ASSIGNMENT_INVITATION, head=paper_id)
        reviewers = [edge.tail for edge in rev_edges2]

        for r in reviewers: reviewer_to_subs[r].append(number)
        maxr = max(maxr, len(reviewers))

        entry = {'ID':paper_id, 'number':number, 'stream':stream, 'title':title, 'reviewers':reviewers, 'authors':authors }
        if typ==1:
            ass1.append(entry)
            maxr1=max(maxr1,len(reviewers))
        else:
            ass2.append(entry)
            maxr2=max(maxr2,len(reviewers))

    ass1.sort(key=lambda x: (x['stream'], x['number']))
    ass2.sort(key=lambda x: (x['stream'], x['number']))

    # Create workbook
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Papers"
    ws2 = wb.create_sheet("Posters")
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

    write_sheet(ws1, ass1);  write_sheet(ws2, ass2)
    filename=f"{VENUE}-reviewers-and-authors.xlsx"
    wb.save(filename)
    print(f"Exported reviewer assignments and authors to {filename}")

if __name__ == '__main__':
    reviewers_and_authors()
