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

import os
from concurrent.futures import ThreadPoolExecutor

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
red_fill = PatternFill( fill_type="solid", start_color="FFC7CE", end_color="FFC7CE" )
green_fill = PatternFill(fill_type="solid", start_color="C6EFCE", end_color="C6EFCE" )
yellow_fill = PatternFill(fill_type="solid", start_color="FFEB9C", end_color="FFEB9C" )
from openpyxl.styles import Alignment
center = Alignment( horizontal="center", vertical="center" )
center_wrap = Alignment( horizontal="center", vertical="center", wrap_text = True )

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
REVIEW_INVITATION = f'{venue_id}/Submission*/-/Official_Review'
MSG_INVITATION=f'{venue_id}/-/Edit'


client=openreview.api.OpenReviewClient(baseurl=BASEURL,username=c.usr,password=c.pas)
submissions = client.get_all_notes(invitation=SUBMISSION_INVITATION, sort='number:asc', details='replies')

print(f"Downloaded {len(submissions)} submissions.")

assigned_reviews = defaultdict(set)
completed_reviews = defaultdict(set)
missing_reviews = defaultdict(set)

decision_map = {
    "A+P": "Accept (Proceedings Paper + Oral Presentation)",
    "A+sp": "Accept (Proceedings Paper + Poster + Spotlight Presentation)",
    "P+sp": "Accept (Poster + Spotlight Presentation)",
    "P": "Accept (Poster)",
    "R": "Reject"}

decision_map2 = {
    "Accept (Proceedings Paper + Oral Presentation)":"A+pres",
    "Accept (Proceedings Paper + Poster + Spotlight Presentation)":"A+spot",
    "Accept (Poster + Spotlight Presentation)":"Po+sp",
    "Accept (Poster)":"Po",
    "Reject":"Rej"}

#reviewer2anonym, anonym2reviewer = {}, {}
from Lists.iwai2026_anonym2reviewers import anonym2reviewer

# https://docs.openreview.net/how-to-guides/data-retrieval-and-modification/how-to-view-camera-ready-revisions

def extract_field(content, key):
    """  OpenReview V2 stores values as:  content[key]["value"] or content[key]  """
    if key not in content:  return None
    value = content[key]
    if isinstance(value, dict) and "value" in value:
        return value["value"]
    return value

def download_cameraready_serial():
    # Location for downloads
    fpath = f"{YEAR}/camera-ready/"
    os.makedirs(fpath, exist_ok=True);
    # Download & save
    for s in submissions:
        title = s.content["title"]["value"]
        print(f"{s.number}-{title} .. ")
        if (s.content.get("camera_ready_pdf")):
            fname = f'{s.number}-{title}-pdf.pdf'
            if not os.path.exists(fpath+fname):
                f = client.get_attachment(id=s.id, field_name='camera_ready_pdf')
                with open(fpath+fname, 'wb') as op: op.write(f)
        if (s.content.get("camera_ready_package")):
            fname = f'{s.number}-{title}-zip.zip'
            if not os.path.exists(fpath+fname):
                f = client.get_attachment(id=s.id, field_name='camera_ready_package')
                with open(fpath+fname, 'wb') as op: op.write(f)
    print('DONE')

def get_decision(paper):
    paper_id = paper.number
    tit = extract_field(paper.content, "title")
    ids = extract_field(paper.content, 'authorids')
    typ = int(extract_field(paper.content, "type")[0])
    if typ == 1:  # Full paper
        for reply in paper.details["replies"]:
            for invitation in reply["invitations"]:
                if invitation.endswith("Decision"):
                    decision = reply["content"]["decision"]["value"]
                    return decision_map2[decision]


def save_attachment(client, s, field_name, fpath):
    if s.content.get(field_name):
        title = s.content.get("title", {}).get("value", f"s{s.number}")
        ext = "pdf" if field_name == "camera_ready_pdf" else "zip"
        fname = f"{s.number}-{title}.{ext}"
        full_path = os.path.join(fpath, fname)
        if not os.path.exists(full_path):
            f = client.get_attachment(id=s.id, field_name=field_name)
            with open(full_path, 'wb') as op:
                op.write(f)

def download_cameraready_parallel(client,submissions,YEAR):
    fpath = f"{YEAR}/camera-ready/"
    os.makedirs(fpath, exist_ok=True)
    flist = os.path.join(f"{YEAR}/",f"IWAI-{YEAR}-CameraReady-State.xlsx")

    # 1. Prepare tasks: one for PDF, one for ZIP if available
    # 2. Make a dictionary to export in xls with info about the availability of pdf/zip

    tasks = []          # List of pdf/zip to download
    has_cameraready={}  # Dictionary with availability of pdf/zip
    for s in submissions:
        decision=get_decision(paper=s)
        if decision and decision[0] == 'A': # Accept for proceedings (with presentation or spotlight)
            title = s.content.get("title", {}).get("value", f"s{s.number}")
            entry = f"{s.number}-{title}"
            has_cameraready[s.number]=[decision,title,"",""]
            for i,field in enumerate(["camera_ready_pdf", "camera_ready_package"]):
                if s.content.get(field):
                    tasks.append((client, s, field, fpath))
                    has_cameraready[s.number][i+2] = "+"

    df = pd.DataFrame([[key] + values for key, values in has_cameraready.items()], columns=["ID","Decision","Title","PDF","ZIP"])
    df = df.sort_values(by=["Decision","ID"], ascending=[True,True])
    df.to_excel(flist, index=False)

    # Download in parallel
    with ThreadPoolExecutor(max_workers=10) as executor:
        executor.map(lambda args: save_attachment(*args), tasks)

    print('DONE')


if __name__ == '__main__':
    #download_cameraready_serial()
    download_cameraready_parallel(client, submissions, YEAR)
