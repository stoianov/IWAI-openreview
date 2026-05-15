import openreview
import config_ivo as c
import pandas as pd
import os
import time
import pprint

# DEFINITIONS
BASE_URL = 'https://api2.openreview.net'
YEAR='2026'
venue = f'IWAI{YEAR}'
venue_id    = f'IWAI/{YEAR}/Workshop'
SUBMISSION_INVITATION = f'{venue_id}/-/Submission'
ASSIGNMENT_INVITATION = f'{venue_id}/Reviewers/-/Assignment'
MSG_INVITATION=f'{venue_id}/-/Edit'
REVIEWER_GROUP = f'{venue_id}/Reviewers'
AUTHORS_GROUP = f'{venue_id}/Authors'

# CLIENT
client=openreview.api.OpenReviewClient(baseurl=BASE_URL,username=c.usr,password=c.pas)
venue_group = client.get_group(venue_id)
submission_str   = venue_group.content['submission_name']['value'] # this query results in text "Submission"
streams = ['1-comp','2-cogn','3-appl']
types=['1-paper','2-abstr']
submissions = client.get_all_notes(invitation=SUBMISSION_INVITATION, sort='number:asc', details='replies')

def list_of_all_authors():
    author_ids = set()
    for paper in submissions:
        ids = paper.content.get('authorids', [])["value"]
        for a_id in ids:
            if a_id and a_id != 'None':
                author_ids.add(a_id)
    print(f"Found {len(author_ids)} unique author IDs.")
    pprint.pprint(author_ids)
    print(f"Found {len(author_ids)} unique author IDs.")


if __name__ == '__main__':
    list_of_all_authors()