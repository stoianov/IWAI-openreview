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

# SUBMISSION TYPES
streams = ['1-comp','2-cogn','3-appl']
types=['1-paper','2-abstr']

submissions = client.get_all_notes(invitation=SUBMISSION_INVITATION, sort='number:asc', details='replies')

def authors():
    author_ids = set()
    for paper in submissions:
        ids = paper.content.get('authorids', [])["value"]
        for a_id in ids:
            if a_id and a_id != 'None':
                author_ids.add(a_id)
    print(f"Found {len(author_ids)} unique author IDs.")
    pprint.pprint(author_ids)
    print(f"Found {len(author_ids)} unique author IDs.")

def monitor():
    authP,authA = {},{}; i_p,i_a = 0,0
    for i,s in enumerate(submissions):
        subm_id = s.number
        aid = s.content['authorids']['value']    # List of all authors id's
        Typ=s.content.get("type")
        Strm=s.content.get("stream")
        typ = int(Typ["value"][0]) if not Typ==None else 1 # Fallback for previous IWAI editions without Type.
        typs = types[typ-1]
        tit = s.content['title']['value']
        kws = s.content['keywords']['value']
        # auth_prof = openreview.tools.get_profiles(client,aid)
        # auth[s.number]=ais
        if typ==1:
            i_p+=1;   authP[s.number]=[i_p,f"#{subm_id}",Strm, tit, aid, kws]
        else:
            i_a+=1;   authA[s.number]=[i_a,f"#{subm_id}",Strm, tit, aid, kws]
        pass
    print(f" ------------- {i_p} FULL PAPERS -------- ")
    pprint.pprint(dict(sorted(authP.items())))
    print(f" ------------- {i_a} EXT ABSTRACTS -------- ")
    pprint.pprint(dict(sorted(authA.items())))
    print(f" ----- TOTAL {i_p+i_a}:  {i_p} full papes and {i_a} ext abstracts -----")

def submissions():
    os.makedirs(YEAR, exist_ok=True)
    xls_fname = f"{venue}-submission-abstracts.xlsx"
    xls_fpath = os.path.join(YEAR,xls_fname)
    data = {0:[],1:[]} # Submission types

    for s in submissions:
        ID = s.number
        typ = int(s.content.get("type")["value"][0]) - 1
        strm = int(s.content.get("stream")["value"][0]) - 1
        stream = streams[strm]
        authors = s.content.get("authors",[])["value"]
        if isinstance(authors,list): authors=", ".join(authors)
        has_pdf= "yes" if s.content.get("pdf") else "no"

        data[typ].append({
            "Stream": stream,
                "ID#": ID,
                "Title": s.content.get("title", "")["value"],
                "Authors": authors,
                "pdf": has_pdf,
                "Abstract": s.content.get("abstract", "")["value"]
            })
    df0 = pd.DataFrame(data[0], columns=["Stream", "ID#", "Title", "Authors", "pdf", "Abstract"]).sort_values(by=["Stream", "ID#"])
    df1 = pd.DataFrame(data[1], columns=["Stream", "ID#", "Title", "Authors", "pdf", "Abstract"]).sort_values(by=["Stream", "ID#"])
    with pd.ExcelWriter(xls_fpath, engine="openpyxl") as writer:
        df0.to_excel(writer, sheet_name=types[0], index=False)
        df1.to_excel(writer, sheet_name=types[1], index=False)
    print(f"Exported {len(df0)} full-paper submissions and {len(df1)} ext.abstr. submissions to {xls_fname}")

def download_pdf():
    for s in submissions:
        pdf_field=s.content.get("pdf")

        if pdf_field:
            pdf_path=pdf_field["value"] if isinstance(pdf_field, dict) else pdf_field

            title = s.content["title"]["value"]
            typ = int(s.content.get("type")["value"][0])-1
            type = types[typ]
            strm = int(s.content.get("stream")["value"][0])-1
            stream = streams[strm]

            fpath = os.path.join(YEAR, "pdf",f"{type}")
            os.makedirs(fpath, exist_ok=True)

            fname = f'{type}-ID{s.number}_{stream}_{title}.pdf'
            ffname=os.path.join(fpath,fname)

            if not os.path.exists(ffname):
                #f = client.get_attachment(s.id, pdf_path) # Does not work
                try:
                    f=client.get_pdf(s.id)
                    with open(ffname, 'wb') as op:
                        op.write(f)

                except Exception as e:
                    print(f"Failed {s.id}: {e}")

"""

TO DO: 
 1. messaging (to authors and potential authors)

"""

if __name__ == '__main__':
    #authors()  # List all author's IDs or emails.
    #monitor()               # List all submissions (type,title,autor-IDs, keywords)
    submissions()
    #download_pdf()  # Download submissions and store them in directories by type
