import openreview
import config_ivo as c
import pandas as pd
import os
import time
import pprint

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

# CLIENT
client=openreview.api.OpenReviewClient(baseurl=BASE_URL,username=c.usr,password=c.pas)
venue_group = client.get_group(venue_id)
submission_str   = venue_group.content['submission_name']['value'] # this query results in text "Submission"

# SUBMISSION TYPES
streams = ['1-comp','2-cogn','3-appl']
types=['1-paper','2-abstr']

submissions = client.get_all_notes(invitation=SUBMISSION_INVITATION, sort='number:asc', details='replies')

def extract_field(content, key):
    """  OpenReview V2 stores values as:  content[key]["value"] or content[key]  """
    if key not in content:  return None
    value = content[key]
    if isinstance(value, dict) and "value" in value:
        return value["value"]
    return value

def authors(print_list=False):
    """ get all unique author IDs and either print the list or just return it """
    author_ids = set()
    for paper in submissions:
        ids = paper.content.get('authorids', [])["value"]
        for a_id in ids:
            if a_id and a_id != 'None':
                author_ids.add(a_id)
    if print_list:
        pprint.pprint(author_ids)
        print(f"Found {len(author_ids)} unique author IDs.")
    else:
        return author_ids

def authors_by_type(stype):
    author_ids = set()
    for paper in submissions:
        ids = paper.content.get('authorids', [])["value"]
        Typ=paper.content.get("type")
        typ = int(Typ["value"][0])
        if typ==stype:
            for a_id in ids:
                if a_id and a_id != 'None':
                    author_ids.add(a_id)
    pprint.pprint(author_ids)
    print(f"Found {len(author_ids)} unique author IDs of type {types[stype-1]}.")


def monitor():
    authP,authA = {},{}; i_p,i_a = 0,0
    for i,s in enumerate(submissions):
        subm_id = s.number
        Typ=s.content.get("type")
        strm=s.content.get("stream")
        Strm = streams[int(strm["value"][0])-1] if strm else ""

        typ = int(Typ["value"][0]) if not Typ==None else 1 # Fallback for previous IWAI editions without Type.
        typs = types[typ-1]
        tit = s.content['title']['value']
        kws = s.content['keywords']['value']
        has_pdf = "pdf" if s.content.get("pdf") else "   "
        authors = s.content.get("authors", [])["value"]
        authorsid = s.content.get("authorids", [])["value"]
        authid0 = authorsid[0]
        if authid0.startswith("~"):
            profile = client.get_profile(authid0)
            email = profile.content.get("preferredEmail")
        else:
            email = authid0
        email_domain=email[-9:]

        # auth_prof = openreview.tools.get_profiles(client,aid)
        # auth[s.number]=ais
        if typ==1:
            i_p+=1; authP[i_p]=[f"ID# {subm_id}", Strm, has_pdf, email_domain, authors[0], tit]
        else:
            i_a+=1; authA[i_a]=[f"ID# {subm_id}", Strm, has_pdf, email_domain, authors[0], tit]
        pass
    print(f" ------------- {i_p} FULL PAPERS -------- ")
    pprint.pprint(dict(sorted(authP.items())), width=300)
    print(f" ------------- {i_a} EXT ABSTRACTS -------- ")
    pprint.pprint(dict(sorted(authA.items())), width=300)
    print(f" ----- TOTAL {i_p+i_a}:  {i_p} full papes and {i_a} ext abstracts -----")

def submissions2xls():
    os.makedirs(YEAR, exist_ok=True)
    xls_fname = f"{VENUE}-submission-abstracts.xlsx"
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
        kwd = s.content['keywords']['value']

        data[typ].append({
            "Stream": stream,
                "ID#": ID,
                "Title": s.content.get("title", "")["value"],
                "Authors": authors,
                "pdf": has_pdf,
                "Keywords": kwd,
                "Abstract": s.content.get("abstract", "")["value"]
            })
    Columns_to_export=["Stream", "ID#", "Title", "Authors", "pdf", "Keywords", "Abstract"]
    df0 = pd.DataFrame(data[0], columns=Columns_to_export).sort_values(by=["Stream", "ID#"])
    df1 = pd.DataFrame(data[1], columns=Columns_to_export).sort_values(by=["Stream", "ID#"])
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

def write_to_myself():
    T=client.post_message(recipients=[c.usr], signature=c.usr, invitation=MSG_INVITATION,
                          subject='Test-OpenReview-Messaging', message="-- Test messaging --")
    print(T)

def write_to(TO,SBJ,MSG):
    try:
        T=client.post_message(recipients=[TO],signature=c.usr, invitation=MSG_INVITATION, subject=SBJ, message=MSG)
        print(T)
    except Exception as e:
        print(f"Failed sending message to {TO}: {e}")

def send_certificates_of_attendance():
    LIST = "xxxx.xlsx"
    SBJ = f"{VENUE} - Certificate of Attendance"
    MSG = """Dear {name},
    Thank you for participating in {VENUE}. Please download your Certificate of Attendance here: {link}
    Best regards,
    The IWAI 2025 Organizing Committee
    """
    df = pd.read_excel(LIST)
    for idx, row in df.iterrows():
        name = row.get("Name")
        email = row.get("Email")
        cert_path = row.get("Certificate")
        attachment = client.put_attachment(invitation=MSG_INVITATION, name=f"certificate_{name}",content=open(cert_path, 'rb'))
        link = attachment['url']
        time.sleep(0.5)
        msg = MSG.format(name=name, link=link, VENUE=VENUE)
        print(f"Try sending to {name} ({email})")
        client.post_message(recipients=[email], subject=SBJ, message=msg, invitation=MSG_INVITATION, signature=c.usr)
        time.sleep(0.5)

def send_invitation_to_contribute():
    # -- LIST = os.path.join("Lists","IWAI-Participants3.xlsx") -- MASTER LIST -- PAY ATTENTION WHEN USING IT. IWAI 2024 and 2025, both authors and actual participants
    # LIST = os.path.join("Lists","IWAI_Added.xlsx") # Non-participant Authors
    LIST = os.path.join("Lists","IWAI_Unmatched2.xlsx") # Non-participant Authors; to
    #LIST = os.path.join("2025","test-list.xlsx")
    import messages.invitation_to_contribute as msg
    df = pd.read_excel(LIST)
    sent=0
    for _, row in df.iterrows():
        name, email = row.get("Name","colleague"), row.get("Email")
        if not email:
            continue
        # This just verifies if the profile is present. The message is then sent to the profile name, not to the email (which is kept partially hidden)
        if email.startswith("~"):
            try:
                profile = client.get_profile(email)
                Email = profile.content.get("preferredEmail")
                print(f"{email} --> {Email}")
            except Exception as e:
                print(f"Cannot resolve {email}: {e}")
                continue

        # Send message to either email or profile-name (! profiles are case-sensitive !)
        try:
            client.post_message(recipients=[email.strip()], signature=c.usr, subject=msg.SBJ, message=msg.MSG.format(name=name), invitation=MSG_INVITATION)
            sent +=1
            time.sleep(0.5)
        except Exception as e:
            print(f"Failed for {email}: {e}")
    print(f"Message send to {sent}")


def send_message_to_authors():
    from Lists.authors2026 import AUTH_PAPER
    import messages.extension_and_anonymization as msg
    for auth in AUTH_PAPER:
        write_to(TO=auth,SBJ=msg.SBJ.format(VENUE=VENUE),MSG=msg.MSG.format(VENUE=VENUE))
        time.sleep(0.3)

def add_reviewers(reviewer_ids):
    """ Add reviewers to the venue's Reviewers group. reviewer_ids : list[str] of openreview profiles/emails """
    group = client.get_group(REVIEWER_GROUP)
    members = set(group.members)
    new_members = members.union(set(reviewer_ids))
    group.members = list(new_members)
    client.post_group_edit( invitation=f"{venue_id}/-/Edit", readers=[venue_id], writers=[venue_id], signatures=[venue_id], group=group)
    print(f"Reviewers group updated: {len(members)} -> {len(new_members)} members.")

def X_lossy_reviewers1():
    logs=client.get_process_logs(venue_id)
    for l in logs:
        print(l.id, l.status)
    expertise = client.get_all_expertise(venue_id)
    pass

def X_reviewers_expertise():
    #reviewers = client.get_group(REVIEWER_GROUP).members
    response = client.request_expertise(name='venue-reviewers-affinity',
                   group_id=REVIEWER_GROUP, venue_id=BL_SUBMISSION_INVITATION, model='specter2+scincl')
    job_id = response['jobId']
    print(f"Expertise job submitted successfully. Job ID: {job_id}")
    n=0
    while n<30:
        n +=1
        status_check = client.get_expertise_status(job_id=job_id)
        status = status_check.get('status','').lower()
        print(f"Current Job Status: {status}")
        if status == 'completed': break
        if status == 'error': raise RuntimeError("Expertise computation failed on OpenReview.")
        time.sleep(30)  # Wait 30 seconds before polling again

    results = client.get_expertise_results(job_id=job_id)
    rows=[]
    for reviewer, papers in results.get('results', {}).items():
        if isinstance(papers, dict):
            for paper, score in papers.items():
                rows.append({"reviewer": reviewer, "paper": paper, "score": score})
    df = pd.DataFrame(rows)
    df.to_excel('expertise_results.xlsx', index=False)
    print("Expertise exported to 'expertise_results.xlsx'")

def lossy_reviewers():
    reviewers = client.get_group(REVIEWER_GROUP).members
    no_pub = []
    for r in reviewers:
        try:
            profile = client.get_profile(r)
            pubs = profile.content.get("publications", [])
            dblp = profile.content.get("dblp", "") # also check DBLP / semantic scholar links if present
            s2 = profile.content.get("semanticScholarId", "")
            if (not pubs or len(pubs) == 0) and not dblp and not s2:
                no_pub.append(r)
        except Exception: no_pub.append(r)
    print(len(no_pub))

def compute_CoI():
    response = client.request_expertise(name='venue_conflicts', group_id=REVIEWER_GROUP,  venue_id=venue_id,  alternate_match_group=None,  model=None)
    job_id = response["jobId"]
    while True:
        status = client.get_expertise_status(job_id=job_id)
        if status["status"].lower() == "completed":   break
        if "error" in status["status"].lower():       raise RuntimeError(status)
        time.sleep(10)
    conflicts = client.get_expertise_results(job_id=job_id)



def all_invitaions():
    invitations = client.get_invitations(prefix=venue_id)
    for inv in invitations:
        if "Reviewer" in inv.id or "Conflict" in inv.id:
            print(inv.id)

if __name__ == '__main__':
    # -- IWAI submissions INFORMATION --
    # authors_by_type(2)  # List all author's IDs or emails.
    monitor()               # List all submissions (type,title,autor-IDs, keywords)
    submissions2xls()
    #download_pdf()  # Download submissions and store them in directories by type

    # ---------  MESSAGING  -----------
    #write_to_myself()
    #send_invitation_to_contribute()

    # import messages.remind_reviewers as msg
    # write_to(TO=msg.TO,SBJ=msg.SBJ.format(VENUE=VENUE),MSG=msg.MSG.format(VENUE=VENUE))

    #send_message_to_authors()

    # --- Reviewers ----
    #from Lists.IWAI2026_All_Reviewers import REVIEWERS
    #add_reviewers(REVIEWERS)
    #add_reviewers(authors())    # Add all authors as potential reviewers
    #authors(print_list=True)
    # lossy_reviewers()
    #get_CoI()
    #all_invitaions()

