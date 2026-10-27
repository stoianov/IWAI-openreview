# IWAI OpenReview python code

### Author: Ivilin Stoianov

This is an official repository of python code used to access the submissions to the IWAI conference. 

Its purpose is to support the Organization Committee in handling the submission and reviewing processes.

v0: 15 May 2026 (authors, monitor, submissions2xls, pdf's)
v1,2,3 ..

Instructions:
- set the *config.py* with your openreview account info. Be careful not to share your access information ..
- install **openreview** library (e.g. with pip)
- uncomment function to be executed

## Scripts:

### iwai.py
This is the main code. List submissions, monitor, etc

    - authors()
List all author's IDs or emails.

    - monitor()             
List all submissions (type,title,autor-IDs, keywords)

    - submissions2xls()
Export all submissions to an xls file, in separate sheets and ordered by stream

    - download_pdf()
Download submissions and store them in directories by type

    - write_to()
Send a message to a list of emails/openreview-accounts. Se example in main.py


### reviewers.py

Handles reviewer-paper match etc

### cameraready.py

Handles camera-ready

    -  download_cameraready_parallel(client, submissions, YEAR)
export in xlsx the cameraready status of all accepted-for-publication papers
and download all pdf/zip files

