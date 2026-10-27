# IWAI OpenReview python code

Author: Ivilin Stoianov

An official repository of python code used to access IWAI @ openreview. 

Its purpose is to support the Organization Committee in handling the submission and reviewing processes.

v1 15 May 2026 (authors, monitor, submissions2xls, pdf's)
regular updates with functionalities

Instructions:
- set *config.py* with your openreview account info. Be careful not to share your access information ..
- install the **openreview** library
- uncomment functions to be executed

## Scripts:

### iwai.py
The main code. List submissions, monitor, etc

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

