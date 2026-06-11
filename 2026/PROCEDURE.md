# MANUAL

Description of activities to handle the openreview submission management.

Venue in October
Timeline relative to begining of process / -venue date)

## Request to open an openreview site
- to be done in November-December (Month +1/-11)
- essentially copy-past the request from the previous year
- attention to set correct days and emails of chairs
- abstract registration deadline is 2 weeks before submission deadline

## Communication to authors of previous years submissions
- extract authors (python); join lists;
- submit message (python)

## Start of submission period including preliminary abstract registration
- submit reminder 2 months and 1 month to all previous authors, before abstract registration deadline

## Switch to "Post Submission" stage to enable revision of registrations
To enable revision of submissions (authors, add pdfs etc):
- press the "Post Submission" button in OpenReview; select "Chairs and Authors" as readers
https://docs.openreview.net/getting-started/hosting-a-venue-on-openreview/enabling-an-abstract-registration-deadline?utm_source=chatgpt.com

## Review Stage 
(4 June 2026)

## Populate the Reviewers List
Reviewer's list is visible here:
In Program Chairs Console, bottom, left, Venue Roles -> Reviewers

Populate it automatically with either
- Selected reviewers:
  reviewer_selectory.py:
    get_submissions()
    extract_reviewers()
    save the string REVIEWERS=[....] into a file in directory Lists. Then import it and run add_reviewers(REVIEWERS)
- OR all authors potential reviewers: add_revuewers(authors())

extract_reviewers() gets the 1st, 2nd, and last (senior) author as reviewers. 
In case it is needed, pen-ultimate could be used as another senior reviewer. (update the script)
    
iwai.py:
    from Lists.IWAI2026_All_Reviewers import REVIEWERS
    add_reviewers(REVIEWERS)

## Compute Paper Matching (blue button on bottom right)
(9 June 2026)
using Comprehensive Conflict computation and Specter2+SciIncl affinity score method; the rest are empty

Then receive email:
Affinity scores and/or conflicts were successfully computed. 
111 Reviewers listed under 'Without Publication' don't have any publications.

To run the matcher, click on the 'Reviewers Paper Assignment' link in the Program Chairs Console: 
https://openreview.net/group?id=IWAI/2026/Workshop/Program_Chairs
(on the Time Line; BOTTOM RIGHT)


invitations = client.get_invitations(prefix=venue_id)
IWAI/2026/Workshop/-/Reviewer
IWAI/2026/Workshop/Reviewers/-/Affinity_Score
IWAI/2026/Workshop/Reviewers/-/Assignment
IWAI/2026/Workshop/Reviewers/-/Conflict
IWAI/2026/Workshop/Reviewers/-/Custom_Max_Papers
IWAI/2026/Workshop/Reviewers/-/Custom_User_Demands
IWAI/2026/Workshop/Reviewers/-/Message
IWAI/2026/Workshop/Reviewers/-/Review_Assignment_Count
IWAI/2026/Workshop/Reviewers/-/Review_Count
IWAI/2026/Workshop/Reviewers/-/Review_Days_Late_Sum
IWAI/2026/Workshop/Reviewers/-/Submission_Group
IWAI/2026/Workshop/Reviewers/-/Submission_Message
