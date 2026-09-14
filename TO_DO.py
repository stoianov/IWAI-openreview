"""

https://docs.openreview.net/venue-request-workflow/conferences

!! Run PostSubmission Stage to set the Submission-Reviewers'groups as readers of the Submissions

1+ verify the anonymization of all PDFs
2. Add abstracts from PDFs if missing; check incomplete submissions and inquire authors
3. Set all authors as reviewers
4. compute paper matching; run reviewer-assignment; etc. See PROCEDURE.md
5. check some special assignments.
6. upload assignments
7. notify reviewers

DECISIONS process

1. At the VenueConfiguration/DecisionStage -> set:

Decision Options:
Accept (Proceedings Paper + Oral Presentation), Accept (Proceedings Paper + Poster + Spotlight Presentation), Accept (Poster + Spotlight Presentation), Accept (Poster), Reject

Accept Decision Options:
Accept (Proceedings Paper + Oral Presentation), Accept (Proceedings Paper + Poster + Spotlight Presentation), Accept (Poster + Spotlight Presentation), Accept (Poster)

Decision Start Date and Deadline

Decisions not public; Release Decisions to Authors: YES; to Reviewers: NO

2. at Decision Stage

- upload a csv with the decisions (number; decision; note)

3. after that, at DecisionStage:

enable the post-decision stage by changing the date to an already past moment

4. at PostDecision stage:

- set Submission readers Program Chairs and Authors only
- Reveal Authors: No

- set the content of all Decision emails (one for each decision stage (see DecisionEmails.txt)
  be very careful with the text

- set Decision Notifications to "YES"

upon "Submit", this will send all decisions to the authors !! So make sure everything is correct.


"""



