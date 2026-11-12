AuditTrail
==========

Purpose
-------
Scripts in this folder read the ProjectWise audit trail to answer "who did what, and when".
Use them to export document history, summarize user activity and support investigations
or usage reviews.

Typical tasks
-------------
- Export audit trail records for a folder, document or date range to Excel.
- Summarize activity counts per user and per action.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module.
- Audit trail queries can be slow on large datasources; narrow the date range first.
- Scripts here are read-only.
- Each script writes a timestamped .LastRun.log transcript beside itself.
