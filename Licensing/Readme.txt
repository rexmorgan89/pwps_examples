Licensing
=========

Purpose
-------
Scripts in this folder report on who uses ProjectWise and how often. Use them for licence
true-ups, finding unused accounts and planning capacity.

Typical tasks
-------------
- Report each user's last login and account status.
- Report which users are connected right now.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module. Excel output needs
  ImportExcel.
- Scripts here are read-only. Login data in ProjectWise is not a licence record; confirm
  licence counts in the Bentley licensing portal.
- Each script writes a timestamped .LastRun.log transcript beside itself.
