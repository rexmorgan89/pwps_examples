Reports
=======

Purpose
-------
Datasource-wide reports for administrators: inactive users and work area activity.
Use them for clean-up planning and licence reviews.

Typical tasks
-------------
- Report when each work area was last active.
- List users who have not logged in for a set number of days.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module. Excel output needs
  ImportExcel.
- Scripts here are read-only but can run for a long time on large datasources.
  Run them outside working hours where possible.
- Each script writes a timestamped .LastRun.log transcript beside itself.
