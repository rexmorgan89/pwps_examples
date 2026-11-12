Groups&UserLists
================

Purpose
-------
Scripts in this folder manage ProjectWise groups and user lists and their members.
Use them to report membership and keep membership in line with a source list.

Typical tasks
-------------
- Export every group and user list with its members to Excel.
- Add and remove members from a CSV.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module. Excel output needs
  ImportExcel.
- Membership changes affect what users can access. Use -WhatIf where supported and test
  in a non-production datasource.
- Each script writes a timestamped .LastRun.log transcript beside itself.
