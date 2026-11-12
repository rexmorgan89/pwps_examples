AccessControl
=============

Purpose
-------
Scripts in this folder report on, repair and replace ProjectWise folder and document
security (access control lists). Use them to audit who can reach which folders, find
folders where explicit security breaks inheritance, and rebuild or replace security in bulk.

Typical tasks
-------------
- Export folder security to Excel for review or audits.
- Count access control entries per project or work area.
- Find folders whose security no longer inherits from the parent.
- Replace or reset folder security for a group, user list or user.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module.
- Scripts that change security can lock users out of folders. Run report scripts first,
  use -WhatIf where supported and test in a non-production datasource.
- Each script writes a timestamped .LastRun.log transcript beside itself.
