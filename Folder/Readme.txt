Folder
======

Purpose
-------
Scripts in this folder walk, create and validate ProjectWise folder structures. Use them
as starting points for any script that has to visit every folder below a parent, and for
checking folder structures against Windows or a template.

Typical tasks
-------------
- Loop through folders with pwps_dab cmdlets or with SQL for speed.
- Compare a Windows folder tree with a ProjectWise folder tree.
- Create a standard folder structure from a CSV template.
- Report folders that contain no documents and no subfolders.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module.
- Folder creation scripts change the datasource and support -WhatIf.
- Each script writes a timestamped .LastRun.log transcript beside itself.
