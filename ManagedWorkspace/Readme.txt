ManagedWorkspace
================

Purpose
-------
Scripts in this folder report on ProjectWise Managed Workspace configuration:
Configuration Settings Blocks (CSBs), their variables and where they are assigned.
Use them to document a workspace, review it before changes and find unused blocks.

Typical tasks
-------------
- Compare the Managed Workspace assignments of folders.
- Export CSB assignments for folders and work areas.
- Find CSBs that are not assigned anywhere.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module.
- Scripts here are read-only.
- Each script writes a timestamped .LastRun.log transcript beside itself.
