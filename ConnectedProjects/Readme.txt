ConnectedProjects
=================

Purpose
-------
Scripts in this folder create ProjectWise work areas (rich projects) and manage their
connection to Bentley Infrastructure Cloud projects (iTwins). Use them to provision
projects consistently and to report on connection status.

Typical tasks
-------------
- Create work areas from a GUI or from an input file.
- Report which work areas are connected to a cloud project and which are not.
- Create many work areas in one run from a CSV.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module.
- Creating work areas changes the datasource. Use -WhatIf where supported and test in a
  non-production datasource.
- Each script writes a timestamped .LastRun.log transcript beside itself.
