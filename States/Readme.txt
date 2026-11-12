States
======

Purpose
-------
Scripts in this folder report on and change the workflow states of ProjectWise documents.
Use them to see where documents sit in a workflow and to move many documents between
states in one run.

Typical tasks
-------------
- Count documents per workflow state, per folder.
- Change document states for a selection of documents or from a CSV list.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module.
- State changes can trigger workflow rules and notifications. Use -WhatIf where
  supported and test in a non-production datasource.
- Each script writes a timestamped .LastRun.log transcript beside itself.
