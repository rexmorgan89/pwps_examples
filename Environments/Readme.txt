Environments
============

Purpose
-------
Scripts in this folder document and compare ProjectWise environment definitions
(tables, columns and interfaces). Use them before changing an environment, when
standardizing environments across datasources or when cleaning up unused ones.

Typical tasks
-------------
- Compare environments within one datasource or between two datasources.
- Export environment columns and interfaces to Excel.
- Report which folders use each environment.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module. Excel output needs
  ImportExcel.
- Scripts here are read-only.
- Each script writes a timestamped .LastRun.log transcript beside itself.
