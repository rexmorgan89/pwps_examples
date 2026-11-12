DatasourceUtilities
===================

Purpose
-------
Scripts in this folder check the health, configuration and connectivity of ProjectWise
datasources. Use them to document a datasource, test logins, read network configuration
and compare datasources before migrations or upgrades.

Typical tasks
-------------
- Test login to every datasource visible on the workstation.
- Export datasource documentation (environments, workflows, groups, users).
- Take a health snapshot (server version, statistics, storage use).
- Compare users and groups between two datasources.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module.
- Most scripts here are read-only; check each script's help for exceptions.
- Each script writes a timestamped .LastRun.log transcript beside itself.
