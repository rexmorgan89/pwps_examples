Tests
=====

Purpose
-------
Scripts in this folder check that a workstation, login and the core pwps_dab cmdlets work
before you run larger scripts. Use them after installing or upgrading ProjectWise or
pwps_dab, or when troubleshooting a script that fails.

Typical tasks
-------------
- Run a read-only smoke test of login and core cmdlets.
- Run a document round trip (create, check out, check in, delete) in a test folder.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module.
- Round-trip tests create and delete a document; point them at a test folder only.
- Each script writes a timestamped .LastRun.log transcript beside itself.
