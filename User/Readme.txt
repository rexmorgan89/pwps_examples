User
====

Purpose
-------
Scripts in this folder manage ProjectWise user accounts: reporting user properties and
disabling accounts.

Typical tasks
-------------
- Export users with their type, email, last login and disabled flag.
- Disable users listed in a CSV.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module.
- Changes to users affect sign-in. Use -WhatIf where supported and test in a
  non-production datasource.
- Each script writes a timestamped .LastRun.log transcript beside itself.
