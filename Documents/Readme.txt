Documents
=========

Purpose
-------
Scripts in this folder find and report on ProjectWise documents that need attention.
Use them for housekeeping such as stale check-outs, documents without files and other
document-level clean-up work.

Typical tasks
-------------
- List documents that have been checked out for longer than a set number of days.
- Find abstract documents and documents whose file is missing or empty.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module.
- Report scripts are read-only. Any script that frees, deletes or changes documents
  says so in its help and supports -WhatIf.
- Each script writes a timestamped .LastRun.log transcript beside itself.

See also
--------
- FileManipulation for importing, exporting and versioning files.
- States for moving documents between workflow states.
