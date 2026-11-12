SavedSearches
=============

Purpose
-------
Scripts in this folder work with ProjectWise saved searches: listing them, documenting
where they live and running them to produce reports without opening ProjectWise Explorer.

Typical tasks
-------------
- Export the list of saved searches, with their owner and location.
- Run a saved search and export the documents it returns to Excel.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module. Excel output needs
  ImportExcel.
- Scripts here are read-only.
- Each script writes a timestamped .LastRun.log transcript beside itself.
