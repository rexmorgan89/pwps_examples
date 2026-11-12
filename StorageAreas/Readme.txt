StorageAreas
============

Purpose
-------
Scripts in this folder report on ProjectWise storage areas and the documents stored in
them. Use them to plan storage migrations, balance disk use and find which work areas
live on which storage area.

Typical tasks
-------------
- Summarize document count and file size per storage area.
- List the documents and folders that use a given storage area.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module. Excel output needs
  ImportExcel.
- Scripts here are read-only.
- Each script writes a timestamped .LastRun.log transcript beside itself.
