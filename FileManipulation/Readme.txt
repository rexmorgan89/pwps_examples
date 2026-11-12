FileManipulation
================

Purpose
-------
Scripts in this folder move files into and out of ProjectWise and work with the files
themselves: importing from Windows, exporting with attributes, creating new versions,
digitally signing and finding duplicates or large files.

Typical tasks
-------------
- Import a Windows folder tree into ProjectWise.
- Export documents and their attributes to disk or to CSV/Excel.
- Create new versions from an Excel list.
- Find duplicate or unusually large files.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module. Some scripts also need
  ImportExcel.
- Import, versioning and signing scripts change ProjectWise data. Test in a
  non-production datasource first.
- Each script writes a timestamped .LastRun.log transcript beside itself.
