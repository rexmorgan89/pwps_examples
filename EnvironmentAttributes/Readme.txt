EnvironmentAttributes
=====================

Purpose
-------
Scripts in this folder read and update the attribute values stored in ProjectWise
environments, and adjust environment column definitions.

Typical tasks
-------------
- Export the attribute values of documents in a folder to Excel.
- Update attribute values in bulk from an edited Excel sheet.
- Change environment column widths.
- Copy geospatial location from a folder to its documents.

Before running
--------------
- Requires Windows PowerShell 5.1 x64 and the pwps_dab module. Excel scripts also need
  ImportExcel.
- Attribute updates change document metadata. Use -WhatIf where supported and test in a
  non-production datasource.
- Each script writes a timestamped .LastRun.log transcript beside itself.
