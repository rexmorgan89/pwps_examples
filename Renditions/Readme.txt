Renditions
==========

Purpose
-------
Scripts in this folder submit documents to ProjectWise rendition profiles (for example
to produce PDFs from plot sets) and report on how rendition profiles are configured
and assigned.

Typical tasks
-------------
- Submit the documents in a folder, or matching a search, to a rendition profile.
- Report rendition profiles and the folders and work areas they are assigned to.

Before running
--------------
- Requires Windows PowerShell 5.1 x64, the pwps_dab module and a working rendition
  service. Plot set (IPS) renditions may need the Plot IPS Files add-in
  (see PlotIPSFilesSetup20210623).
- Submitting requests can create or replace output files. Test with a small folder first.
- Each script writes a timestamped .LastRun.log transcript beside itself.
