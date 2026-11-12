Scheduled Tasks
===============

Purpose
-------
Scripts in this folder set up and monitor unattended jobs that run other scripts in this
library on a schedule, using Windows Task Scheduler or PowerShell scheduled jobs.

Typical tasks
-------------
- Register a scheduled job or task that runs a report script.
- Check the last run time and result of registered tasks.

Before running
--------------
- Requires Windows PowerShell 5.1 x64. Registering tasks that run as another account may
  need an elevated session.
- Unattended ProjectWise logins should use a credential file created with Export-Clixml
  by the account that runs the task. Never put passwords in the task or the script.
- Each script writes a timestamped .LastRun.log transcript beside itself.
