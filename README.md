# getLog_script
Get Windows sign-in information with PowerShell and save logs to SharePoint via OneDrive.

## prerequisite
* Operating as a file server using a document library on SharePoint
* Creating a shortcut to a SharePoint library in OneDrive for Business
* The SharePoint directory where the script stores logs must be accessible from the device running the script.

## Directory structure
* local
```
C:\pcmaint -  getLog_task.vbs
           |- getLog_task.bat
```
* SharePoint
```
\document library - pcmaint -  getLog_task.vbs
                            |- getLog_task.bat
                            |- check_holiday.ps1
                            |- getLog_script.ps1
                            |- getLog_taskschedule_set.ps1
                            |- system_chk.ps1
                  |- timecard_log - path_chk.txt
                  |- system_log
```
## execution procedure
Edit $serverPath = "" in each script.  
Run the getLog_taskschedule_set.ps1 file on each user's terminal.
