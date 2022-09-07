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
## Detailed information (Japanese)
https://qiita.com/shokun0803@github/items/c5a25af344af49fa7fcf#windows-%E3%81%AE%E3%82%B5%E3%82%A4%E3%83%B3%E3%82%A4%E3%83%B3%E3%83%AD%E3%82%B0%E3%82%92%E5%8B%A4%E6%80%A0%E3%81%AB%E6%B4%BB%E7%94%A8%E3%81%97-sharepoint-%E3%81%AB%E3%83%AD%E3%82%B0%E4%BF%9D%E5%AD%98%E3%81%99%E3%82%8B
