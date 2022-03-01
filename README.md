# Sit Stand Timer
This set of powershell scripts for standing desk users creates a pop-up dialog to remind you to change between sitting and standing.  

By default, sitting is set for 45 minutes, and standing for 90. If the position interval is 30 minutes or more, a balloon in the task bar give a reminder 15 minutes before the change. When the time interval has expired, a popup dialog reminds you to change position. After you click 'OK', the timer starts for the next position. 'Quit' causes the script to exit. The script runs for 8 hours after it is invoked. (Note: it will not stop running while a dialog window is open.)  

### Prerequisites
Sit Stand Timer was tested on:
- Windows 10 Enterprise
- Powershell 5.1.x

### Installation
The installer creates 3 files in the installation directory 'C:\SIS\SitStandTimer\:
- SitStandWrapper.ps1: this script reads the settings.ini file and invokes the SitStandTimer.ps1 script
- SitStandTimer.ps1: this script is the application
- settings.ini: this file contains initialization parameters

The installer also creates a shortcut on the desktop. The shortcut launches SitStandWrapper.ps1.

### Configuration
The file 'settings.ini' controls the script behaviour.

```
# 
# SitStandTimer Settings
#
# StartPosition: accepts string values "Sit" or "Stand"
# SitMinutes: accepts integer values
# StandMinutes: accepts integer values
# RunHours: accepts integer values
#
# DEFAULT VALUES
# [General]
# StartPosition="Sit"
# SitMinutes=45
# StandMinutes=90
# RunHours=8
#

StartPosition="Sit"
SitMinutes=45
StandMinutes=90
RunHours=8
```

## Built with
- Powershell
- [Advanced Installer](https://www.advancedinstaller.com/) - used to create the msi

## Authors
- [Charles Macdonald](mailto:charles.macdonald@telus.com)