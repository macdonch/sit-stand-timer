param (
    [string]$StartPosition,
    [int]$SitMinutes,
    [int]$StandMinutes,
    [int]$RunHours
)

#
# Powershell Script: SitStandWrapper.ps1
#
# This script for standing desk users creates a pop-up dialog reminder
# to change between sitting and standing.
#
#
# INVOCATION FLAGS
#
# -StartPosition (accepts string values "Sit" or "Stand")
# -SitMinutes (integer)
# -StandMinutes (integer)
# -RunHours (integer)
#
# Created by: charles.macdonald@telus.com
#

# FUNCTIONS
function Format-InvocationParamsAsString
{
    param (
        $Params
    )

    begin  {
    }

    process {
        try {
            $strParams = ''
            foreach($key in $Params.Keys) {
                # add a space if it isn't the first key
                if ($strParams -ne '') {
                    $strParams += ' '
                }
                if ($Params.$key -Is [string]) {
                    $strParams += "-$key '$($Params.$key)'"
                } else {
                    $strParams += "-$key $($Params.$key)"
                }
            }
        }
            
        catch {
            # on error return null
            $strParams = $null
        }
		return $strParams
    }

    End {
    }
}

function Show-BalloonTip {            
    [cmdletbinding()]            
    param(            
        [parameter(Mandatory=$true)]            
        [string]$Title,            
        [ValidateSet("Info","Warning","Error")]             
        [string]$MessageType = "Info",            
        [parameter(Mandatory=$true)]            
        [string]$Message,            
        [string]$Duration=10000            
    )            
    
    [system.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null            
    $balloon = New-Object System.Windows.Forms.NotifyIcon            
    $path = Get-Process -id $pid | Select-Object -ExpandProperty Path            
    $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)            
    $balloon.Icon = $icon            
    $balloon.BalloonTipIcon = $MessageType            
    $balloon.BalloonTipText = $Message            
    $balloon.BalloonTipTitle = $Title            
    $balloon.Visible = $true
    $balloon.ShowBalloonTip($Duration)
    
    $balloon.Dispose()
    
}

function ClearAndClose()
{
   $Form.Dispose();
   if ($Timer) {
     $Timer.Dispose();
     $Script:CountDown=5
   }
}

function Button_Click()
{
   ClearAndClose
}

function ButtonQuit_Click()
{
   ClearAndClose
}

function Timer_Tick()
{
    --$Script:CountDown
    if ($Script:CountDown -lt 0)
    {
        ClearAndClose
    }
}

function WelcomeForm
{
    param (
        [Parameter(Mandatory)]
        [string]$Message
    )
   
    # Create Welcome Form
   
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $Form = New-Object system.Windows.Forms.Form
    $Form.Text = "Sit Stand Timer"
    $Form.Size = New-Object System.Drawing.Size(500,170)
    $Form.StartPosition = "CenterScreen"
    $Form.Topmost = $True
    # check icon
    $InstallDir = split-path $PSCommandPath -Parent
    $iconpath = "$InstallDir\icons\SitStandIcon.ico"
    if (Test-Path $iconpath -PathType leaf) {
        $Form.Icon = New-Object system.drawing.icon($iconpath)
    }

    $Label = New-Object System.Windows.Forms.Label
    $Label.AutoSize = $true
    $Label.Location = New-Object System.Drawing.Size(20,10)
    $Label.Text = $Message
   
    $Button = New-Object System.Windows.Forms.Button
    $Button.Location = New-Object System.Drawing.Size(190,100)
    $Button.Size = New-Object System.Drawing.Size(120,23)
    $Button.Text = "OK"
    $Button.DialogResult=[System.Windows.Forms.DialogResult]::OK
   
    $Timer = New-Object System.Windows.Forms.Timer
    $Timer.Interval = 1000
   
    $Form.Controls.Add($Label)
    $Form.Controls.Add($Button)
   
    $Script:CountDown = 30
   
    $Button.Add_Click({Button_Click})
    $Timer.Add_Tick({ Timer_Tick})
   
    $Timer.Start()
    $Form.ShowDialog()
}
function PositionForm
{
    param (
        [Parameter(Mandatory)]
        [string]$NewPosition
    )
  
    # Create Position Form
   
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $Form = New-Object system.Windows.Forms.Form
    $Form.Text = "Time to $NewPosition"
    $Form.Size = New-Object System.Drawing.Size(500,170)
    $Form.StartPosition = "CenterScreen"
    $Form.Topmost = $True
    # check icon
    $InstallDir = split-path $PSCommandPath -Parent
    $iconpath = "$InstallDir\icons\SitStandIcon.ico"
    if (Test-Path $iconpath -PathType leaf) {
        $Form.Icon = New-Object system.drawing.icon($iconpath)
    }

    $Label = New-Object System.Windows.Forms.Label
    $Label.AutoSize = $true
    $Label.Location = New-Object System.Drawing.Size(20,10)
    $Label.Text = "Click 'OK' after changing positions, or 'Quit' to exit"
   
    $Button = New-Object System.Windows.Forms.Button
    $Button.Location = New-Object System.Drawing.Size(110,100)
    $Button.Size = New-Object System.Drawing.Size(120,23)
    $Button.Text = "OK"
    $Button.DialogResult=[System.Windows.Forms.DialogResult]::OK

    $ButtonQuit = New-Object System.Windows.Forms.Button
    $ButtonQuit.Location = New-Object System.Drawing.Size(270,100)
    $ButtonQuit.Size = New-Object System.Drawing.Size(120,23)
    $ButtonQuit.Text = "Quit"
    $ButtonQuit.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
   
    $Form.Controls.Add($Label)
    $Form.Controls.Add($Button)
    $Form.Controls.Add($ButtonQuit)
   
    $Button.Add_Click({Button_Click})
    $ButtonQuit.Add_Click({ButtonQuit_Click})

    $Form.ShowDialog()
}

function KillExistingInstanceDialog
{
 
    # Create Form

    $Message = "A previous instance of Sit Stand Timer is still running."
    $Message += "`n`nClick 'OK' to end the existing process, or 'Quit' to exit"
   
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $Form = New-Object system.Windows.Forms.Form
    $Form.Text = "SitStandTimer Already Running"
    $Form.Size = New-Object System.Drawing.Size(500,170)
    $Form.StartPosition = "CenterScreen"
    $Form.Topmost = $True
    # check icon
    $InstallDir = split-path $PSCommandPath -Parent
    $iconpath = "$InstallDir\icons\SitStandIcon.ico"
    if (Test-Path $iconpath -PathType leaf) {
        $Form.Icon = New-Object system.drawing.icon($iconpath)
    }

    $Label = New-Object System.Windows.Forms.Label
    $Label.AutoSize = $true
    $Label.Location = New-Object System.Drawing.Size(20,10)
    $Label.Text = $Message
   
    $Button = New-Object System.Windows.Forms.Button
    $Button.Location = New-Object System.Drawing.Size(110,100)
    $Button.Size = New-Object System.Drawing.Size(120,23)
    $Button.Text = "OK"
    $Button.DialogResult=[System.Windows.Forms.DialogResult]::OK

    $ButtonQuit = New-Object System.Windows.Forms.Button
    $ButtonQuit.Location = New-Object System.Drawing.Size(270,100)
    $ButtonQuit.Size = New-Object System.Drawing.Size(120,23)
    $ButtonQuit.Text = "Quit"
    $ButtonQuit.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
   
    $Form.Controls.Add($Label)
    $Form.Controls.Add($Button)
    $Form.Controls.Add($ButtonQuit)
   
    $Button.Add_Click({Button_Click})
    $ButtonQuit.Add_Click({ButtonQuit_Click})

    $Form.ShowDialog()
}

# BODY

# Preference Settings
# change the debug preference to Continue for verbose screen debug output
$DebugPreference = "SilentlyContinue"
# $DebugPreference = 'Continue'
$ErrorActionPreference = 'Continue'

try {

    # Hide Window
    Add-Type -Name win -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);' -Namespace native
    [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle,0)

    # Look for running instances
    $existingProcesses = (Get-WMIObject -Class Win32_Process -Filter "Name='PowerShell.EXE'" | `
        Where-Object {$_.CommandLine -Like "*SitStandTimer.ps1*"} | Select-Object Handle,CommandLine)

    # Set boolContinue
    $boolContinue = $true

    if (($existingProcesses | Measure-Object).Count -gt 1) {
        # ask if existing process should be killed
        $killExisting = KillExistingInstanceDialog
        if ($killExisting -like '*OK') {
            foreach ($existingProcess in $existingProcesses) {
                if ($existingProcess.handle -ne $PID) {
                    Stop-Process -Id $existingProcess.handle | Out-Null
                }
            }
        } else {
            $boolContinue = $false
        }
    }

    if ($boolContinue) {
        # set variables
        $startTime = Get-Date
        $position = $StartPosition

        # Create Welcome Dialogue Message
        $StartMessage = 'Welcome!'
        if ($position -like 'Stand*') {
            $StartMessage += "`n`nSitStandTimer will remind you in $StandMinutes minutes to change positions from standing to sitting."
        } else {
            $StartMessage += "`n`nSitStandTimer will remind you in $SitMinutes minutes to change positions from sitting to standing."
        }
        $StartMessage += "`nSitStandTimer will run for $RunHours hours."
        $StartMessage += "`n`nThis window will close in 30 seconds."

        # Create Welcome Form
        (WelcomeForm -Message $StartMessage) | Out-Null

        Write-Debug $position

        # run for $RunHours
        while ((New-Timespan -Start (Get-Date) -End $startTime.AddHours($RunHours)).TotalHours -gt 0) {
            if ($position -like 'Stand*') {
                $position = 'Stand'
                $positionMinutes = $StandMinutes
                $nextPosition = 'Sit'
            } else {
                $position = 'Sit'
                $positionMinutes = $SitMinutes
                $nextPosition = 'Stand'
            }

            Write-Debug $nextPosition
            Write-Debug $positionMinutes

            # set sleep and notifications
            if ($positionMinutes -ge 30) {
                $timeExpiryWarning = 15
                Write-Debug "Sleeping for $($positionMinutes - $timeExpiryWarning) minutes"
                Start-Sleep (($positionMinutes - $timeExpiryWarning) * 60)
                # create popup warning
                Show-BalloonTip -Title "Change Positions Soon" -Message "Time to $nextPosition in $timeExpiryWarning minutes" -Duration 30000
                Write-Debug "Sleeping for $timeExpiryWarning minutes"
                Start-Sleep ($timeExpiryWarning * 60)
            } else {
                Start-Sleep ($positionMinutes * 60)
            }
            # create popups for position change
            Show-BalloonTip -Title "Change Positions Now" -Message "Time to $nextPosition" -Duration 30000
            $resultPositionForm = PositionForm -NewPosition $nextPosition

            if ($resultPositionForm -match 'Cancel') {
                break
            }

            # change position
            $position = $nextPosition
        }

        Show-BalloonTip -Title "SitStandTimer Complete" -Message "$PsCommandPath has exited normally." -Duration 30000

    }

    $exitcode = 0
}

catch {
    try {
        [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle,1)
        Write-Host $_.Exception
    }
    catch {}
    try {
        Show-BalloonTip -Title "SitStandTimer Failed" -Message "$PsCommandPath has exited abnormally with $($_.Exception.Mesage)." -Duration 30000
    }
    catch {}
    $exitcode = 1
}

finally {
    exit $exitcode
}