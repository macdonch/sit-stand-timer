#
# Powershell Script: SitStandWrapper.ps1
#
# This script reads the settings.ini file and then invokes SitStandTimer.ps1
#
#
# INVOCATION FLAGS
#
# none
#
# Created by: charles.macdonald@telus.com
#

try {
    # Hide Window
    Add-Type -Name win -MemberDefinition '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);' -Namespace native
    [native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle,0)

    $InstallDir = split-path $PSCommandPath -Parent

    # Read settings.ini
    $settings_regex = '^(?!#)(?<property>\S.*\S)=(?<value>.*\S)$'
    $arrSettings = (Get-Content "$InstallDir\settings.ini" | Select-String -pattern $settings_regex)

    # initialize empty hash table
    $settings = @{}
    for ($i = 0; $i -lt $arrSettings.length; $i++) {
        if ($arrSettings[$i] -match $settings_regex) {
            $settings.Add($matches['property'].Trim(),$matches['value'])
        }
    }

    $StartPosition = $settings.StartPosition
    $SitMinutes = $settings.SitMinutes
    $StandMinutes = $settings.StandMinutes
    $RunHours = $settings.RunHours
        
    $command = "powershell.exe " + '"' + $InstallDir + '\SitStandTimer.ps1'
    $command += " -StartPosition $StartPosition"
    $command += " -SitMinutes $SitMinutes"
    $command += " -StandMinutes $StandMinutes"
    $command += " -RunHours $RunHours"
    $command += '"'

    Invoke-WmiMethod -Class Win32_Process -Name Create -ArgumentList $command

    [System.Diagnostics.Process]::GetCurrentProcess() | Get-Process | Stop-Process

    $exitcode = 0

}
catch {
    $exitcode = 1
}
finally {
    exit $exitcode
}
