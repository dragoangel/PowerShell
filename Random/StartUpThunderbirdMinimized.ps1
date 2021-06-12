param (
    [Parameter(Mandatory=$false)]
    [Switch]$Install,
    [Parameter(Mandatory=$false)]
    [Switch]$Uninstall
)

if ($Install){
    if (-not (Test-Path -Path $MyInvocation.MyCommand.Definition)){
        Write-Warning "Installation of scheduled job failed. You should run this script directly from Powershell terminal to get it work."
        Pause
    } else {
        $VbsContent = @"
command = `"PowerShell.exe -NoLogo -File `"`"$($MyInvocation.MyCommand.Definition)`"`"`"
set shell = CreateObject("WScript.Shell")
shell.Run command,0
"@
        $VbsPath = $MyInvocation.MyCommand.Definition -Replace "\.ps1$",".vbs"
        $VbsContent | Out-File $VbsPath
        $Action = New-ScheduledTaskAction -Execute “wscript” -Argument "//nologo `"$VbsPath`""
        $Principal = New-ScheduledTaskPrincipal -UserId (Get-CimInstance –ClassName Win32_ComputerSystem | Select-Object -Expand UserName)
        $Trigger = New-ScheduledTaskTrigger -AtLogOn -User (Get-CimInstance –ClassName Win32_ComputerSystem | Select-Object -Expand UserName)
        $SettingsParams = @{
            "ExecutionTimeLimit"         = (New-TimeSpan -Minutes 2)
            "AllowStartIfOnBatteries"    = $True
            "DontStopIfGoingOnBatteries" = $True
            "RestartCount"               = 0
        }
        $Settings = New-ScheduledTaskSettingsSet @SettingsParams
        $TaskParams = @{
            "Action"    = $Action
            "Principal" = $Principal
            "Trigger"   = $Trigger
            "Setting"   = $Settings
        }
        $Task = New-ScheduledTask @TaskParams
        Get-ScheduledTask -TaskName "Thunderbird_StartUp_${env:USERNAME}" -ErrorAction Ignore | Unregister-ScheduledTask -Confirm:$False
        $Task | Register-ScheduledTask -TaskName "Thunderbird_StartUp_${env:USERNAME}"
        Exit 0
    }
}

if ($Uninstall){
    Get-ScheduledTask -TaskName "Thunderbird_StartUp_${env:USERNAME}" -ErrorAction Continue | Unregister-ScheduledTask -Confirm:$False
    Exit 0
}

$ErrorActionPreference = "SilentlyContinue"
$ThunderbirdEML = Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\ThunderbirdEML\shell\open\command"
if ($ThunderbirdEML.'(default)'.ToLower().Contains("\thunderbird.exe")){
    $ThunderbirdExecutable = (Split-path -Parent $ThunderbirdEML.'(default)'.Replace('"','')) + "\thunderbird.exe"
    if (-not (Test-Path -Path $ThunderbirdExecutable)){Exit 1}
} else {
    if (Test-Path -Path "${env:ProgramFiles}\Mozilla Thunderbird\thunderbird.exe"){$ThunderbirdExecutable = "${env:ProgramFiles}\Mozilla Thunderbird\thunderbird.exe"}
    if (Test-Path -Path "${env:ProgramFiles(x86)}\Mozilla Thunderbird\thunderbird.exe"){$ThunderbirdExecutable = "${env:ProgramFiles(x86)}\Mozilla Thunderbird\thunderbird.exe"}
    if ($ThunderbirdExecutable -eq $null){Exit 1}
}

function Get-ProcessState {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [String]$Name
    )
    begin {
        Add-Type -AssemblyName UIAutomationClient
    }
    process {
        $ProcessList = Get-Process -Name $Name
        foreach ($Process in $ProcessList) {
            $AutomationElement = [System.Windows.Automation.AutomationElement]::FromHandle($Process.MainWindowHandle)
            $ProcessPattern = $AutomationElement.GetCurrentPattern([System.Windows.Automation.WindowPatternIdentifiers]::Pattern)
            [PSCustomObject]@{
                Process = $Process.MainWindowTitle
                ProcessState = $ProcessPattern.Current.WindowVisualState
            }
        }
    }
}

function Set-WindowStyle {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [String]$Name,
        [Parameter(Mandatory=$true)]
        [ValidateSet('FORCEMINIMIZE', 'HIDE', 'MAXIMIZE', 'MINIMIZE', 'RESTORE', 
                     'SHOW', 'SHOWDEFAULT', 'SHOWMAXIMIZED', 'SHOWMINIMIZED', 
                     'SHOWMINNOACTIVE', 'SHOWNA', 'SHOWNOACTIVATE', 'SHOWNORMAL')]
        $Style = 'SHOW'
    )
    $WindowStates = @{
        'FORCEMINIMIZE'   = 11
        'HIDE'            = 0
        'MAXIMIZE'        = 3
        'MINIMIZE'        = 6
        'RESTORE'         = 9
        'SHOW'            = 5
        'SHOWDEFAULT'     = 10
        'SHOWMAXIMIZED'   = 3
        'SHOWMINIMIZED'   = 2
        'SHOWMINNOACTIVE' = 7
        'SHOWNA'          = 8
        'SHOWNOACTIVATE'  = 4
        'SHOWNORMAL'      = 1
    }
    $MemberDefinition = @'
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
'@
    $Win32ShowWindowAsync = Add-Type -MemberDefinition $MemberDefinition -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
    $ProcessList = Get-Process -Name $Name
    foreach ($Process in $ProcessList) {
        $Win32ShowWindowAsync::ShowWindowAsync($Process.MainWindowHandle, $WindowStates[$Style]) | Out-Null
    }
}

$ThunderbirdProcess = Get-Process -Name thunderbird
if ($ThunderbirdProcess -eq $null){Start-Process -FilePath $ThunderbirdExecutable}
do {$ThunderbirdProcessState = Get-ProcessState -Name thunderbird} while ($ThunderbirdProcessState.ProcessState -eq $null)
if ($ThunderbirdProcessState.ProcessState -ne "Minimized"){
    do {
        Set-WindowStyle -Name thunderbird -Style MINIMIZE
        $ThunderbirdProcessState = Get-ProcessState -Name thunderbird
    } while ($ThunderbirdProcessState.ProcessState -ne "Minimized")
}
