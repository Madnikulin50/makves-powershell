
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")] $computers = ("acme.local"),
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $outfilename = "events",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $Count = 3000,
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $user = "current",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $pwd = "",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $start = "",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $fwd = "",
    [Parameter(Mandatory = $False, Position = 10, ParameterSetName = "NormalRun")] [ValidateSet("All","Logon","Service","User","Computer", "Clean", "File", "MSSQL", "RAS", "USB", "Printer", "Sysmon", "TS")] [array]$target="All",
    [string]$startfn = "", ##".event-monitor.time_mark",
    [string]$makves_url = "", ##"http://10.0.0.10:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$exclude_user = ""
)

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$env:Path += ";$scriptPath\MakvesLogs"
Write-Host $env:Path

Import-Module -Name $scriptPath"\MakvesLogs" -Verbose

Test-EventLog -Computers $computers -Target $target -Outfilename $outfilename -User $user `
 -Pwd $pwd -Fwd $fwd -Exclude_user $exclude_user -Count $count `
 -Start $start -StartFn $startfn `
 -Makves_url $makves_url -Makves_user $makves_user -Makves_pwd $makves_pwd
