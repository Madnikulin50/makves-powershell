param (
    [string]$base = 'DC=acme,DC=local',
    [string]$server = 'acme.local',
    [string]$outfilename = 'export_ad',
    [string]$user = "current",
    [string]$pwd = "",
    [switch]$notping = $true,
    [string]$start = "",
    [string]$startfn = "", ##".ad-monitor.time_mark",
    [string]$makves_url = "",##"http://10.0.0.10:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [int]$timeout = 0
 )

Write-Host "base: " $base
Write-Host "server: " $server

Write-Host "user: " $user
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$env:Path += ";$scriptPath\MakvesActiveDirectory"

Import-Module -Name $scriptPath"\MakvesActiveDirectory" -Verbose

Test-ActiveDirectory -Base $base -Server $Server -Outfilename $outfilename `
 -User $user -Pwd $pwd `
 -NotPing $notping `
 -Start $start -StartFn $startfn `
 -Makves_url $makves_url -Makves_user $makves_user -Makves_pwd $makves_pwd

Write-Host "Export finished..."  -ForegroundColor Green
