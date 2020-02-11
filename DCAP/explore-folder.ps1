param (
    [string]$folder = "C:\work",
    [string]$outfilename = "folder", ##"",
    [string]$base = "",
    [string]$server = "",
    [int]$hashlen = 1048576,
    [switch]$no_hash = $false,
    [switch]$extruct = $false,
    [switch]$compliance = $false,
    [switch]$monitor = $false,
    [string]$start = "",
    [string]$startfn = "", ##".file-monitor.time_mark",
    [string]$makves_url =  "", ##"http://192.168.2.22:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$logfilename = "log"
)

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$env:Path += ";$scriptPath\MakvesFileSystem"
Write-Host $env:Path

Import-Module -Name $scriptPath"\MakvesFileSystem" -Verbose

Test-FileSystem -folder $folder -Outfilename $outfilename -Base $base `
 -Server $server -hashlen $hashlen -no_hash $no_hash -extruct $extruct `
 -Compliance $compliance -Monitor $monitor `
 -Start $start "" -StartFn $startfn `
 -Makves_url $makves_url -Makves_user $makves_user -Makves_pwd $makves_pwd -Logfilename $logfilename