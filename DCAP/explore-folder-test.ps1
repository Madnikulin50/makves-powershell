param (
    [string]$folder = "C:\work\FinCert",
    [string]$outfilename = "folder", ##"",
    [string]$computer = "",
    [string]$base = "",
    [string]$server = "",
    [int]$hashlen = 1048576,
    [switch]$no_hash = $false,
    [switch]$extruct = $false,
    [switch]$compliance = $true,
    [switch]$monitor = $false,
    [int16]$threads = 1,
    [string]$start = "",
    [string]$startfn = "", ##".file-monitor.time_mark",
    [string]$makves_url =  "", ##"http://192.168.2.22:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string[]]$KB = ("C:\work\kb\research\test.json"),
    [string]$logfilename = ""
)

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$env:Path += ";$scriptPath\MakvesFileSystem"
Write-Host $env:Path

Import-Module -Name $scriptPath"\MakvesFileSystem" -Verbose

Test-FileSystem -folder $folder -Outfilename $outfilename -Base $base `
 -Server $server -Computer $computer `
 -hashlen $hashlen -no_hash $no_hash -extruct $extruct `
 -Compliance $compliance -Monitor $monitor -Threads $threads `
 -Start $start "" -StartFn $startfn -KB $KB `
 -Makves_url $makves_url -Makves_user $makves_user -Makves_pwd $makves_pwd -Logfilename $logfilename