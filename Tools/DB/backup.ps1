
param (
    [string]$user = "postgres",
    [string]$pwd = "Zse45rdx",
    [string]$server = "localhost",
    [string]$db = "makves",
    [string]$folder= ".",
    [string]$outfilename = "database"
 )

$postgresDir = "C:\Program Files\PostgreSQL\11"

Write-Host "db: " $db
Write-Host "server: " $server

Write-Host "user: " $user
Write-Host "pwd: " $pwd

$outfile = ""
$LogDate = get-date -f yyyyMMddhhmm

if ($outfilename -ne "") {
    $outfile = "$($folder)\$($outfilename)_$LogDate.dump"
    if (Test-Path $outfile) {
        Remove-Item $outfile
    } 
}

$connection = "postgresql://$($user):$($pwd)@$($server):5432/$($db)"

&($postgresDir + "\bin\pg_dump.exe") ("-Fc") ("-U " + $user) ("--dbname="+$connection) ("--file=`""+$outfile+"`"")