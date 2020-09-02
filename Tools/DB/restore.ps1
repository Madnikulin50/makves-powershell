
param (
 [string]$user = "postgres",
 [string]$pwd = "Zse45rdx",
 [string]$server = "localhost",
 [string]$db = "makvest",
 [string]$file = "C:\work\makves\makves-powershell\Tools\DB\database_202007170809.dump"
)

$postgresDir = "C:\Program Files\PostgreSQL\11"

Write-Host "db: " $db
Write-Host "server: " $server

Write-Host "user: " $user
Write-Host "pwd: " $pwd

$connection = "postgresql://$($user):$($pwd)@$($server):5432/$($db)"

&($postgresDir + "\bin\pg_restore.exe") ("--clean")("--dbname="+$connection) ("`""+$file+"`"")

##pg_restore  -U postgres -d makvesr  C:\work\makves\makves-powershell\Tools\DB\makves.dump

