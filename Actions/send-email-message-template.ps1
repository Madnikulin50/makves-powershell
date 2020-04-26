param (
    [string]$from = "admin@makves.ru",
        [string[]]$to = (""),
        [string]$server = $PSEmailServer,
        [int]$port = 25,
        [bool]$usessl = $false,
        [string]$subject = "",
        [string]$body = ""
 )

Write-Host "user: " $user
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$env:Path += ";$scriptPath\MakvesSendEmail"

Import-Module -Name $scriptPath"\MakvesSendEmail" -Verbose

Send-Email -From $from -To $To -Server $server `
 -Port $port -UseSsl $usessl `
 -Subject $subject `
 -Body $body

Write-Host "Send finished..."  -ForegroundColor Green
Remove-Module MakvesSendEmail
