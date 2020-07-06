param(
    [string]$makves_url = "http://127.0.0.1:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$outfile = "",
    [string]$path = '/ldap/explore/computers',
    [string]$search = '{"filter":{"type":{"type":"equals","filter":"computer"}},"sort":[{"colId":"cn"}],"page":1,"limit":100}',
    [string]$mail = $null
)
<# Перед запуском необходимо выполнить установку Postache из Интернет

Install-Module Poshstache

#>

<# 
#Пример  настройки почты

$mail = @{
    Server="smtp.gmail.com"
    Port=587
    EnableSSL=$true
    From="madnikulin50@gmail.com"
    To="mn@makves.ru"
    User="madnikulin50"
    Pwd="<pwd>"
}
#>

Import-Module Poshstache -Verbose

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$env:Path += ";$scriptPath\MakvesHtmlReport"
Write-Host $env:Path
if ("" -eq $outfile) {
    $outfile = "$scriptPath\computers.html"
}

Import-Module -Name $scriptPath"\MakvesHtmlReport" -Verbose
$mail = @{
    Server="smtp.gmail.com"
    Port=587
    EnableSSL=$true
    From="madnikulin50@gmail.com"
    To="mn@makves.ru"
    User="madnikulin50"
    Pwd="<pwd>"
}


New-MakvesSimpleReport -templatefile "$scriptPath\template-computers.mustache" -outfilename $outfile `
 -Makves_url $makves_url -Makves_user $makves_user -Makves_pwd $makves_pwd `
 -Title "Компьютеры" -path $path -search $search

 #-mail $mail 

Remove-Module MakvesHtmlReport
