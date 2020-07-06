param(
    [string]$makves_url = "http://127.0.0.1:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$outfile = "",
    [string]$path = '/ldap/view/S-1-5-21-2950234198-1677066530-3853381158-6127',
    [string]$search = '',
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
    $outfile = "$scriptPath\computer.html"
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


function preprocesItem {
    param (
        $data
    )
    if ("computer" -in $data.type) {
        $data | Add-Member -MemberType NoteProperty -Name is_computer -Value $true -Force
    }
    if ("group" -in $data.type) {
        $data | Add-Member -MemberType NoteProperty -Name is_group -Value $true -Force
    }
}
function preprocess {
    param (
        $data
    )
    preprocesItem $data.data
        
}


New-MakvesSimpleReport -templatefile "$scriptPath\template-computer.mustache" -outfilename $outfile `
 -Makves_url $makves_url -Makves_user $makves_user -Makves_pwd $makves_pwd `
 -Title "Компьютер" -path $path -search $search -preprocessor $function:preprocess

 #-mail $mail 

Remove-Module MakvesHtmlReport
