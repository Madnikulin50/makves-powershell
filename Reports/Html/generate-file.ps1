param(
    [string]$makves_url = "http://127.0.0.1:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$outfile = "",
    [string]$path = '/file/view/5327008',
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
    $outfile = "$scriptPath\file.html"
}

Import-Module -Name $scriptPath"\MakvesHtmlReport" -Verbose



function preprocesItem {
    param (
        $data
    )
    if ("file" -in $data.type) {
        $data | Add-Member -MemberType NoteProperty -Name is_file -Value $true -Force
    }
    if ("folder" -in $data.type) {
        $data | Add-Member -MemberType NoteProperty -Name is_folder -Value $true -Force
    }

    if ("access" -in $data.PSobject.Properties.Name) {
        $data.access | ForEach-Object {
            if (10 -eq $_.read) {
                $_ | Add-Member -MemberType NoteProperty -Name allow_read -Value $true -Force
            }
            if (20 -eq $_.read) {
                $_ | Add-Member -MemberType NoteProperty -Name deny_read -Value $true -Force
            }
            if (10 -eq $_.write) {
                $_ | Add-Member -MemberType NoteProperty -Name allow_write -Value $true -Force
            }
            if (20 -eq $_.write) {
                $_ | Add-Member -MemberType NoteProperty -Name deny_write -Value $true -Force
            }
        } 
    }

}
function preprocess {
    param (
        $data
    )
    preprocesItem $data.data
        
}


New-MakvesSimpleReport -templatefile "$scriptPath\template-file.mustache" -outfilename $outfile `
 -Makves_url $makves_url -Makves_user $makves_user -Makves_pwd $makves_pwd `
 -Title "Файл" -path $path -search $search -preprocessor $function:preprocess -mail $mail

Remove-Module MakvesHtmlReport
