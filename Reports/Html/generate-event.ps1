param(
    [string]$makves_url = "http://127.0.0.1:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$outfile = "",
    [string]$path = '/events/view/2707222',
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
    $outfile = "$scriptPath\event.html"
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

function preprocess {
    param (
        $data
    )

    if ("data1" -in $data.data.PSobject.Properties.Name) {
   
        if ("aux1" -in $data.data.PSobject.Properties.Name) {
            if ($data.data.aux1.length -ge 256) {
                $data.data.aux1 = $data.data.aux1.SubString(0, 253) + "..."
            }
        }
        if ("aux2" -in $data.data.PSobject.Properties.Name) {
            if ($data.data.aux2.length -ge 256) {
                $data.data.aux2 = $data.aux2.SubString(0, 253) + "..."
            }
        }

        if ("contents" -in $data.data.PSobject.Properties.Name) {
            if ($data.data.contents.length -ge 256) {
                $data.data.contents = $data.data.contents.SubString(0, 253) + "..."
            }
        }
    }
    
}

New-MakvesSimpleReport -templatefile "$scriptPath\template-event.mustache" -outfilename $outfile `
 -Makves_url $makves_url -Makves_user $makves_user -Makves_pwd $makves_pwd `
 -Title "Событие" -path $path -search $search -preprocessor $function:preprocess -mail $mail 

Remove-Module MakvesHtmlReport

