param(
    [string]$makves_url = "http://127.0.0.1:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$outfile = "",
    [string]$path = '/events/explore',
    [string]$search = '{"sort":[{"colId":"time", "sort":"desc"}],"page":1,"limit":100}'
)

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$env:Path += ";$scriptPath\MakvesHtmlReport"
Write-Host $env:Path
if ("" -eq $outfile) {
    $outfile = "$scriptPath\events.html"
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

    $data.items | ForEach-Object {
        if ("aux1" -in $_.PSobject.Properties.Name) {
            if ($_.aux1.length -ge 256) {
                $_.aux1 = $_.aux1.SubString(0, 253) + "..."
            }
        }
        if ("aux2" -in $_.PSobject.Properties.Name) {
            if ($_.aux2.length -ge 256) {
                $_.aux2 = $_.aux2.SubString(0, 253) + "..."
            }
        }

        if ("contents" -in $_.PSobject.Properties.Name) {
            if ($_.contents.length -ge 256) {
                $_.contents = $_.contents.SubString(0, 253) + "..."
            }
        }
    }
    
}

New-MakvesSimpleReport -templatefile "$scriptPath\template-events.mustache" -outfilename $outfile `
 -Makves_url $makves_url -Makves_user $makves_user -Makves_pwd $makves_pwd `
 -Title "Последние события" -path $path -search $search -preprocessor $function:preprocess

 #-mail $mail 

Remove-Module MakvesHtmlReport

