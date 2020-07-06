param(
    [string]$makves_url = "http://127.0.0.1:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$outfile = "./summary.html",
    [string]$title = $null,
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

$templatefile = "./template-summary.mustache"

Import-Module Poshstache -Verbose

function preprocesItem {
    param (
        $data
    )
    if ("user" -in $data.type) {
        $data | Add-Member -MemberType NoteProperty -Name is_user -Value $true -Force
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


Import-Module Poshstache -Verbose
$pair = "${makves_user}:${makves_pwd}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue}

Add-Type -AssemblyName 'System.Net.Http'

function convertSize (
    $bytes, 
    [int]$precision
    ) 
{
    foreach ($i in ("bytes","KB","MB","GB","TB")) {
        if (($bytes -lt 1000) -or ($i -eq "TB")) {
            $bytes = ($bytes).ToString("F0$precision")
            return $bytes + " $i"
        } else {
            $bytes /= 1024
        }
    }
}


function preprocess ($data) {
    if ("items" -in $data.PSobject.Properties.Name) {

        $data.items | ForEach-Object {
            if ("basic_score" -in $_.PSobject.Properties.Name) {
                $score = [math]::Ceiling($_.basic_score * 100)
                $_ | Add-Member -MemberType NoteProperty -Name score -Value $score -Force

                if ($_.basic_score -ge 0.8) {
                    $_ | Add-Member -MemberType NoteProperty -Name score_color -Value "#f86c6b" -Force
                } else {
                    if ($_.basic_score -ge 0.3) {
                        $_ | Add-Member -MemberType NoteProperty -Name score_color -Value "#ffc107" -Force
                    } else {
                        $_ | Add-Member -MemberType NoteProperty -Name score_color -Value "#4dbd74" -Force
                    }
                }
            }

            if ("size" -in $_.PSobject.Properties.Name) {
                $size_string = convertSize $_.size 2
                $_ | Add-Member -MemberType NoteProperty -Name size_string -Value $size_string -Force            
            }
        }
    }
    if ("data" -in $data.PSobject.Properties.Name) {
        $d =  $data.data
        if ("basic_score" -in $d.PSobject.Properties.Name) {
            $score = [math]::Ceiling($d.basic_score * 100)
            $d | Add-Member -MemberType NoteProperty -Name score -Value $score -Force

            if ($d.basic_score -ge 0.8) {
                $d | Add-Member -MemberType NoteProperty -Name score_color -Value "#f86c6b" -Force
            } else {
                if ($d.basic_score -ge 0.3) {
                    $d | Add-Member -MemberType NoteProperty -Name score_color -Value "#ffc107" -Force
                } else {
                    $d | Add-Member -MemberType NoteProperty -Name score_color -Value "#4dbd74" -Force
                }
            }
        }
        
        if ("size" -in $d.PSobject.Properties.Name) {
            $size_string = convertSize $d.size 2
            $d | Add-Member -MemberType NoteProperty -Name size_string -Value $size_string -Force            
        }
    }
    if ("risk" -in $data.PSobject.Properties.Name) {
        $score = [math]::Ceiling($data.risk * 100)
        $data | Add-Member -MemberType NoteProperty -Name score -Value $score -Force

        if ($data.risk -ge 0.8) {
            $data | Add-Member -MemberType NoteProperty -Name score_color -Value "#f86c6b" -Force
        } else {
            if ($d.risk -ge 0.3) {
                $data | Add-Member -MemberType NoteProperty -Name score_color -Value "#ffc107" -Force
            } else {
                $data | Add-Member -MemberType NoteProperty -Name score_color -Value "#4dbd74" -Force
            }
        }
    }

    return $data
}

function getdata { 
    param (
        [string]$makves_url = "http://127.0.0.1:8000/events/explore",
        [string]$makves_user = "admin",
        [string]$makves_pwd = "admin"
    )
    
    Add-Type -AssemblyName 'System.Net.Http'
    
    Try
    {
        $pair = "${makves_user}:${makves_pwd}"
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
        $base64 = [System.Convert]::ToBase64String($bytes)
        $basicAuthValue = "Basic $base64"
        [psobject]$headers = @{ Authorization = $basicAuthValue}

        $response = Invoke-WebRequest -Uri $makves_url -Method Get -Headers $headers
        $jsonObj = ConvertFrom-Json $([String]::new($response.Content))

        return preprocess $jsonObj
    }
    Catch {
        Write-Host "Error send data to server:" + $PSItem.Exception.Message
        return $Null
    }
}

$requests = @()

$requests += New-Object PSObject -Property @{
    Field= "users"
    Url="/ldap/stat?type=users"
}

$requests += New-Object PSObject -Property @{
    Field= "computers"
    Url="/ldap/stat?type=computers"
}

$requests += New-Object PSObject -Property @{
    Field= "files"
    Url="/file/stat?not_actual=true&total=true&byType=true&compliance=true&byCompliance=true&stolled=true&openAccess=true&duplicates=true"
}

$requests += New-Object PSObject -Property @{
    Field= "events"
    Url="/events/stat"
}

$requests += New-Object PSObject -Property @{
    Field= "mailboxes"
    Url="/mailboxes/stat"
}

$reportTime = get-date -f "dd.MM.yyyy hh:mm"


$data = New-Object PSObject -Property @{
    title=$title
    report_time=$reportTime
}



$requests | ForEach-Object {
    $cur = $_
    $url = $makves_url + $cur.Url
    $d = getdata $url $makves_user $makves_pwd
    if ($null -ne $d) {
        $data | Add-Member -MemberType NoteProperty -Name $cur.Field -Value $d -Force   
    }    
}

$JSON = $data | ConvertTo-Json -Depth 5

$jsonString = $JSON.ToString()
Import-Module Poshstache
$res = ConvertTo-PoshstacheTemplate -InputFile $templatefile -ParametersObject $jsonString 
if ("" -ne $outfile) {
    $res | Out-File $outfile -Force -Encoding "UTF8"
}




if (($null -ne $mail) -and ("" -ne $mail)) {

    $SMTPServer = $mail.server
    $SMTPClient = New-Object Net.Mail.SMTPClient($SmtpServer, $mail.port)
    $SMTPClient.EnableSSL = $mail.EnableSSL
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($mail.user, $mail.pwd);

    # From Core @ stackoverflow.com
    $emailMessage = New-Object System.Net.Mail.MailMessage
    $emailMessage.From = $mail.from
    foreach ($recipient in $mail.to)
    {
        $emailMessage.To.Add($recipient)
    }
    $emailMessage.IsBodyHtml = $true
    $emailMessage.Subject = $title
    $emailMessage.Body = $res
    # Do we have any attachments?
    # If yes, then add them, if not, do nothing
    ##if ($Arry_EmailAttachments.Count -ne $NULL)
    ##{
    ##    $emailMessage.Attachments.Add()
    ##}
    $SMTPClient.Send($emailMessage)

}




