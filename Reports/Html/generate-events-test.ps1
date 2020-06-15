param(
    [string]$template = "template-events.tmpl",
    [string]$makves_url = "http://127.0.0.1:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$outfile = "file"
)

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
        return $jsonObj
    }
    Catch {
        Write-Host "Error send data to server:" + $PSItem.Exception.Message
        return $Null
    }
}

$events = getdata "$makves_url/events/explore" $makves_user $makves_pwd

if ($null -eq $events) {
    return
}

$events | Add-Member -MemberType NoteProperty -Name title -Value  "События" -Force

$JSON = $events | ConvertTo-Json

$jsonString = $JSON.ToString()
$res = ConvertTo-PoshstacheTemplate -InputFile "$scriptPath\template-events.mustache" -ParametersObject $jsonString 
$res | Out-File ".\events.html" -Force -Encoding "UTF8"

$mail = @{
    Server="smtp.gmail.com"
    Port=587
    EnableSSL=$true
    From="madnikulin50@gmail.com"
    To="mn@makves.ru"
    User="madnikulin50"
    Pwd="oe3014n;"
}

if ($null -ne $mail) {

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
