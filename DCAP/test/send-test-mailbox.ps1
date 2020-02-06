param (
    [string]$fn="C:\work\test\Matrix23112019\mailboxes_201911230532.json",
    [string]$makves_url = "http://127.0.0.1:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin"
 )


 ## Init web server 
$uri = $makves_url + "/data/upload/agent"
$pair = "${makves_user}:${makves_pwd}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue}

if ($makves_url -eq "") {
    $uri = ""
    Add-Type -AssemblyName 'System.Net.Http'
}

function store($cur) {
    if ($uri -ne "") {
      $cur | Add-Member -MemberType NoteProperty -Name Type -Value "exchange-mailbox" -Force 
      $cur | Add-Member -MemberType NoteProperty -Name Forwarder -Value "exchange-mailboxes-forwarder" -Force
      $JSON = $cur | ConvertTo-Json
      $body = [System.Text.Encoding]::UTF8.GetBytes($JSON.ToString());
      Invoke-WebRequest -Uri $uri -Method Post -Body $body -ContentType "application/json" -Headers $headers
    }
}

$contents = get-content $fn -encoding UTF8 -Raw

$items = $contents -split "`n}`r`n{"

$items | Foreach-Object {
    Try {
        $cur = $_
        if ($cur[0] -ne '{`r`n') {
            $cur = "{"  + $cur
        }
        if ($cur[$cur.length - 1] -ne '}') {
            $cur = $cur + "`r`n}"
        }
        $json = $cur | ConvertFrom-Json
        store($json)
    } Catch {
        Write-Host "$($_.Exception.Message)"
      }
}
