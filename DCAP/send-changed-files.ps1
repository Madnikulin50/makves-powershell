param (
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")] [string]$folder = "C:\work\test",
    [string]$include = "*",
    [string]$exclude = "",
    [string]$tempfolder = "",
    [string]$makves_url = "http://127.0.0.1:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$start = "",
    [string]$startfn = "", ##".file-monitor.time_mark",
    [string]$timespan = 60,
    [string]$loop = 0
 )

$uri = $makves_url + "/data/upload"
$pair = "${makves_user}:${makves_pwd}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue}
Add-Type -AssemblyName 'System.Net.Http'

if ($startfn -ne "") {
    Try
    {
        $start = Get-Content $startfn
    }
    Catch {
        Write-Host "Error read time mark:" + $PSItem.Exception.Message
        $start = ""
    }
} 


function sendFile($entry) {
    $fullname = $entry.Fullname
    $name = $entry.Name
    $tempfile = ""
    $readfn = $fullname
    if ($tempfolder -ne "") {
        $tempfile = $fullname + "copy"
        Copy-Item $fullname $tempFile
        $readfn = $tempfile
    }

    Try
    {
        $fileBin = [IO.File]::ReadAllBytes($readfn)
        $enc = [System.Text.Encoding]::GetEncoding("iso-8859-1")
        $fileEnc = $enc.GetString($fileBin)

        $boundary = [System.Guid]::NewGuid().ToString()
        $LF = "`r`n"
        $bodyLines = (
            "--$boundary",
            "Content-Disposition: form-data; name=`"file`"; filename=`"$name`"",
            "Content-Type: application/octet-stream$LF",
            $fileEnc,
            "--$boundary--$LF"
        ) -join $LF

        Write-Host "send file " + $name

        Invoke-RestMethod -Uri $uri -Method Post -ContentType "multipart/form-data; boundary=`"$boundary`"" -Body $bodyLines -Headers $headers
    }
    Catch {
        Write-Host "send file " $name " error:" + $PSItem.Exception.Message
    }
    if ($tempfile -ne "") {
        Remove-Item -path $tempfile
    }
}



function processFile($entry) {
    sendFile $entry
}

function worker() {
    $markTime = Get-Date -format "yyyyMMddHHmmss"
    
    Try
    {
        $childs = Get-ChildItem $folder -Filter $include -Include $include -Recurse 

        if ($start -ne "") {
            Write-Host "start: " $start
            $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)
    
            $childs = $childs | Where-Object { $_.LastWriteTime -gt $starttime } 
        }
        $childs | Foreach-Object {
            Try
            {
                if ($_.PSIsContainer -eq $False) {
                    processFile $_
                }
                
            }
            Catch {
                Write-Host "processFile error:" + $PSItem.Exception.Message
            }
        }
        
    } Catch {
        Write-Host "Error Get-ChildItem:" + $f + ":" $PSItem.Exception.Message
        return;
    }
    if ($startfn -ne "") {
        $markTime | Out-File -FilePath $startfn -Encoding UTF8
        Write-Host "Store new mark: " $markTime
    }
    $start = $markTime
}

if ($loop -eq 0) {
    worker
}
 


 