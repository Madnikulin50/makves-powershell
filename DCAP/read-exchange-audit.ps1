param (
    [string]$filter = "*",
    $outfilename = "mailbox-audit_"
)

$outfile = ""
$LogDate = get-date -f yyyyMMddhhmm 

if ($outfilename -ne "") {
    $outfile = "$($outfilename)_$LogDate.json"
    if (Test-Path $outfile) 
    {
        Remove-Item $outfile
    }
}

function store($cur) {
    $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
}

Write-Host "filter: $($filter)"

function inspect($cur) {
    Write-Host $cur
    try {       
        $data = Search-MailboxAuditLog -Identity $cur -ResultSize 2000 -LogonTypes Admin,Delegate -ShowDetails
        write-host "$($cur) Admin,Delefate: $($data.Count) Total entries Found"
        $data | ForEach-Object {
            $cur | Add-Member -MemberType NoteProperty -Name Forwarder -Value "exchange-mailbox-audit" -Force
            store $_
        }
        $data = Search-MailboxAuditLog -Identity $cur -ResultSize 2000 -LogonTypes Owner -ShowDetails
        write-host "$($cur) Owner: $($data.Count) Total entries Found"
        $data | ForEach-Object {
            $cur | Add-Member -MemberType NoteProperty -Name Forwarder -Value "exchange-mailbox-audit" -Force
            store $_
        }
        
    }
    Catch {
        Write-Host "Error on $($cur): $($_.Exception.Message)"
    }
}

#inspect mikulka@urd.local  
#Get-Mailbox -Identity mikulka@urd.local | Format-List Name,Audit*


Search-AdminAuditLog | ForEach-Object {
    $cur = $_
    $cur | Add-Member -MemberType NoteProperty -Name Forwarder -Value "exchange-admin-audit" -Force
    store $cur    
 }

 
Get-Mailbox -ResultSize Unlimited -Filter $filter | ForEach-Object {
    inspect $_.Name    
 }
 

#New-ManagementRoleAssignment -Name "AuditLogsRole" -User urd\root -Role "Audit Logs"


Write-Host OK -ForegroundColor Green

