param (
    [string]$makves_url = "http://localhost:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin",
    [string]$services = "service.csv",
    [string]$resources = "resource.csv",
    [string]$acnts = "acnt.csv",
    [string]$out = "services.json"
)

$s = Import-Csv $services -Delimiter ';'

$r = Import-Csv $resources -Delimiter ';'

$a = Import-Csv $acnts -Delimiter ';'



foreach($item in $s) {

    $components = @()
    foreach($ritem in $r) {
        if ($ritem.ResSvc -eq $item.SvcID) {
            $accounts = @()
            foreach($aitem in $a) {
                if ($aitem.AcntType -eq $ritem.ResID) {
                    
                    $ahash = @{
                        Type= "iam-account"
                        ID=$ritem.AcntID
                        Name= $aitem.AcntName
                        Note= $aitem.AcntNotes
                        Org= $aitem.AcntIntOrg
                        Disabled= $aitem.AcntDisabled -eq "ИСТИНА"
                        Dep = $aitem.AcntIntDep
                        Div	= $aitem.AcntIntDiv
                        Owner =	$aitem.AcntIntOwner
                        Expiration = $aitem.AcntExpDate
                        Rights = $aitem.AcntRights
                        IsSystem = $aitem.IsSystem -eq "ИСТИНА"
                    }
                    $accounts += New-Object PSObject -Property $ahash
                }
            }
            $conv = $accounts | ConvertTo-Json
            $chash = @{
                
                Type= "iam-resource"
                ID=$ritem.ResID
                Name= $ritem.ResName
                Note= $ritem.ResNote
                BaseLdap=$ritem.ResBaseLDAP
                Accounts= $conv
            }
            $components += New-Object PSObject -Property $chash
        }
       
    }

    $hash = @{
        Type= "iam-service"
        ID=$ritem.SvcID
        Name= $item.SvcName
        Note= $item.SvcNote
        Code= $item.SvcCode
        Components=$components
    }
    $cur = New-Object PSObject -Property $hash

    $cur | ConvertTo-Json | Out-File -FilePath $out -Encoding UTF8 -Append
}
