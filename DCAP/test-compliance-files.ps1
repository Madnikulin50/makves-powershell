
 param (
    [string]$file = "C:\work\test\test"
)

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$env:Path += ";$scriptPath\MakvesCompliance"

Import-Module -Name $scriptPath"\MakvesCompliance" -Verbose

$file | Search-FileCompliance

Get-ChildItem $file -Recurse | Foreach-Object { $_.FullName } | Search-FileCompliance