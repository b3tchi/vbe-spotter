
Import-Module ./AccessRunDb.ps1

# $scriptPath = Split-Path $psise.CurrentFile.FullPath #$Pwd.Path.ToString()

# $scriptPath = $PSScriptRoot
# $scriptPath = Split-Path -Parent $PSCommandPath
# $scriptPath = "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps"
# $shelperName = "shelper.accdb"

# $shelperPath = $scriptPath + "\" + $shelperName
$dbFullPath = "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\TestDb.accdb"

$app = GetApp $dbFullPath #$shelperPath

# $app.CurrentProject.FullName

$db = $app.CurrentDb()

$data = ConvertFromRs $db "AllItems"

$data | ConvertTo-Json
