#$xxx = 'Attribute VB_Name=`"Test`"' | Select-String '^Attribute VB_Name = `"(.*)`"$' -AllMatches
$modulFileu = @'
Attribute VB_Name = "Module2"
Option Compare Database
'End Function
'@

  $dbCodePath= "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\src\Modules\Module2.bas"
  $modulFile = Get-Content $dbCodePath -Raw

function x(){
$pattern = 'Attribute VB_Name = "(.*)"'


  $rxMatches = $modulFile `
    | Select-String $pattern -AllMatches `
    | Foreach-Object {$_.Matches}



  $rxGroups = $rxMatches | Foreach-Object {$_.Groups[1].Value} | Select-Object -First 1

#Wait-Debugger
}

function ModuleNameFromFile($moduleFile){

  # $sln = Get-Content $PathToSolutionFile
  $rxPattern = 'Attribute VB_Name = "(.*)"'
  $rxMatches = $modulFile `
    | Select-String $rxPattern -AllMatches `
    | Foreach-Object {$_.Matches}

  $rxGroup = $rxMatches | Foreach-Object {$_.Groups[1].Value} | Select-Object -First 1
  # Wait-Debugger
  # Write-Debug $modrx

  return $rxGroup
}


Write-Host $modulFile.GetType()
$ret = ModuleNameFromFile $modulFile

Write-Host "ret:$ret"