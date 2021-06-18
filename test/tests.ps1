Import-Module $PSScriptRoot/../src/lib_spotter.ps1

# [string]$dbFile = "~/Repositories/VbaSpotter/src/TestDb.accdb"
# [string]$ExportLocation ="~/Repositories/VbaSpotter/test/codes/"

# [string]$dbFile = "C:\Users\czJaBeck\Onedrive - LEGO\Documents\Wdd_v2.accdb"
# [string]$exportLocation ="C:\Users\czJaBeck\Repositories\PPM_vba_codes\"


[string]$dbFile = "C:\Users\czJaBeck\Repositories\LocalWeb_Ps\TestDb.accdb"
[string]$ExportLocation ="C:\Users\czJaBeck\Repositories\LocalWeb_Ps\codes\"

[DateTime]$changeDate = (Get-Date).AddDays(-5)

[Hashtable]$modules= @{}

[string]$excelFile = "C:\Users\czJaBeck\OneDrive - LEGO\Documents\Unicodefix.xlsm"

#Test1
# RepoChanged $dbFile $exportLocation $changeDate

#Test2

Function Test2(){


  # ExportCode()
}

# function GetExcel_T0(){
#     GetExcel $excelFile
#
# }
# GetExcel_T0
#

function GetExcel_T(){
  try{
    $app = GetExcel $excelFile

    #validation
    # $app.Name
    $test = 'ok'

  }catch{
    $test = 'nok'

    Write-Warning $Error[0]

  }finally{

    $msg = "{0} - {1}" -f $MyInvocation.MyCommand, $test
    Write-Information $msg  -InformationAction Continue

  }
    return $app
}

function GetProject_T($app){

  try{
    $proj = GetProject $app


    # Write-Debug  $proj.Name
     $test = 'ok'
  }catch{
    Write-Warning $Error[0]
    $test = 'nok'
  }finally{
    $msg = "{0} - {1}" -f $MyInvocation.MyCommand, $test
    Write-Information $msg  -InformationAction Continue
  }

  return $proj

}

function CodeModule_T($proj, $moduleName){

  try{
    $module = GetCodeModule $proj $moduleName

    $test = 'ok'
  }catch{
    $test = 'nok'
    Write-Warning $Error[0]
  }finally{
    $msg = "{0} - {1}" -f $MyInvocation.MyCommand, $test
    Write-Information $msg  -InformationAction Continue
  }

  if ($test -eq 'ok') {
    return $module
  }

}

function GetCode_T($module){

  try{
    $code = GetCode $module
    $test = 'ok'
  }catch{
    $test = 'nok'
    Write-Warning $Error[0]
  }finally{
    $msg = "{0} - {1}" -f $MyInvocation.MyCommand, $test
    Write-Information $msg  -InformationAction Continue
  }

  if ($test -eq 'ok') {
    return $code
  }

}

function ExportForm_T($proj,$moduleName,$ExportPath){

  try{
    $module = GetCodeModule $proj $moduleName
    $code = ExportCode $module $ExportPath
    $test = 'ok'
  }catch{
    $test = 'nok'
    Write-Warning $Error[0]
  }finally{
    $msg = "{0} - {1}" -f $MyInvocation.MyCommand, $test
    Write-Information $msg  -InformationAction Continue
  }

}

function ImportForm_T($proj,$moduleName,$ExportLocation){

  # $moduleFilename = $moduleName+'.frm'
  $moduleDestination = [IO.Path]::Combine($ExportLocation, $moduleName+'.frm')

  try{
    $module = ImportCode $proj $moduleDestination
    $test = 'ok'
  }catch{
    $test = 'nok'
    Write-Warning $Error[0]
  }finally{
    $msg = "{0} - {1}" -f $MyInvocation.MyCommand, $test
    Write-Information $msg  -InformationAction Continue
  }

  if ($test -eq 'ok') {
    return $module
  }

}

function TestRun(){
  # $moduleName = 'lib_symbols'
  $moduleName = 'Sheet2'

  $moduleFilename = 'lib_symbols2.bas'
  $moduleDestination = [IO.Path]::Combine($ExportLocation, $moduleFilename)

  $moduleFilename2 = 'Sheet11.cls'
  $moduleDestination2 = [IO.Path]::Combine($ExportLocation, $moduleFilename2)

  $app = GetExcel_T

  # $app.Name

  $proj = GetProject_T $app

  # $proj.Name

  $module = CodeModule_T $proj $moduleName

  # $module.Name

  $code = GetCode_T $module

  #FormInOut Test
  ImportForm_T $proj 'UserForm1' $ExportLocation
  ExportForm_T $proj 'UserForm1' $ExportLocation



  # $modules cleanup
  # for ($i = 1; $i -lt 2; $i++) {
  #   $module1 = GetCodeModule $proj "lib_symbols2$i"
  #   RemoveCodeModule $proj $module1
  # }

  #import standard class
  # $module = ImportCode $proj $moduleDestination

  #import special class
  # $module = ImportCode $proj $moduleDestination2

}



# TestExcel
TestRun
