# Start-Process powershell -ArgumentList "-noexit", ".\Testing\AccessWorker.ps1"

Import-Module $PSScriptRoot/../AccessRunDb.ps1
Import-Module $PSScriptRoot/../Config.ps1

# Invoke-Command { & "powershell.exe" } -NoNewScope
# . $PSHOME\Profile.ps1
# . $Profile
# $scriptPath = Split-Path $psise.CurrentFile.FullPath #$Pwd.Path.ToString()
# $scriptPath = $PSScriptRoot
# $scriptPath = Split-Path -Parent $PSCommandPath
# $dbName = "Test.accdb"
# $dbFullPath= $scriptPath + "\" + $dbName
# $dbFullPath= "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\TestDb.accdb"

# $app = GetApp $dbFullPath

function jsonTest(){
  $app = GetApp $dbFullPath

  $jsonR = AccessJSON $app "Test"

  Write-Information $jsonR

}

function CreateRec(){

  $app = GetApp $dbFullPath
  $db = $app.CurrentDb()

  ConvertToRs3 $db


}

function AdoCreate(){

  $provider =
    "Provider=Microsoft.ACE.OLEDB.12.0;" +
    "Data Source=$dbFullPath"

  $cn = new-object -comObject ADODB.connection
  $cn.connectionString = $provider
  $cn.open()

  # $adOpenDynamic = 2
  # $adOpenStatic= 3
  # $adLockOptimistic = 3

  $rs = new-object -comObject ADODB.recordset
  # Wait-Debugger

  $rs.Open("SELECT * FROM 04_Batch", $cnd, 3, 3)

  $rs.AddNew()
  $rs.Fields("BatchID").Value
  $rs.Update()

  $rs.Close()
  $cn.close()

}

function GetBatchID_TEST(){

  $app = GetApp $dbFullPath
  $db = $app.CurrentDb()

  $id = GetBatchID $db "04_Batch"

  $id

}

function AccessProcedureTest(){
  $json = @'
{
  "name":"p04_StageSequence"
  ,"arguments":{
    "StageID":"1"
    ,"04_StageSequence":[
      {"ItemID":1,"Sequence":1}
      ,{"ItemID":2,"Sequence":2}
    ]
  }
}
'@

  $pson = $json | ConvertFrom-Json
  $app = GetApp $dbFullPath

  AccessProcedure $app $pson

}

function ParseJsonPson(){
  $json = @'
{
  "name":"p04_StageSequence"
  ,"arguments":{
    "StageID":"1"
    ,"04_StageSequence":[
      {"ItemID":1,"Sequence":1}
      ,{"ItemID":2,"Sequence":2}
    ]
  }
}
'@

  $pson = $json | ConvertFrom-Json
  $app = GetApp $dbFullPath

  $db = $app.CurrentDb()

  # $pson.name
  # $pson.arguments

  $batchNr = ($pson.name).SubString(1,2)
  $batchTable = $batchNr + "_Batch"

  [int]$batchID = GetBatchID $db $batchTable
  # $batchID

  ## Open batch record
  $rs = $db.OpenRecordset("SELECT TOP 1 * FROM $batchTable WHERE BatchID = $batchID")
  $rs.Edit()

  foreach($prop in $pson.arguments.PsObject.Properties){
    # $prop.Name
    if($prop.Value -is [array]) {

      RsBatchTable $db $prop $batchID

    }else{

      RsEnterValue $rs $prop.Name $prop.Value
      # $rs.fields($prop.Name).Value = $prop.Value
    }
  }

  $rs.Update()
  $rs.Close()

  ## CALL VBA PROCEDURE
  $app.Run($pson.name, [ref]$batchID) #use [ref] for optional COM parameters

}


function cmdTestx(){
  $app = GetApp $dbFullPath

  $jsonS = @'
{"name":"SaveTitle","arguments":{"pText":"Itemx","pItemID":23}}
'@

  $pson = $jsonS | ConvertFrom-Json

  $res = AccessCmd $app $pson.name $pson.arguments
  Write-Host ($res | Format-List | Out-String)
}

function cmdTest(){
  $app = GetApp $dbFullPath

  $jsonS = @'
{"name":"UpdateStage","arguments":{"pStageID":6,"pItemID":8}}
'@

  $pson = $jsonS | ConvertFrom-Json

  $res = AccessCmd $app $pson.name $pson.arguments
  Write-Host ($res | Format-List | Out-String)
}

function ExportCodeTest(){
  $app = GetApp $dbFullPath

  $dbCodePath= "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\"
  $vbProj = AppVbProj $app
  CodeExport $vbProj $dbCodePath
  Write-Host "test"
}

function ReadModuleName(){

  $dbCodePath= "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\src\Modules\Module2.bas"
  $codeFile = Get-Content $dbCodePath -Raw

  # Write-Host $codeFile.GetType()
  # Write-Host "code:$codeFile"
  $rxPattern = 'Attribute VB_Name = "(.*)"'
  $ret =  FirstStringByPatternFile $codeFile $rxPattern
  Write-Host "ret: $ret"
  # Write-Host "end"
}

#what to test
# ExportCodeTest
# cmdTestx
# ReadModuleName
# jsonTest
# AdoCreate
# ParseJsonPson
AccessProcedureTest
# GetBatchID_TEST
# CreateRec
