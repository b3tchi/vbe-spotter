function GetApp($scriptPath) {
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

  # $TargetApp = $Access.Run("GetApp", [ref]$scriptPath) #use [ref] for optinal COM parameters
  $TargetApp = [Microsoft.VisualBasic.Interaction]::GetObject($scriptPath)

  return $TargetApp
}

function CloseDb($Access) {
  $Access.Quit(2)
}

function AccessJSON($Access, $command) {
  $rs = $Access.Run("QueryGet", [ref]$command) #use [ref] for optional COM parameters

  if($null -eq $rs){
    $json = ""
  }else{
    $json = ConvertFromRs($rs) | ConvertTo-Json
  }
  # Write-Information $myTestObject
  return $json
}

function AccessProcedure_old($app, $command, $arguments) {
  $data = $arguments."data" #get first object in array

  #temp vars add

  # Fill Json Data
  $db = $app.CurrentDb()

  foreach ($item in $data) {
    ConvertToRs $db $item
  }

  $output = $app.Run("ExecCommand", [ref]$command) #use [ref] for optinal COM parameters
  # $myTestObject = $output | ConvertFrom-Json
  # Write-Information $myTestObject

  #return output tbd
  return $output
}

function AccessProcedure($app, $pson){
#   $json = @'
# {
#   "name":"p04_StageSequence"
#   ,"arguments":{
#     "StageID":"1"
#     ,"04_StageSequence":[
#       {"ItemID":1,"Sequence":1}
#       ,{"ItemID":2,"Sequence":2}
#     ]
#   }
# }
# '@
#
#   $pson = $json | ConvertFrom-Json
  # $app = GetApp $dbFullPath

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
  $respAr = $app.Run($pson.name, [ref]$batchID) #use [ref] for optional COM parameters

  $respHash = @{
    # RecordsAffected = "$respAr[0]"
    Status = "$respAr[0]"
    Error = "$respAr[1]"
  }

  # Write-Host $respHash

  #return output tbd
  return $respHash
}

function AccessCmd($app, $command, $arguments) {

  # Fill Json Data
  $db = $app.CurrentDb()

  # Write-Host $app.DBEngine.Errors.Count

  $qdf = $db.QueryDefs($command)
  $pars = $qdf.Parameters


  foreach ($par in $pars){

    $parName = $par.Name
    [string]$parVal = "" #work only with string for parameter value ?
    $parVal = $arguments."$parName"

    # Write-Host $parName
    # Write-Host $par.Type
    # Write-Host $arguments."$parName"
    # Write-Host $parVal.ToString() # | Get-TypeData

    $par.Value = $parVal
    # $par = $null
  }

  $err = ""
  $status = 0

  try {
    $qdf.Execute(128) #128 = dbFailOnError 512 = dbSeeChanges
  }
  catch {
    $status = 500
    $err = $app.DBEngine.Errors($app.DBEngine.Errors.Count - 1).Description
  }

  # Write-Host $db.RecordsAffected
  $respHash = @{
    RecordsAffected = $qdf.RecordsAffected
    Status = "$status"
    Error = "$err"
  }

  Write-Host $respHash

  #return output tbd
  return $respHash
}



function GetBatchID($db, $tableName) {

  $creationGUID = '{' + [guid]::NewGuid() + '}'

  #InsertID
  $rs = $db.OpenRecordset($tableName) #dynaset

  $rs.AddNew()
  $rs.Fields("CreationGUID").Value = $creationGUID
  $rs.Update()
  $rs.close()

  #get ID
  $rs = $db.OpenRecordset("SELECT * FROM $tableName WHERE CreationGUID = $creationGUID")

  ##return value
  Write-Output $rs.fields("BatchID").Value

  $rs.close()

}

function RsBatchTable($db, $table, $batchID) {
  # $itemprops = $psO.PsObject.Properties
  # $table = $itemprops | Select-Object -First 1

  $tableName = $table.Name
  $records = $table.Value

  #Open recordset
  # $db.Execute("DELETE FROM $tableName")
  $rs = $db.OpenRecordset($tableName)

  foreach ($record in $records) {

    # $tableName
    $fields = $record.PsObject.Properties
    # $fields = $fields | Get-Member -MemberType NoteProperty # | Select-Object -Property Name
    # write-host ------
    $rs.AddNew()

    foreach ($field in $fields) {
      # Access the name of the property
      # write-host $object_properties.Name
      # Access the value of the property

      RsEnterValue $rs $field.Name $field.Value

      # $fld = $rs.Fields($field.Name)
      # write-host $field.Name $field.Value $fld.Name
    }

    RsEnterValue $rs "batchID" $batchID

    $rs.Update()
  }

  $rs.close()
}

function RsEnterValue($rs, $fieldName, $fieldValue){

  try {
    $rsfld = $rs.Fields($fieldName)
  }
  catch {
    $rsfld = $null
    write-host $fieldName + " not found"
  }

  if ($null -ne $rsfld) {
    # $fieldValue = $field.Value
    # if ($fieldValue.GetType().Name -eq 'String') {
      # $rs.Fields($fieldName).Value = "$fieldValue" #$strA
    # }
    # else {
    # }
    try{
      # $rs.Fields($fieldName).Value = $fieldValue
      $rs.Fields($fieldName).Value = "$fieldValue"
    }
    catch{
      write-Host $error[0]
      write-host "types not matching slot:" + $rs.Fields($fieldName).Value.GetType().Name + "in:" + $fieldValue.GetType().Name
      write-host "name:" + $fieldName + " value:" $fieldValue
    }
  }
}
function ConvertToRs($db, $psO) {
  $itemprops = $psO.PsObject.Properties
  $table = $itemprops | Select-Object -First 1

  $tableName = $table.Name
  $records = $table.Value

  #Open recordset
  $db.Execute("DELETE FROM $tableName")
  $rs = $db.OpenRecordset($tableName)

  foreach ($record in $records) {

    # $tableName
    $fields = $record.PsObject.Properties
    # $fields = $fields | Get-Member -MemberType NoteProperty # | Select-Object -Property Name
    # write-host ------
    $rs.AddNew()

    foreach ($field in $fields) {
      # Access the name of the property
      # write-host $object_properties.Name
      # Access the value of the property
      try {
        $rsfld = $rs.Fields($field.name)
      }
      catch {
        $rsfld = $null
        write-host $field.name + " not found in $tablename"
      }

      if ($null -ne $rsfld) {
        $value = $field.Value
        if ($value.GetType().Name -eq 'String') {
          $rs.Fields($field.Name).Value = "$value" #$strA
        }
        else {
          $rs.Fields($field.Name).Value = $value
        }
      }
      # $fld = $rs.Fields($field.Name)
      # write-host $field.Name $field.Value $fld.Name
    }

    $rs.Update()
  }

  $rs.close()
}

function ConvertFromRs($rs) {
  $rs.MoveLast()
  $rs.MoveFirst()
  # $rs.RecordCount

  $fldCount = $rs.Fields.Count
  $data = @()
  while ($rs.EOF -ne $true) {
    $rec = @{}

    for ($i = 0; $i -lt $fldCount; $i++) {
      $rec | Add-Member -NotePropertyName $rs.Fields($i).Name -NotePropertyValue $rs.Fields($i).Value
    }

    # $rec | ConvertTo-Json
    $data += $rec

    $rs.MoveNext()
  }

  # $rs.Close()

  return $data
}

function AppVbProj($app) {

  #for access
  return $app.VBE.VBProjects(1)
}

function CodeExport($vbproj, $CWD) {
  $COMPONENT_TYPE_MODULE = 1
  $COMPONENT_TYPE_CLASS = 2
  $modules = $vbproj.VBComponents;
  # $exportedModules = 0

  # Write-Host $module.Type
  foreach ($module in $modules) {

    $moduleFilename = switch ($module.Type){
      $COMPONENT_TYPE_MODULE { "src\Modules\$($module.Name).bas" }
      $COMPONENT_TYPE_CLASS { "src\Classes\$($module.Name).cls" }
      default { "" }
    }

    if ($moduleFilename -eq ""){
      continue
    }

    $moduleDestination = [IO.Path]::Combine($CWD, $moduleFilename)
    $module.Export($moduleDestination)
    # $exportedModules += 1
  }

  #TODO decide if return exported pieces
  #TODO file exitsts delete ?

}

function FirstStringByPatternFile($moduleFile, $rxPattern){

  $rxMatches = $moduleFile `
    | Select-String $rxPattern -AllMatches `
    | Foreach-Object {$_.Matches}

  $rxGroup = $rxMatches | Foreach-Object {$_.Groups[1].Value} | Select-Object -First 1

  return $rxGroup
}

function ModuleImport($vbproj, $moduleFile) {

}
