$getObjCode = {
    $sc = New-Object -ComObject MSScriptControl.ScriptControl.1
    $sc.Language = 'JScript'

    $sc.AddCode('function myFunction(x){return GetObject(x)}')
  
    #$Ac = New-Object -ComObject 
    $Ac = $sc.codeobject.myFunction("C:\\Users\\czJaBeck\\Documents\\Vbox\\LocalWeb_Ps\\TestDb.accdb")
    # $Ac.Forms
    #$Ac
    # $Ac.Run("Test")
    # $rs = $Ac.CurrentDb.OpenRecordset("SELECT * FROM Table1")
    # $rsText = $rs.GetString()
    # $sc.codeobject.MyFunction(2)
  } #-runas32 | wait-job | receive-job

$job = Start-Job -ScriptBlock $getObjCode -runAs32 #b-ArgumentList @($jscode, $jsoncode)
$output = $job | Wait-Job | Receive-Job
# $output.GetType() #return output
$App = $output
# $App = $db.Parent
Remove-Job $job