$dbfilePath = "C:\Users\czJaBeck\Repositories\LocalWeb_Ps\TestDb.accdb"
$ExportLocation ="C:\Users\czJaBeck\Repositories\LocalWeb_Ps\codes\"
#loop test
$access = New-Object -ComObject Access.Application

$access.OpenCurrentDatabase($dbfilePath)

$db = $access.CurrentDb

$access.visible = $true
# Write-Host $access.CurrentProject.Fullname
# Write-Host $access.CurrentProject.AllModules.Count

$conts = $access.CurrentProject.AllModules

$dbModified = (Get-Item $dbfilePath).LastWriteTime

$keephook = $true

While ($keephook) {

    Write-Host $dbModified

    #     $modules= @()
    #
    #     foreach($module in $conts){
    #     # Write-Host $module.Name $module.DateModified
    #         $modDate = $module.DateModified.ToString('yyyy-mm-dd HH:mm:ss')
    #             $access.SaveAsText(5, $module.Name, $ExportLocation +  $module.Name + ".txt")
    #             $modules += @{"module"= $module.Name; "modified"= $modDate}
    #     }
    #     # Write-Host ($modules | ConvertTo-JSON)
    #
    # $modules | ConvertTo-JSON | Out-File $ExportLocation"modules.json"

    if ($looping -ge 100){
        $keephook = $false
    }
}
    # For Each d In c.Documents
        # Application.SaveAsText acForm, d.Name, sExportLocation & "Form_" & d.Name & ".txt"
    # Next d

#     Set c = db.Containers("Reports")
#     For Each d In c.Documents
#         Application.SaveAsText acReport, d.Name, sExportLocation & "Report_" & d.Name & ".txt"
#     Next d
#
#     Set c = db.Containers("Scripts")
#     For Each d In c.Documents
#         Application.SaveAsText acMacro, d.Name, sExportLocation & "Macro_" & d.Name & ".txt"
#     Next d
#     
#     Set c = db.Containers("Modules")
#     For Each d In c.Documents
#         Application.SaveAsText acModule, d.Name, sExportLocation & "Module_" & d.Name & ".txt"
#     Next d
#     
#     For i = 0 To db.QueryDefs.Count - 1
#         Application.SaveAsText acQuery, db.QueryDefs(i).Name, sExportLocation & "Query_" & db.QueryDefs(i).Name & ".txt"
#     Next i
#     
#     Set db = Nothing
#     Set c = Nothing
#     
#     MsgBox "All database objects have been exported as a text file to " & sExportLocation, vbInformation
#     
# Exit_ExportDatabaseObjects:
#     Exit Sub
#     
# Err_ExportDatabaseObjects:
#     MsgBox Err.Number & " - " & Err.Description
#     Resume Exit_ExportDatabaseObjects
#     
# End Sub
$db.close
$access.Quit()
