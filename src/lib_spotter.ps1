# Write-Information 'lib loaded' -InformationAction Continue

function GetExcel($scriptPath) {

  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

  $TargetApp = [Microsoft.VisualBasic.Interaction]::GetObject($scriptPath)

  # $TargetApp

  return $TargetApp.Application

}

function GetProject($officeApp){

    $appName = $officeApp.Name
    $appName = $appName.split(" ")[1]

    # $officeApp
    # Write-Debug $appName

    if ($appName -eq "Access"){
        $vbproj = $officeApp.VBE.VBProjects(1)
    }elseif ($appName -eq "Excel"){
        $vbproj = $officeApp.workbooks(1).vbProject
    }


    return $vbproj
}

function GetCodeModule($vbproj, $moduleName) {

    $codeModule = $vbproj.VBComponents($moduleName).CodeModule

    return $codeModule

}

function GetCode($codeModule){

    return $codeModule.lines(1,$codeModule.CountOfLines)

}

function RemoveCode($codeModule){

    return $codeModule.DeleteLines(1,$codeModule.CountOfLines)

}

function ExportCode($codeModule, $path){
  $COMPONENT_TYPE_MODULE = 1
  $COMPONENT_TYPE_CLASS = 2
  $COMPONENT_TYPE_FORM = 3
  $COMPONENT_TYPE_SPECIAL = 100

    switch($codeModule.Parent.Type){
        $COMPONENT_TYPE_FORM {$suffix = '.frm'}
        $COMPONENT_TYPE_CLASS {$suffix = '.cls'}
        $COMPONENT_TYPE_MODULE {$suffix = '.bas'}
        $COMPONENT_TYPE_SPECIAL {$suffix = '.cls'}
        default{1}
    }

    $moduleFilename = $codeModule.Name + $suffix

    $moduleDestination = [IO.Path]::Combine($path, $moduleFilename)

    $codeModule.Parent.Export($moduleDestination)

}

function RemoveCodeModule($vbProj,$codeModule){

    $vbProj.VBComponents.Remove($codeModule.Parent)
}

function ImportCode($vbProj, $path){
  $COMPONENT_TYPE_MODULE = 1
  $COMPONENT_TYPE_CLASS = 2
  $COMPONENT_TYPE_FORM = 3
  $COMPONENT_TYPE_SPECIAL = 100

    $moduleName = (Get-Item $path).Basename
    Write-Debug $moduleName

    #check if component exists
    $component = $null
    $componentType = -1
    try{
        $component = $vbProj.VBComponents($moduleName)
        switch($component.Type){
            $COMPONENT_TYPE_FORM {$componentType = 1}
            $COMPONENT_TYPE_CLASS {$componentType = 1}
            $COMPONENT_TYPE_MODULE {$componentType = 1}
            $COMPONENT_TYPE_SPECIAL {$componentType = 2}
            default{1}
        }
    }catch{

    }
    #special modules like sheets,workbooks,accessforms

    #exists normal - remove old
    if ($componentType -eq 1){
        RemoveCodeModule $vbProj $component.CodeModule
    }

    #import code into
    $newComponent = $vbProj.VBComponents.Import($path)

    #exists special
    if ($componentType -eq 2){
        $curModule = $component.CodeModule
        $newModule = $newComponent.CodeModule

        $newCode = GetCode $newModule
        # $vbProj.VBComponents.Remove($newComponent)


        RemoveCodeModule $vbProj $newModule
        $newModule = $curModule

        RemoveCode $curModule
        $curModule.AddFromString($newCode)
    }

    return $newModule

}

function RepoChanged($dbFile,$ExportLocation,$dteChange) {

    Write-Information "repo changed $dteChange" -InformationAction Continue

    $filesAll=Get-ChildItem -Path "${ExportLocation}*.*"
    $filesChanged = $filesAll | Where-Object {$_.LastWriteTime -gt $dteChange}

    # $filesChanged
    $m = $filesChanged | measure
    $m = $filesAll | measure

    Write-Information "repo changed $filesChanged.Count" -InformationAction Continue
    # Write-Host "RepoChanged"

    #loop all changed files
    $filesChanged | ForEach-Object {
        $code = $_ | Get-Content
        $name = $_.Name

        Write-Information "repo changed $name - $code" -InformationAction Continue

    }

    $accessRun = GetApp $dbFile

    if(!$accessRun) {
        #     $modules.add($module.Name, $modDate)
        Write-Information 'file closed' -InformationAction Continue
    }else{
        #     if($modDate -eq $modLog){
        #
        Write-Information "$accessRun is running" -InformationAction Continue

        # $accessRun.LoadFromText(5, "Testing", $ExportLocation+"Testing.txt")
        #
        #     }else{
        #         # Write-Information "$modDate is newer will be update" -InformationAction Continue
        #
        #     }
    }


}

