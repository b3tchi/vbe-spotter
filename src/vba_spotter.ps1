Import-Module $PSScriptRoot/lib_spotter.ps1

function GetApp($scriptPath) {
  [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")

  # $TargetApp = $Access.Run("GetApp", [ref]$scriptPath) #use [ref] for optinal COM parameters
  $TargetApp = [Microsoft.VisualBasic.Interaction]::GetObject($scriptPath)

  return $TargetApp
}

Function AccChanged{ #($access, $dbFile) {

    Write-Information 'acc change' -InformationAction Continue
    # this is the bit we want to happen when the file changes
    # Clear-Host # remove previous console output
    # & 'C:\Program Files\erl7.3\bin\erlc.exe' 'program.erl' # compile some erlang
    # erl -noshell -s program start -s init stop # run the compiled erlang program:start()

    Copy-Item $dbFile -Destination $ExportLocation
    $FileName = Split-Path $dbFile -Leaf
    $access.OpenCurrentDatabase($ExportLocation+$FileName)
    $db = $access.CurrentDb

    # $dbModified = (Get-Item $dbfile).LastWriteTime

    $conts = $access.CurrentProject.AllModules

    $vbproj = $access.VBE.VBProjects(1)

    foreach($module in $conts){

        # Write-Host $module.Name $module.DateModified
        # $modDate = $module.DateModified.ToString('yyyy-mm-dd HH:mm:ss')

        # $moduleLog = $modules | Where-Object { $_.psobject.module -eq $module.Name}

        # Write-Information ($modules | ConvertTo-JSON) -InformationAction Continue
        # $modLog = $modules[$module.name]
        # $modLog = $modules['m_FormFx']

        # Write-Information $modLog 'log' -InformationAction Continue

        # if(!$modLog) {
        #     $modules.add($module.Name, $modDate)
        #     # Write-Information 'adding module null' -InformationAction Continue
        # }else{
        #     if($modDate -eq $modLog){
        #
        #         # Write-Information "$modDate is same nothing to do" -InformationAction Continue
        #
        #     }else{
        #         # Write-Information "$modDate is newer will be update" -InformationAction Continue
        #
        #     }
        # }

        ##main function
        $access.SaveAsText(5, $module.Name, $ExportLocation +  $module.Name + ".txt")

        # Write-Information $module -InformationAction Continue

        # $codeModule = $vbproj.VBComponents($module.Name).CodeModule.lines(0, $codeModule.CountOfLines)
        $codeModule = $vbproj.VBComponents($module.Name).CodeModule

        $code = $codeModule.lines(1,$codeModule.CountOfLines)

        # Write-Information $code -InformationAction Continue

        # $code = $codeModule.lines(0,$codeModule.CountOfLines)
        # Write-Information $codeModule -InformationAction Continue


        # Write-Host $modLog -eq $null

        # Write-Host $moduleLog 'item'

        # $access.SaveAsText(5, $module.Name, $ExportLocation +  $module.Name + ".txt")

        # $object = New-Object -TypeName PSObject
        # $object | Add-Member -Name 'module' -MemberType Noteproperty -Value $module.Name
        # $object | Add-Member -Name 'modified' -MemberType Noteproperty -Value $modDate
        # $modules += $object

        # $modules += @{$module.Name=$modDate}
        # $modules.add($module.Name, $modDate)
    }

    # Write-Information ($modules | ConvertTo-JSON) -InformationAction Continue

    # $modules | ConvertTo-JSON | Out-File $ExportLocation"modules.json"
    # }

    # $dbModified = (Get-Item $dbFile).LastWriteTime
    # Write-Information "file changed" -InformationAction Continue
    $access.CloseCurrentDatabase()

    # return $modules

}

Function RepoChanged() {

    Write-Information 'repo change' -InformationAction Continue
    # Write-Host "RepoChanged"
    $accessRun = GetApp $dbFile

    if(!$accessRun) {
        #     $modules.add($module.Name, $modDate)
        Write-Information 'file closed' -InformationAction Continue

    }else{
        #     if($modDate -eq $modLog){
        #
        Write-Information "$accessRun is running" -InformationAction Continue

        $accessRun.LoadFromText(5, "Testing", $ExportLocation+"Testing.txt")
        #
        #     }else{
        #         # Write-Information "$modDate is newer will be update" -InformationAction Continue
        #
        #     }
    }



}

Function Watch{#($access, $dbFile) {
    $global:FileChanged = 0 # dirty... any better suggestions?

    Write-Information $dbFile -InformationAction Continue

    $FilePath = Split-Path $dbFile -Parent
    $FileName = Split-Path $dbFile -Leaf
    # $folder = "M:\dev\Erlang"
    # $filter = "*.erl"
    $watcherAcc = New-Object IO.FileSystemWatcher $FilePath, $FileName -Property @{
        IncludeSubdirectories = $false
        EnableRaisingEvents = $true
    }

    Register-ObjectEvent $watcherAcc "Changed" -Action {$global:FileChanged = 1} > $null

    $watcherRepo = New-Object IO.FileSystemWatcher $ExportLocation, "*.txt" -Property @{
        IncludeSubdirectories = $false
        EnableRaisingEvents = $true
    }

    Register-ObjectEvent $watcherRepo "Changed" -Action {$global:FileChanged = 2} > $null

    # $dbModified = (Get-Item $dbfile).LastWriteTime
    # $dbActual = (Get-Item $dbfile).LastWriteTime

    while ($true){
        while ($FileChanged -eq 0){
            # We need this to block the IO thread until there is something to run
            # so the script doesn't finish. If we call the action directly from
            # the event it won't be able to write to the console
            Start-Sleep -Milliseconds 500
            # $dbActual = (Get-Item $dbfile).LastWriteTime

            # Write-Host 'loop-check-'+$FileChanged+$dbModified+$dbActual

            # if ($dbActual -gt $dbModified){
            #     $FileChanged = $true
            #     Write-Host 'changed-'+$dbModified+$dbActual
            #     $dbModified = $dbActual
            #
            #     # Write-Host 'loop-lt'+$FileChanged
            # }

        }

        $localChanged = $global:FileChanged
        Write-Host 'loop-runevent'+$localChanged

        # a file has changed, run our stuff on the I/O thread so we can see the output
        if($localChanged -eq 1){
            AccChanged # $access $dbFile        # reset and go again
            Write-Host 'acc if'+$localChanged
        }

        if($localChanged -eq 2){
            RepoChanged
            Write-Host 'repo if'+$localChanged
        }

        $global:FileChanged = 0

        Write-Host 'loop-run'+$global:FileChanged




        #reregister action
        # Register-ObjectEvent $Watcher "Changed" -Action {$global:FileChanged = $true} > $null

    }
}

# [string]$ExportLocation ="C:\Users\czJaBeck\Repositories\PPM_vba_codes\"
# [string]$dbFile = "C:\Users\czJaBeck\Onedrive - LEGO\Documents\Wdd_v2.accdb"

[string]$dbFile = "~/Repositories/VbaSpotter/src/TestDb.accdb"
[string]$ExportLocation ="~/Repositories/VbaSpotter/test/codes/"

[Hashtable]$modules= @{}

try {

    $access = New-Object -ComObject Access.Application
    # $access.visible = $true

    Watch # $access $GdbFile
    # RepoChanged
    # AccChanged

} finally {

    #execute where broken
    Write-Information "Clearing running access" -InformationAction Continue
    $access.Quit()
}

