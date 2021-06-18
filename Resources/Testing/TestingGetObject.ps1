    [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
    
    $Path = "C:\Users\czJaBeck\Documents\Vbox\LocalWeb_Ps\TestDb.accdb"
    $Access = [Microsoft.VisualBasic.Interaction]::GetObject("$Path")

    $Access.Run('Test')