$w=[System.Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
$w.documents |ft name,path -auto