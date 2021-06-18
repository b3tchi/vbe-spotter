Option Compare Database

Sub ShowUserRosterMultipleUsers()

  # Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
  # Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
  # Dim i, j As Long

 $cnstring = @'
 Provider=Microsoft.ACE.OLEDB.12.0
 ;User ID=Admin
 ;Data Source=C:\Users\czJaBeck\Onedrive - LEGO\Documents\Wdd_v2.accdb
 ;Mode=Share Deny None;Extended Properties=""
 ;Jet OLEDB:System database=C:\Users\czJaBeck\AppData\Roaming\Microsoft\Access\System.mdw
 ;Jet OLEDB:Registry Path=Software\Microsoft\Office\16.0\Access\Access Connectivity Engine
 ;Jet OLEDB:Database Password=""
 ;Jet OLEDB:Engine Type=6
 ;Jet OLEDB:Database Locking Mode=1
 ;Jet OLEDB:Global Partial Bulk Ops=2
 ;Jet OLEDB:Global Bulk Transactions=1
 ;Jet OLEDB:New Database Password=""
 ;Jet OLEDB:Create System Database=False
 ;Jet OLEDB:Encrypt Database=False
 ;Jet OLEDB:Don't Copy Locale on Compact=False
 ;Jet OLEDB:Compact Without Replica Repair=False
 ;Jet OLEDB:SFP=False
 ;Jet OLEDB:Support Complex Data=True
 ;Jet OLEDB:Bypass UserInfo Validation=False
 ;Jet OLEDB:Limited DB Caching=False
 ;Jet OLEDB:Bypass ChoiceField Validation=False
'@


  $cn = cn.OpenConnection($cnstring)

  # ' The user roster is exposed as a provider-specific schema rowset
  # ' in the Jet 4.0 OLE DB provider.  You have to use a GUID to
  # ' reference the schema, as provider-specific schemas are not
  # ' listed in ADO's type library for schema rowsets

  $adSchemaProviderSpecific = -1

  $rs = cn.OpenSchema($adSchemaProviderSpecific, $null  , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")

  'Output the list of all users in the current database.

  Debug.Print rs.Fields(0).Name, "", rs.Fields(1).Name, _
  "", rs.Fields(2).Name, rs.Fields(3).Name

  While Not rs.EOF
    Debug.Print rs.Fields(0), rs.Fields(1), _
    rs.Fields(2), rs.Fields(3)
    ListUsers.AddItem "'" & rs.Fields(0) & "-" & rs.Fields(1) & "'"
    rs.MoveNext
  Wend

End Sub