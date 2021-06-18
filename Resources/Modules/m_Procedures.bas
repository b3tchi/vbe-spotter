Attribute VB_Name = "m_Procedures"
Option Compare Database

Public Function p01_CreateItem() As Long
    
    'Execution
    Call CurrentDb.Execute("01_CreateItem_Append")
    Call CurrentDb.Execute("01_NewItemIDs")

    'Return Status
    CreateItem = 1
        

End Function




