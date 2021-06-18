Attribute VB_Name = "Testing"
Option Compare Database

Sub qryx()


    Dim db As Database
    
    Set db = CurrentDb

    Dim qdf As QueryDef
    
    Set qdf = db.QueryDefs("UpdateStage")
    
    
    qdf.Parameters(0).Value = 4
    qdf.Parameters(1).Value = 4

    On Error Resume Next
    qdf.Execute (dbFailOnError)

    Debug.Print DBEngine.Errors(0).Description


    Debug.Print qdf.RecordsAffected

End Sub

Sub qryx2()

    Dim db As Database
    Set db = CurrentDb

    Dim qdf As QueryDef
    Set qdf = db.QueryDefs("SaveTitle")
    
    Set par = qdf.Parameters(0)
    With par
        Debug.Print .Type, .Name
        .Value = "someText2"
    End With
    
    Set par = qdf.Parameters(1)
    With par
        Debug.Print .Type, .Name
        .Value = 23
    End With
    
    On Error Resume Next
    Debug.Print DBEngine.Errors.Count
    qdf.Execute (dbFailOnError)

    Debug.Print DBEngine.Errors.Count
    Debug.Print DBEngine.Errors(0).Description
    Debug.Print qdf.RecordsAffected

End Sub

