Option Compare Database
Option Private Module
Option Explicit



Public Sub ImportRelation(ByRef filePath As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim thisDb As Database
    Set thisDb = appInstance.CurrentDb
    
    Dim fileLines() As String
    With FSO.OpenTextFile(filePath, iomode:=ForReading, create:=False, Format:=TristateFalse)
        fileLines = Split(.ReadAll, vbCrLf)
        .Close
    End With
    
    Dim newRelation As Relation
    Set newRelation = thisDb.CreateRelation(fileLines(1), fileLines(2), fileLines(3), fileLines(0))
    
    Dim newField As Field
    Dim thisLine As Long
    For thisLine = 4 To UBound(fileLines)
        If "Field = Begin" = fileLines(thisLine) Then
            thisLine = thisLine + 1
            Set newField = newRelation.CreateField(fileLines(thisLine))  ' Name set here
            thisLine = thisLine + 1
            newField.ForeignName = fileLines(thisLine)
            thisLine = thisLine + 1
            If "End" <> fileLines(thisLine) Then
                Set newField = Nothing
                Err.Raise 40000, "ImportRelation", "Missing 'End' for a 'Begin' in " & filePath
            End If
            
            newRelation.Fields.Append newField
        End If
    Next thisLine
        
    ' Remove conflicting Index entries because adding the relation creates new indexes causing "Error 3284 Index already exists"
    On Error Resume Next
    With thisDb
        .Relations.Delete fileLines(1)  ' Avoid 3012 Relationship already exists
        .TableDefs(fileLines(2)).Indexes.Delete fileLines(1)
        .TableDefs(fileLines(3)).Indexes.Delete fileLines(1)
    End With
    On Error GoTo ErrorHandler
    
    With thisDb.Relations
        .Append newRelation
    End With
    
ErrorHandler:
    Select Case Err.Number
    Case 0
    Case 3012
        Form_LogWindow.WriteError "Relationship already exists: """ & newRelation.Name & """ "
    Case 3284
        Form_LogWindow.WriteError "Index already exists for: """ & newRelation.Name & """ "
    Case Else
        Form_LogWindow.WriteError "Failed to add: """ & newRelation.Name & """ " & Err.Number & " " & Err.Description
    End Select
End Sub