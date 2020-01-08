Option Compare Database
Option Private Module
Option Explicit

Public Sub VCS_ImportRelation(ByVal filePath As String, Optional appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim thisDB As Database
    Set thisDB = appInstance.CurrentDb
    Dim inputFile As Object
    Set inputFile = FSO.OpenTextFile(filePath, iomode:=ForReading, create:=False, Format:=TristateFalse)

    Dim newRelation As Relation
    Set newRelation = thisDB.CreateRelation
    With newRelation
        .Attributes = inputFile.ReadLine
        .Name = GetRelationRelativeName(inputFile.ReadLine)
        .Table = inputFile.ReadLine
        .ForeignTable = inputFile.ReadLine
    End With
    
    Dim newField As Field
    Do Until inputFile.AtEndOfStream
        If "Field = Begin" = inputFile.ReadLine Then
            Set newField = newRelation.CreateField(inputFile.ReadLine) ' Name set here
            'newField.Name = inputFile.ReadLine
            newField.ForeignName = inputFile.ReadLine
            If "End" <> inputFile.ReadLine Then
                Set newField = Nothing
                Err.Raise 40000, "VCS_ImportRelation", "Missing 'End' for a 'Begin' in " & filePath
            End If
            
            newRelation.Fields.Append newField
            Set newField = Nothing
        End If
    Loop
    
    inputFile.Close
    
    ' Skip if relationship already exists and make a note of it. It was embedded in the table schema.
    On Error GoTo ErrorHandler
    thisDB.Relations.Append newRelation
    
ErrorHandler:
    Select Case Err.Number
    Case 0
    Case 3012
        If DebugOutput Then Debug.Print "Relationship already exists: """ & newRelation.Name & """ "
    Case Else
        Debug.Print "Failed to add: """ & newRelation.Name & """ " & Err.Number & " " & Err.Description
    End Select
End Sub

' Remove linked table from relation because that table is now a local object
Public Function GetRelationRelativeName(ByVal RelationName As String) As String
    If InStr(1, RelationName, "].") > 0 Then
        ' Need to remove path to linked file
        GetRelationRelativeName = CStr(Split(RelationName, "].")(1))
    Else
        GetRelationRelativeName = RelationName
    End If
End Function