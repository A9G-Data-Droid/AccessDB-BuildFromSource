Option Compare Database
Option Explicit

Private Sub writeTextToFile(ByVal textToWrite As String, ByVal file_path As String)
    Dim oFile As Object
    Set oFile = FSO.CreateTextFile(file_path)

    oFile.WriteLine textToWrite
    oFile.Close
End Sub

Private Function readFromTextFile(ByVal file_path As String) As String
    Dim oFile As Object
    Set oFile = FSO.OpenTextFile(file_path, ForReading)
       
    readFromTextFile = oFile.ReadAll
    oFile.Close
End Function

Public Sub ImportQueryFromSQL(ByVal obj_name As String, ByVal file_path As String, _
                              Optional ByVal Ucs2Convert As Boolean = False, Optional theDatabase As Database)
    On Error GoTo ErrorHandler
    If theDatabase Is Nothing Then Set theDatabase = CurrentDb
    
    If Not VCS_Dir.VCS_FileExists(file_path) Then Exit Sub
    
    If Ucs2Convert Then
        Dim tempFileName As String
        tempFileName = FileAccess.GetTempFile()
        FileAccess.ConvertUtf8Ucs2 file_path, tempFileName
        theDatabase.QueryDefs.Delete (obj_name)
        theDatabase.CreateQueryDef obj_name, readFromTextFile(file_path)
        
        FSO.DeleteFile tempFileName
    Else
        theDatabase.QueryDefs.Delete (obj_name)
        theDatabase.CreateQueryDef obj_name, readFromTextFile(file_path)
    End If

ErrorHandler:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Sub