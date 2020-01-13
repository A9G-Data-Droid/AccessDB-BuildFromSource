Option Compare Database
Option Explicit
'--------------------------------------------------------------------
' Complier to Make and Build Access DB from source files.
'   by Adam Kauffman on 2020-01-07
'
'   Designed to work with code exported by msaccess-vcs-integration
'       https://github.com/joyfullservice/msaccess-vcs-integration
'       https://github.com/timabell/msaccess-vcs-integration
'--------------------------------------------------------------------
Const Version = "0.2.0"

' Keep a persistent reference to file system object after initializing version control.
Private m_FSO As Object
Public Function FSO() As Object
    If m_FSO Is Nothing Then Set m_FSO = CreateObject("Scripting.FileSystemObject")
    Set FSO = m_FSO
End Function

' This is the main entry point called by the Compiler GUI
Public Sub Build(ByVal sourceFolder As String, ByVal outputFile As String, Optional ByVal overwrite As Boolean = False)
    Const cstrSpacer As String = "-------------------------------"
    Dim startTime As Single
    startTime = Timer
    Debug.Print  'String(32, vbNewLine)
    Debug.Print cstrSpacer
    Debug.Print " -= " & CurrentProject.Name & " =-"
    Debug.Print " Version: " & Version
    Debug.Print " Started: " & Now
    Debug.Print
    Debug.Print cstrSpacer
    Debug.Print " Source Folder:"
    Debug.Print sourceFolder
    If overwrite Then DestoryDB outputFile
    
    Dim newApp As Application
    Set newApp = New Access.Application
    newApp.NewCurrentDatabase outputFile
    On Error GoTo ErrorHandler
    Debug.Print " Created new DB: "
    Debug.Print newApp.CurrentDb.Name
    Debug.Print cstrSpacer
    Debug.Print
    If DebugOutput Then newApp.Visible = True

    ImportAllSource False, sourceFolder, newApp
    Debug.Print
    Debug.Print " Runtime: " & Round(Timer - startTime, 2) & " seconds"
    
ErrorHandler:
    If Err.Number > 0 Then
        Debug.Print "Error: " & Err.Number & " " & Err.Description
    End If

    newApp.Quit acQuitSaveAll
End Sub

' Used by overwrite to delete the DB before it is created.
Public Sub DestoryDB(ByVal dbFullPath As String)
    On Error GoTo ErrorHandler
    
    Dim theFile As Object
    Set theFile = FSO.GetFile(dbFullPath)
    theFile.Delete
    Debug.Print "Deleted DB: " & dbFullPath
    
ErrorHandler:
    If Err.Number = 53 Then
        ' TODO: Handle a filename given with no extension
        Dim thisFileName As String
        thisFileName = FSO.GetFileName(dbFullPath)
        If InStr(1, thisFileName, ".") = 0 Then
            ' This is too dangerous. Let's not.
            thisFileName = Dir(dbFullPath & ".*")
        End If
    ElseIf Err.Number > 0 Then
        Debug.Print Err.Number & " " & Err.Description
        Debug.Assert Err.Number = 0
    End If
End Sub