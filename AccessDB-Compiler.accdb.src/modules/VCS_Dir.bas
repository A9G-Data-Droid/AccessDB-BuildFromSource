Option Compare Database
Option Private Module
Option Explicit

' Path/Directory of the current database file.
Public Function VCS_ProjectPath() As String
    VCS_ProjectPath = CurrentProject.path
    If Right$(VCS_ProjectPath, 1) <> "\" Then VCS_ProjectPath = VCS_ProjectPath & "\"
End Function

' Create folder `Path`. Silently do nothing if it already exists.
Public Sub VCS_MkDirIfNotExist(ByVal path As String)
    On Error GoTo ErrorHandler
    MkDir path
ErrorHandler:
End Sub

' Delete a file if it exists.
Public Sub VCS_DelIfExist(ByVal path As String)
    On Error GoTo ErrorHandler
    Kill path
ErrorHandler:
End Sub

' Erase all *.`ext` files in `Path`.
Public Sub VCS_ClearTextFilesFromDir(ByVal path As String, ByVal Ext As String)
    If Not FSO.FolderExists(path) Then Exit Sub

    On Error GoTo ErrorHandler
    If Dir$(path & "*." & Ext) <> vbNullString Then
        FSO.DeleteFile path & "*." & Ext
    End If
    
ErrorHandler:
End Sub

Public Function VCS_FileExists(ByVal strPath As String) As Boolean
    On Error GoTo ErrorHandler
    VCS_FileExists = False
    VCS_FileExists = ((GetAttr(strPath) And vbDirectory) <> vbDirectory)

ErrorHandler:
End Function