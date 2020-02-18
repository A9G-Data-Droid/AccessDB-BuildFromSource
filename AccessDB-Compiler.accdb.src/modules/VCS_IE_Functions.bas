Option Compare Database
Option Private Module
Option Explicit

' Constants for Scripting.FileSystemObject API
Public Const ForReading = 1
Public Const ForWriting = 2
Public Const ForAppending = 8
Public Const TristateTrue = -1
Public Const TristateFalse = 0
Public Const TristateUseDefault = -2

' Export a database object with optional UCS2-to-UTF-8 conversion.
Public Sub VCS_ExportObject(ByVal obj_type_num As Integer, ByVal obj_name As String, _
                            ByVal file_path As String, Optional ByVal Ucs2Convert As Boolean = False, Optional ByRef appInstance As Application)
                    
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    VCS_Dir.VCS_MkDirIfNotExist Left$(file_path, InStrRev(file_path, "\"))
    If Ucs2Convert Then
        Dim tempFileName As String
        tempFileName = FileAccess.GetTempFile()
        appInstance.SaveAsText obj_type_num, obj_name, tempFileName
        If obj_type_num = acModule Then
            With FSO.OpenTextFile(tempFileName, ForAppending, False, TristateTrue)
                .WriteBlankLines 1
            End With
        End If
        
        FileAccess.ConvertUcs2Utf8 tempFileName, file_path
    Else
        appInstance.SaveAsText obj_type_num, obj_name, file_path
        If obj_type_num = acModule Then
            With FSO.OpenTextFile(file_path, ForAppending, False, TristateTrue)
                .WriteBlankLines 1
            End With
        End If
    End If
End Sub

' Import a database object with optional UTF-8-to-UCS2 conversion.
Public Sub VCS_ImportObject(ByVal obj_type_num As Integer, ByVal obj_name As String, _
                            ByVal file_path As String, Optional ByVal Ucs2Convert As Boolean = False, Optional ByRef appInstance As Application)
                    
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    If Not VCS_Dir.VCS_FileExists(file_path) Then Exit Sub
    
    If Ucs2Convert Then
        Dim tempFileName As String
        tempFileName = FileAccess.GetTempFile()
        FileAccess.ConvertUtf8Ucs2 file_path, tempFileName
        appInstance.LoadFromText obj_type_num, obj_name, tempFileName
        
        FSO.DeleteFile tempFileName
    Else
        appInstance.LoadFromText obj_type_num, obj_name, file_path
    End If
End Sub