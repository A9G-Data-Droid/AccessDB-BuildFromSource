Option Compare Database
Option Private Module
Option Explicit

#If VBA7 Then

Private Declare PtrSafe _
        Function getTempPath Lib "kernel32" _
        Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                              ByVal lpBuffer As String) As Long
Private Declare PtrSafe _
        Function getTempFileName Lib "kernel32" _
        Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                                  ByVal lpPrefixString As String, _
                                  ByVal wUnique As Long, _
                                  ByVal lpTempFileName As String) As Long
#Else
Private Declare _
        Function getTempPath Lib "kernel32" _
        Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                              ByVal lpBuffer As String) As Long
Private Declare _
        Function getTempFileName Lib "kernel32" _
        Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                                  ByVal lpPrefixString As String, _
                                  ByVal wUnique As Long, _
                                  ByVal lpTempFileName As String) As Long
#End If



' Keep a persistent reference to file system object after initializing version control.
Private m_FSO As Object
Public Function FSO() As Object
    If m_FSO Is Nothing Then Set m_FSO = CreateObject("Scripting.FileSystemObject")
    Set FSO = m_FSO
End Function



'---------------------------------------------------------------------------------------
' Procedure : ConvertUcs2Utf8
' Author    : Adam Waller
' Date      : 1/23/2019
' Purpose   : Convert a UCS2-little-endian encoded file to UTF-8.
'---------------------------------------------------------------------------------------
'
Public Sub ConvertUcs2Utf8(strSourceFile As String, strDestinationFile As String)

    Dim stmNew As Object
    Set stmNew = CreateObject("ADODB.Stream")
    Dim strText As String
    
    ' Read file contents
    With FSO.OpenTextFile(strSourceFile, , , TristateTrue)
        strText = .ReadAll
        .Close
    End With
    
    ' Write as UTF-8
    With stmNew
        .Open
        .Type = 2 'adTypeText
        .Charset = "utf-8"
        .WriteText strText
        .SaveToFile strDestinationFile, 2 'adSaveCreateOverWrite
        .Close
    End With
    
    Set stmNew = Nothing
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ConvertUtf8Ucs2
' Author    : Adam Waller
' Date      : 1/24/2019
' Purpose   : Convert the file to old UCS-2 unicode format
'---------------------------------------------------------------------------------------
'
Public Sub ConvertUtf8Ucs2(strSourceFile As String, strDestinationFile As String)

    Dim stmNew As Object
    Set stmNew = CreateObject("ADODB.Stream")
    Dim strText As String
    
    ' Read file contents
    With FSO.OpenTextFile(strSourceFile, , , TristateTrue)
        strText = .ReadAll
        .Close
    End With
    
    ' Write as UCS-2 LE (BOM)
    With stmNew
        .Open
        .Type = 2 'adTypeText
        .Charset = "unicode"  ' The original Windows "Unicode" was UCS-2
        .WriteText strText
        .SaveToFile strDestinationFile, 2  'adSaveCreateOverWrite
        .Close
    End With
    
    Set stmNew = Nothing
    
End Sub



' Determine if this database imports/exports code as UCS-2-LE. (Older file
' formats cause exported objects to use a Windows 8-bit character set.)
Public Function UsingUcs2(Optional ByRef appInstance As Application) As Boolean
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim obj_name As String
    Dim obj_type As Variant
    Dim fn As Long
    Dim bytes As String
    Dim obj_type_split() As String
    Dim obj_type_name As String
    Dim obj_type_num As Long
    Dim thisDb As Database
    Set thisDb = appInstance.CurrentDb

    If CurrentProject.ProjectType = acMDB Then
        If thisDb.QueryDefs.Count > 0 Then
            obj_type_num = acQuery
            obj_name = thisDb.QueryDefs.Item(0).Name
        Else
            For Each obj_type In Split( _
                "Forms|" & acForm & "," & _
                "Reports|" & acReport & "," & _
                "Scripts|" & acMacro & "," & _
                "Modules|" & acModule _
            )
                DoEvents
                obj_type_split = Split(obj_type, "|")
                obj_type_name = obj_type_split(0)
                obj_type_num = Val(obj_type_split(1))
                If thisDb.Containers(obj_type_name).Documents.Count > 0 Then
                    obj_name = thisDb.Containers(obj_type_name).Documents(0).Name
                    Exit For
                End If
            Next
        End If
    Else
        ' ADP Project
        If CurrentData.AllQueries.Count > 0 Then
            obj_type_num = acServerView
            obj_name = CurrentData.AllQueries(1).Name
        ElseIf CurrentProject.AllForms.Count > 0 Then
            ' Try a form
            obj_type_num = acForm
            obj_name = CurrentProject.AllForms(1).Name
        Else
            ' Can add more object types as needed...
        End If
    End If

    If obj_name = vbNullString Then
        ' No objects found that can be used to test UCS2 versus UTF-8
        UsingUcs2 = True
        Exit Function
    End If

    Dim tempFileName As String: tempFileName = GetTempFile()
    Application.SaveAsText obj_type_num, obj_name, tempFileName
    
    bytes = "  "
    fn = FreeFile
    Open tempFileName For Binary Access Read As fn
    Get fn, 1, bytes
    Close fn
    FSO.DeleteFile tempFileName
    
    ' If the 16-bit units use little-endian order, the BOM will appear in the sequence of bytes as 0xFF 0xFE
    UsingUcs2 = (Asc(Mid$(bytes, 1, 1)) = &HFF) And (Asc(Mid$(bytes, 2, 1)) = &HFE)
End Function



'---------------------------------------------------------------------------------------
' Procedure : GetTempFile
' Author    : Adapted by Adam Waller
' Date      : 1/23/2019
' Purpose   : Generate Random / Unique temporary file name.
'---------------------------------------------------------------------------------------
'
Public Function GetTempFile(Optional strPrefix As String = "VBA") As String

    Dim strPath As String * 512
    Dim strName As String * 576
    Dim lngReturn As Long
    
    lngReturn = getTempPath(512, strPath)
    lngReturn = getTempFileName(strPath, strPrefix, 0, strName)
    If lngReturn <> 0 Then GetTempFile = Left$(strName, InStr(strName, vbNullChar) - 1)
    
End Function



'---------------------------------------------------------------------------------------
' Procedure : WriteFile
' Author    : Adam Waller
' Date      : 1/23/2019
' Purpose   : Save string variable to text file.
'---------------------------------------------------------------------------------------
'
Public Sub WriteFile(strContent As String, strPath As String, Optional blnUnicode As Boolean = False)
    With FSO.CreateTextFile(strPath, True, blnUnicode)
        .Write strContent
        .Close
    End With
End Sub



' Test shows that UCS-2 files exported by Access make round trip through our conversions.
Public Sub TestTextModes()
    Dim tempFileName As String
    tempFileName = FileAccess.GetTempFile()
    
    Application.SaveAsText acQuery, CurrentDb.QueryDefs.Item(0).Name, tempFileName
    
    ConvertUtf8Ucs2 tempFileName, tempFileName & "UCS2UCS"
    ConvertUcs2Utf8 tempFileName, tempFileName & "UTF8"
    ConvertUtf8Ucs2 tempFileName, tempFileName & "UTF82UCS"
End Sub