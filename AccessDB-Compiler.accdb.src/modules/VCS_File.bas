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

' --------------------------------
' Structures
' --------------------------------

' Structure to track buffered reading or writing of binary files
Private Type BinFile
    file_num As Integer
    file_len As Long
    file_pos As Long
    buffer As String
    buffer_len As Integer
    buffer_pos As Integer
    at_eof As Boolean
    mode As String
End Type

' --------------------------------
' Basic functions missing from VB 6: buffered file read/write, string builder, encoding check & conversion
' --------------------------------

' Open a binary file for reading (mode = 'r') or writing (mode = 'w').
Private Function BinOpen(ByVal file_path As String, ByVal mode As String) As BinFile
    Dim f As BinFile

    f.file_num = FreeFile
    f.mode = LCase$(mode)
    If f.mode = "r" Then
        Open file_path For Binary Access Read As f.file_num
        f.file_len = LOF(f.file_num)
        f.file_pos = 0
        If f.file_len > &H4000 Then
            f.buffer = String$(&H4000, " ")
            f.buffer_len = &H4000
        Else
            f.buffer = String$(f.file_len, " ")
            f.buffer_len = f.file_len
        End If
        f.buffer_pos = 0
        Get f.file_num, f.file_pos + 1, f.buffer
    Else
        VCS_DelIfExist file_path
        Open file_path For Binary Access Write As f.file_num
        f.file_len = 0
        f.file_pos = 0
        f.buffer = String$(&H4000, " ")
        f.buffer_len = 0
        f.buffer_pos = 0
    End If

    BinOpen = f
End Function

' Buffered read one byte at a time from a binary file.
Private Function BinRead(ByRef f As BinFile) As Integer
    If f.at_eof = True Then
        BinRead = 0
        Exit Function
    End If

    BinRead = Asc(Mid$(f.buffer, f.buffer_pos + 1, 1))

    f.buffer_pos = f.buffer_pos + 1
    If f.buffer_pos >= f.buffer_len Then
        f.file_pos = f.file_pos + &H4000
        If f.file_pos >= f.file_len Then
            f.at_eof = True
            Exit Function
        End If
        If f.file_len - f.file_pos > &H4000 Then
            f.buffer_len = &H4000
        Else
            f.buffer_len = f.file_len - f.file_pos
            f.buffer = String$(f.buffer_len, " ")
        End If
        f.buffer_pos = 0
        Get f.file_num, f.file_pos + 1, f.buffer
    End If
End Function

' Buffered write one byte at a time from a binary file.
Private Sub BinWrite(ByRef f As BinFile, b As Integer)
    Mid(f.buffer, f.buffer_pos + 1, 1) = Chr$(b)
    f.buffer_pos = f.buffer_pos + 1
    If f.buffer_pos >= &H4000 Then
        Put f.file_num, , f.buffer
        f.buffer_pos = 0
    End If
End Sub

' Close binary file.
Private Sub BinClose(ByRef f As BinFile)
    If f.mode = "w" And f.buffer_pos > 0 Then
        f.buffer = Left$(f.buffer, f.buffer_pos)
        Put f.file_num, , f.buffer
    End If
    Close f.file_num
End Sub

' Binary convert a UCS2-little-endian encoded file to UTF-8.
'---------------------------------------------------------------------------------------
' Procedure : ConvertUcs2Utf8
' Author    : Adam Waller
' Date      : 1/23/2019
' Purpose   : Convert the file to unicode format
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

' Binary convert a UTF-8 encoded file to UCS2-little-endian.
Public Sub VCS_ConvertUtf8Ucs2(ByVal Source As String, ByVal dest As String)
    Dim f_in As BinFile
    Dim f_out As BinFile
    Dim in_1 As Integer
    Dim in_2 As Integer
    Dim in_3 As Integer

    f_in = BinOpen(Source, "r")
    f_out = BinOpen(dest, "w")

    Do While Not f_in.at_eof
        in_1 = BinRead(f_in)
        If (in_1 And &H80) = 0 Then
            ' U+0000 - U+007F   0LLLLLLL
            BinWrite f_out, in_1
            BinWrite f_out, 0
        ElseIf (in_1 And &HE0) = &HC0 Then
            ' U+0080 - U+07FF   110HHHLL 10LLLLLL
            in_2 = BinRead(f_in)
            BinWrite f_out, ((in_1 And &H3) * &H40) + (in_2 And &H3F)
            BinWrite f_out, (in_1 And &H1C) / &H4
        Else
            ' U+0800 - U+FFFF   1110HHHH 10HHHHLL 10LLLLLL
            in_2 = BinRead(f_in)
            in_3 = BinRead(f_in)
            BinWrite f_out, ((in_2 And &H3) * &H40) + (in_3 And &H3F)
            BinWrite f_out, ((in_1 And &HF) * &H10) + ((in_2 And &H3C) / &H4)
        End If
    Loop

    BinClose f_in
    BinClose f_out
End Sub

' Determine if this database imports/exports code as UCS-2-LE. (Older file
' formats cause exported objects to use a Windows 8-bit character set.)
Public Function VCS_UsingUcs2(Optional ByRef appInstance As Application) As Boolean
    If appInstance Is Nothing Then Set appInstance = Application.Application
    Dim obj_name As String
    Dim obj_type As Variant
    Dim fn As Integer
    Dim bytes As String
    Dim obj_type_split() As String
    Dim obj_type_name As String
    Dim obj_type_num As Integer
    Dim thisDB As Database
    Set thisDB = appInstance.CurrentDb
    
    If thisDB.QueryDefs.Count > 0 Then
        obj_type_num = acQuery
        obj_name = thisDB.QueryDefs(0).Name
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
            If thisDB.Containers(obj_type_name).Documents.Count > 0 Then
                obj_name = thisDB.Containers(obj_type_name).Documents(0).Name
                Exit For
            End If
        Next
    End If

    If obj_name = vbNullString Then
        ' No objects found that can be used to test UCS2 versus UTF-8
        VCS_UsingUcs2 = True
        Exit Function
    End If

    Dim tempFileName As String
    tempFileName = VCS_File.VCS_TempFile()
    
    appInstance.SaveAsText obj_type_num, obj_name, tempFileName
    
    fn = FreeFile
    Open tempFileName For Binary Access Read As fn
    bytes = "  "
    Get fn, 1, bytes
    VCS_UsingUcs2 = Asc(Mid$(bytes, 1, 1)) = &HFF And Asc(Mid$(bytes, 2, 1)) = &HFE
    Close fn
    
    FSO.DeleteFile (tempFileName)
End Function

' Generate Random / Unique temporary file name.
Public Function VCS_TempFile(Optional ByVal sPrefix As String = "VBA") As String
    Dim sTmpPath As String * 512
    Dim sTmpName As String * 576
    Dim nRet As Long
    Dim sFileName As String
    
    nRet = getTempPath(512, sTmpPath)
    nRet = getTempFileName(sTmpPath, sPrefix, 0, sTmpName)
    If nRet <> 0 Then sFileName = Left$(sTmpName, InStr(sTmpName, vbNullChar) - 1)
    VCS_TempFile = sFileName
End Function