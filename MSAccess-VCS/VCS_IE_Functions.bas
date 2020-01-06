Option Compare Database

Option Private Module
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Private Const AggressiveSanitize As Boolean = True
Private Const StripPublishOption As Boolean = True

' Constants for Scripting.FileSystemObject API
Public Const ForReading = 1, ForWriting = 2, ForAppending = 8
Public Const TristateTrue = -1, TristateFalse = 0, TristateUseDefault = -2


' Can we export without closing the form?

' Export a database object with optional UCS2-to-UTF-8 conversion.
Public Sub VCS_ExportObject(ByVal obj_type_num As Integer, ByVal obj_name As String, _
                    ByVal file_path As String, Optional ByVal Ucs2Convert As Boolean = False)

    VCS_Dir.VCS_MkDirIfNotExist Left$(file_path, InStrRev(file_path, "\"))
    If Ucs2Convert Then
        Dim tempFileName As String
        tempFileName = VCS_File.VCS_TempFile()
        Application.SaveAsText obj_type_num, obj_name, tempFileName
        If obj_type_num = acModule Then
            Open tempFileName For Append As #1
            Write #1,
            Close #1
        End If
        VCS_File.VCS_ConvertUcs2Utf8 tempFileName, file_path
    Else
        Application.SaveAsText obj_type_num, obj_name, file_path
        If obj_type_num = acModule Then
            Open file_path For Append As #1
            Write #1,
            Close #1
        End If
    End If
End Sub

' Import a database object with optional UTF-8-to-UCS2 conversion.
Public Sub VCS_ImportObject(ByVal obj_type_num As Integer, ByVal obj_name As String, _
                    ByVal file_path As String, Optional ByVal Ucs2Convert As Boolean = False)
    
    If Not VCS_Dir.VCS_FileExists(file_path) Then Exit Sub
    
    If Ucs2Convert Then
        Dim tempFileName As String
        tempFileName = VCS_File.VCS_TempFile()
        VCS_File.VCS_ConvertUtf8Ucs2 file_path, tempFileName
        Application.LoadFromText obj_type_num, obj_name, tempFileName
        
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        FSO.DeleteFile tempFileName
    Else
        Application.LoadFromText obj_type_num, obj_name, file_path
    End If
End Sub

'shouldn't this be SanitizeTextFile (Singular)?

' For each *.txt in `Path`, find and remove a number of problematic but
' unnecessary lines of VB code that are inserted automatically by the
' Access GUI and change often (we don't want these lines of code in
' version control).
Public Sub VCS_SanitizeTextFiles(ByVal Path As String, ByVal Ext As String)

    Dim counter As Integer
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    '
    '  Setup Block matching Regex.
    Dim rxBlock As Object
    Set rxBlock = CreateObject("VBScript.RegExp")
    rxBlock.ignoreCase = False
    '
    '  Match PrtDevNames / Mode with or  without W
    Dim srchPattern As String
    srchPattern = "PrtDev(?:Names|Mode)[W]?"
    If (AggressiveSanitize = True) Then
      '  Add and group aggressive matches
      srchPattern = "(?:" & srchPattern
      srchPattern = srchPattern & "|GUID|""GUID""|NameMap|dbLongBinary ""DOL"""
      srchPattern = srchPattern & ")"
    End If
    '  Ensure that this is the beginning of a block.
    srchPattern = srchPattern & " = Begin"
'Debug.Print srchPattern
    rxBlock.Pattern = srchPattern
    '
    '  Setup Line Matching Regex.
    Dim rxLine As Object
    Set rxLine = CreateObject("VBScript.RegExp")
    srchPattern = "^\s*(?:"
    srchPattern = srchPattern & "Checksum ="
    srchPattern = srchPattern & "|BaseInfo|NoSaveCTIWhenDisabled =1"
    If (StripPublishOption = True) Then
        srchPattern = srchPattern & "|dbByte ""PublishToWeb"" =""1"""
        srchPattern = srchPattern & "|PublishOption =1"
    End If
    srchPattern = srchPattern & ")"
'Debug.Print srchPattern
    rxLine.Pattern = srchPattern
    Dim fileName As String
    fileName = Dir$(Path & "*." & Ext)
    Dim isReport As Boolean
    isReport = False
    
    counter = 0
    Do Until Len(fileName) = 0
        DoEvents
        counter = counter + 1
        Dim obj_name As String
        obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)

        Dim InFile As Object
        Set InFile = FSO.OpenTextFile(Path & obj_name & "." & Ext, iomode:=ForReading, create:=False, Format:=TristateFalse)
        Dim OutFile As Object
        Set OutFile = FSO.CreateTextFile(Path & obj_name & ".sanitize", overwrite:=True, Unicode:=False)
    
        Dim getLine As Boolean
        getLine = True
        
        Do Until InFile.AtEndOfStream
            DoEvents
            Dim txt As String
            '
            ' Check if we need to get a new line of text
            If getLine = True Then
                txt = InFile.ReadLine
            Else
                getLine = True
            End If
            '
            ' Skip lines starting with line pattern
            If rxLine.Test(txt) Then
                Dim rxIndent As Object
                Set rxIndent = CreateObject("VBScript.RegExp")
                rxIndent.Pattern = "^(\s+)\S"
                '
                ' Get indentation level.
                Dim matches As Object
                Set matches = rxIndent.Execute(txt)
                '
                ' Setup pattern to match current indent
                Select Case matches.Count
                    Case 0
                        rxIndent.Pattern = "^" & vbNullString
                    Case Else
                        rxIndent.Pattern = "^" & matches(0).SubMatches(0)
                End Select
                rxIndent.Pattern = rxIndent.Pattern + "\s"
                '
                ' Skip lines with deeper indentation
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If Not rxIndent.Test(txt) Then Exit Do
                Loop
                ' We've moved on at least one line so do get a new one
                ' when starting the loop again.
                getLine = False
            '
            ' skip blocks of code matching block pattern
            ElseIf rxBlock.Test(txt) Then
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If InStr(txt, "End") Then Exit Do
                Loop
            ElseIf InStr(1, txt, "Begin Report") = 1 Then
                isReport = True
                OutFile.WriteLine txt
            ElseIf isReport = True And (InStr(1, txt, "    Right =") Or InStr(1, txt, "    Bottom =")) Then
                'skip line
                If InStr(1, txt, "    Bottom =") Then
                    isReport = False
                End If
            Else
                OutFile.WriteLine txt
            End If
        Loop
        OutFile.Close
        InFile.Close

        FSO.DeleteFile (Path & fileName)

        Dim thisFile As Object
        Set thisFile = FSO.GetFile(Path & obj_name & ".sanitize")
        
        ' Error Handling to deal with errors caused by Dropbox, VirusScan,
        ' or anything else touching the file.
        Dim ErrCounter As Integer
        On Error GoTo ErrorHandler
        thisFile.Move (Path & fileName)
        fileName = Dir$()
        SysCmd acSysCmdUpdateMeter, counter + 1
    Loop
    
    Exit Sub
ErrorHandler:
    ErrCounter = ErrCounter + 1
    If ErrCounter = 20 Then  ' 20 attempts seems like a nice arbitrary number
        MsgBox "This file could not be moved: " & vbNewLine, vbCritical + vbApplicationModal, _
            "Error moving file..."
        Resume Next
    End If
    Select Case Err.Number
        Case 58    ' "File already exists" error.
            DoEvents
            Sleep 10
            Resume    ' Go back to what you were doing
        Case Else
            MsgBox "This file could not be moved: " & vbNewLine, vbCritical + vbApplicationModal, _
                "Error moving file..."
    End Select
    Resume Next
End Sub
