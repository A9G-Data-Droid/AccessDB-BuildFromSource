Option Compare Database
Option Private Module
Option Explicit

' Import References from a GUID or file, true=SUCCESS
Public Function VCS_ImportReferences(ByVal obj_path As String, Optional ByRef appInstance As Application) As Boolean
    If appInstance Is Nothing Then Set appInstance = Application.Application
    Dim InFile As Object
    Dim line As String
    Dim strParts() As String
    Dim GUID As String
    Dim Major As Long
    Dim Minor As Long
    Dim FileName As String
    Dim refName As String
    
    FileName = Dir$(obj_path & "references.csv")
    If Len(FileName) = 0 Then
        VCS_ImportReferences = False
        Exit Function
    End If
    
    Set InFile = FSO.OpenTextFile(obj_path & FileName, iomode:=ForReading, create:=False, Format:=TristateFalse)
    
    Dim refCount As Long
    Debug.Print VCS_String.VCS_PadRight("Importing References...", 24);
    On Error GoTo failed_guid
    Do Until InFile.AtEndOfStream
        line = InFile.ReadLine
        strParts = Split(line, ",")
        If UBound(strParts) = 2 Then      'a ref with a guid @Timabell branch
            GUID = Trim$(strParts(0))
            Major = CLng(strParts(1))
            Minor = CLng(strParts(2))
            appInstance.References.AddFromGuid GUID, Major, Minor
        ElseIf UBound(strParts) = 3 Then  'a ref with a guid @JoyfullService branch
            GUID = Trim$(strParts(0))
            Major = CLng(strParts(2))
            Minor = CLng(strParts(3))
            appInstance.References.AddFromGuid GUID, Major, Minor
        Else
            refName = Trim$(strParts(0))
            appInstance.References.AddFromFile refName
        End If
        
        refCount = refCount + 1
go_on:
    Loop
    
    On Error GoTo 0
    VCS_ImportReferences = True
    Debug.Print "[" & refCount & "]"
    InFile.Close
    
failed_guid:
    If Err.Number = 32813 Then
        'The reference is already present in the access project - so we can ignore the error
        Resume Next
    ElseIf Err.Number > 0 Then
        MsgBox "Failed to register " & GUID, , "Error: " & Err.Number
        'Do we really want to carry on the import with missing references??? - Surely this is fatal
        Resume go_on
    End If
End Function