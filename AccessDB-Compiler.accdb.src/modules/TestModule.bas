Option Compare Database
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

' Test shows that UCS-2 files exported by Access make round trip through our conversions.
'@TestMethod("TextConversions")
Public Sub TestUCS2toUTF8RoundTrip()
    On Error GoTo TestFail
    
    'Arrange:
    Dim queryName As String
    queryName = "Temp_Test_Query_Delete_Me"
    Dim tempFileName As String
    tempFileName = GetTempFile()
    
    CurrentDb.CreateQueryDef queryName, "SELECT * FROM TEST WHERE TESTING=TRUE"
    Application.SaveAsText acQuery, queryName, tempFileName
    CurrentDb.QueryDefs.Delete queryName
    
    'Act:
    ConvertUtf8Ucs2 tempFileName, tempFileName & "UCS2UCS"
    ConvertUcs2Utf8 tempFileName & "UCS2UCS", tempFileName & "UTF8"
    ConvertUcs2Utf8 tempFileName & "UTF8", tempFileName & "UTF82UTF8"
    ConvertUtf8Ucs2 tempFileName & "UTF82UTF8", tempFileName & "UTF82UCS"
    
    ' Read original export
    Dim originalExport As String
    With FSO.OpenTextFile(tempFileName, , , TristateTrue)
        originalExport = .ReadAll
        .Close
    End With
    
    ' Read final file that went through all permutations of conversion
    Dim finalFile As String
    With FSO.OpenTextFile(tempFileName & "UTF82UCS", , , TristateTrue)
        finalFile = .ReadAll
        .Close
    End With
    
    'Assert:
    Assert.AreEqual originalExport, finalFile
    
    GoTo TestExit
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

TestExit:
    
End Sub