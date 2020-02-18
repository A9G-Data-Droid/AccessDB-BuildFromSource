Option Compare Database
Option Private Module
Option Explicit

Private LinkedDBPath As String

Public Sub VCS_ImportLinkedTable(ByVal tblName As String, ByRef obj_path As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim Db As Database
    Dim InFile As Object
    
    Set Db = appInstance.CurrentDb
    
    Dim tempFilePath As String
    tempFilePath = FileAccess.GetTempFile()
    
    ConvertUtf8Ucs2 obj_path & tblName & ".LNKD", tempFilePath
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFilePath, iomode:=ForReading, create:=False, Format:=TristateTrue)
    
    On Error Resume Next
    appInstance.DoCmd.DeleteObject acTable, tblName
    Err.Clear
    
    On Error GoTo Err_CreateLinkedTable:
    
    Dim td As TableDef
    Set td = Db.CreateTableDef(InFile.ReadLine())
    
    Dim connect As String
    connect = InFile.ReadLine()
    If InStr(1, connect, "DATABASE=.\") Then     'replace relative path with literal path
        If LinkedDBPath = vbNullString Then
            LinkedDBPath = FSO.GetParentFolderName(FSO.GetFolder(obj_path).ParentFolder)
        End If
        
        connect = Replace(connect, "DATABASE=.\", "DATABASE=" & LinkedDBPath & "\")
        If Not ConnectFileValid(connect) Then  ' Let the user select the backend
            LinkedDBPath = FileDialogWindows.BrowseFolder("Select the location of the backend that contains your linked tables:", LinkedDBPath)
            connect = Replace(connect, "DATABASE=.\", "DATABASE=" & LinkedDBPath & "\")
        End If
    End If
    
    td.Attributes = dbAttachSavePWD
    td.connect = connect
    
    td.SourceTableName = InFile.ReadLine()
    Db.TableDefs.Append td
    
    GoTo Err_CreateLinkedTable_Fin
    
Err_CreateLinkedTable:
    MsgBox Err.Description, vbCritical, "ERROR: IMPORT LINKED TABLE"
    Resume Err_CreateLinkedTable_Fin
    
Err_CreateLinkedTable_Fin:
    'this will throw errors if a primary key already exists or the table is linked to an access database table
    'will also error out if no pk is present
    On Error GoTo Err_LinkPK_Fin:
    
    Dim Fields As String
    Fields = InFile.ReadLine()
    Dim Field As Variant
    Dim sql As String
    sql = "CREATE INDEX __uniqueindex ON " & td.Name & " ("
    
    For Each Field In Split(Fields, ";+")
        sql = sql & "[" & Field & "]" & ","
    Next
    
    'remove extraneous comma
    sql = Left$(sql, Len(sql) - 1)
    
    sql = sql & ") WITH PRIMARY"
    Db.Execute sql
    
Err_LinkPK_Fin:
    On Error GoTo ErrorHandler
    InFile.Close

ErrorHandler:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Sub



Public Function TableExists(ByVal tableName As String) As Boolean
    TableExists = Not IsNull(DLookup("Name", "MSysObjects", "Name='" & tableName & "'"))
End Function



' Import Table Definition
Public Sub VCS_ImportTableDef(ByVal tblName As String, ByVal directory As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim filePath As String
    Dim tbl As Object
    Dim prefix As String
    
    Dim thisDb As Database
    Set thisDb = appInstance.CurrentDb
    
    ' Drop table first.
    If TableExists(tblName) Then
        thisDb.Execute "Drop Table [" & tblName & "]"
    End If
    
    filePath = directory & tblName & ".xml"
    appInstance.ImportXML DataSource:=filePath, ImportOptions:=acStructureOnly
    
    prefix = Left$(tblName, 2)
    If prefix = "t_" Or prefix = "u_" Then
        appInstance.SetHiddenAttribute acTable, tblName, True
    End If
End Sub

' Import the lookup table `tblName` from `source\tables`.
Public Sub ImportTableData(tblName As String, obj_path As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim thisDb As Object        ' DAO.Database
    Dim tableRecords As Object  ' DAO.Recordset
    Dim thisField As Object     ' DAO.Field
    Dim tempFile As Object      ' FSO.File
    Dim currentRecord As Long
    Dim lineBuffer As String
    Dim fieldRecords() As String
    Dim thisRecord As Variant
    
    Dim tempFileName As String
    tempFileName = obj_path & tblName & ".txt" 'VCS_File.GetTempFile()

    ' Open file for reading with                              Create=False, Unicode=True (USC-2 Little Endian format)
    Set tempFile = FSO.OpenTextFile(tempFileName, ForReading, False, TristateFalse)
        
    Set thisDb = appInstance.CurrentDb
    thisDb.Execute "DELETE FROM [" & tblName & "]"
    Set tableRecords = thisDb.OpenRecordset(tblName)
    
    On Error GoTo ErrorHandler
    
    lineBuffer = tempFile.ReadLine  'Text(adReadLine)  ' Discard Header Line
    Do Until tempFile.AtEndOfStream   'EOS
        lineBuffer = tempFile.ReadLine  'Text(adReadLine)
        If Len(Trim$(lineBuffer)) > 0 Then
            fieldRecords = Split(lineBuffer, vbTab)
            currentRecord = 0
            tableRecords.AddNew
            For Each thisField In tableRecords.Fields
                DoEvents
                thisRecord = fieldRecords(currentRecord)
                Dim destinationField As Object
                Set destinationField = tableRecords.Fields(thisField.Name)
                If Len(thisRecord) = 0 Then
                    If destinationField.Required Then
                        Select Case destinationField.Type
                        Case 10 Or 12 Or 18  ' String type
                            thisRecord = vbNullString
                        Case Else  ' Numeric type
                            thisRecord = 0
                        End Select
                    Else  ' Not required can be null
                        thisRecord = Null
                    End If
                ElseIf IsNumeric(thisRecord) Then
                ' Don't process numbers
                Else ' Convert symbols back to their original state, as exported.
                    thisRecord = Replace(thisRecord, "\t", vbTab)
                    thisRecord = Replace(thisRecord, "\n", vbCrLf)
                    thisRecord = Replace(thisRecord, "\\", "\")
                End If
                
                ' Explicit data type conversion avoids Error: 3421 Data type conversion error.
                Select Case destinationField.Type
                Case dbBoolean
                    destinationField.Value = CBool(thisRecord)
                Case dbDate Or dbTimeStamp
                    destinationField.Value = CDate(thisRecord)
                Case Else
                    destinationField.Value = thisRecord
                End Select
                                
                currentRecord = currentRecord + 1
            Next
            
            tableRecords.Update
        End If
    Loop
    
ErrorHandler:
    If Err.Number > 0 Then 'TODO: handle specific error
        Form_LogWindow.WriteError vbCrLf & "Failed to Import Table Data: """ & tblName & """" & "<br/>" & vbCrLf & _
            "Field: " & currentRecord & "<br/>" & vbCrLf & _
            "Field Value: " & thisRecord & "<br/>" & vbCrLf & _
            "Line Buffer: " & lineBuffer & "<br/>" & vbCrLf & _
            "Error: " & Err.Number & " " & Err.Description
        
        Resume Next
    End If
    
    tableRecords.Close
    tempFile.Close
End Sub

' Check if the file in the connection string is valid
'   Return true if the file is found.
Private Function ConnectFileValid(ByVal connection As String) As Boolean
    Dim fileStart As Long
    fileStart = InStr(1, connection, "DATABASE=", vbTextCompare) + 9
    Dim FileName As String
    FileName = Mid$(connection, fileStart, Len(connection) - fileStart + 1)
    ConnectFileValid = FSO.FileExists(FileName)
End Function