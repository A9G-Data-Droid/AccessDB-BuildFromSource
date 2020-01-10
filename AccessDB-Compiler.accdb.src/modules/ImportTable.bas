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
    tempFilePath = VCS_File.VCS_TempFile()
    
    VCS_ConvertUtf8Ucs2 obj_path & tblName & ".LNKD", tempFilePath
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

' Import Table Definition
Public Sub VCS_ImportTableDef(ByVal tblName As String, ByVal directory As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim filePath As String
    Dim tbl As Object
    Dim prefix As String
    
    Dim thisDB As Database
    Set thisDB = appInstance.CurrentDb
    
    ' Drop table first.
    On Error Resume Next
    Set tbl = thisDB.TableDefs(tblName)
    On Error GoTo 0

    If Not tbl Is Nothing Then
        thisDB.Execute "Drop Table [" & tblName & "]"
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
    
    Dim Db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim InFile As Object
    Dim c As Long
    Dim buf As String
    Dim Values() As String
    Dim Value As Variant
    
    Dim tempFileName As String
    tempFileName = VCS_File.VCS_TempFile()
    VCS_File.VCS_ConvertUtf8Ucs2 obj_path & tblName & ".txt", tempFileName
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFileName, ForReading, False, TristateTrue)
    Set Db = appInstance.CurrentDb
    On Error GoTo ErrorHandler
    Db.Execute "DELETE FROM [" & tblName & "]"
    Set rs = Db.OpenRecordset(tblName)
    buf = InFile.ReadLine()
    Do Until InFile.AtEndOfStream
        buf = InFile.ReadLine()
        If Len(Trim$(buf)) > 0 Then
            Values = Split(buf, vbTab)
            c = 0
            rs.AddNew
            For Each fieldObj In rs.Fields
                DoEvents
                Value = Values(c)
                If Len(Value) = 0 Then
                    Value = Null
                Else
                    Value = Replace(Value, "\t", vbTab)
                    Value = Replace(Value, "\n", vbCrLf)
                    Value = Replace(Value, "\\", "\")
                End If
                rs(fieldObj.Name) = Value
                c = c + 1
            Next
            rs.Update
        End If
    Loop
    
ErrorHandler:
    rs.Close
    InFile.Close
    FSO.DeleteFile tempFileName
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