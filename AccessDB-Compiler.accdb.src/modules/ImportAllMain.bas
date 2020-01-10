Option Compare Database
Option Explicit

' Turn on to get additional debug messages in your output.
Public Const DebugOutput = False

'export/import all Queries as plain SQL text
Const HandleQueriesAsSQL = False

Public Function VCS_SourcePath() As String
    VCS_SourcePath = VCS_Dir.VCS_ProjectPath() & "source\"
End Function

Public Function ImportObjType(ByVal FileName As String, ByVal obj_type_label As String, ByVal obj_type_num As Integer, Optional ByVal ignoreVCS As Boolean = False, Optional ByVal src_path As String, Optional ByRef appInstance As Application) As Integer
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim obj_path As String
    Dim obj_name As String
    Dim ucs2 As Boolean

    If src_path = vbNullString Then src_path = VCS_SourcePath

    ImportObjType = 0
    obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)
    obj_path = src_path & obj_type_label & "\"
    If obj_type_label = "modules" Then
        ucs2 = False
    Else
        ucs2 = VCS_File.VCS_UsingUcs2(appInstance)
    End If

    VCS_IE_Functions.VCS_ImportObject obj_type_num, obj_name, obj_path & FileName, ucs2, appInstance
    ImportObjType = 1
End Function

Public Sub ImportObjTypeSource(ByVal obj_type As Variant, Optional ByVal ignoreVCS As Boolean = False, Optional ByVal src_path As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_num As Integer
    Dim obj_path As String
    Dim obj_name As String
    Dim obj_count As Integer
    Dim FileName As String
   
    'InitVCS_UsingUcs2
    
    If src_path = vbNullString Then src_path = VCS_SourcePath
    
    obj_type_split = Split(obj_type, "|")
    obj_type_label = obj_type_split(0)
    obj_type_num = Val(obj_type_split(1))
    obj_path = src_path & obj_type_label & "\"

    If (obj_type_label = "modules") Then
        If Not ImportReference.VCS_ImportReferences(src_path, appInstance) Then
            Debug.Print "Info: no references file in " & src_path
            Debug.Print
        End If
    End If
    
    FileName = Dir$(obj_path & "*.bas")
    If Len(FileName) > 0 Then
        Debug.Print VCS_String.VCS_PadRight("Importing " & obj_type_label & "...", 24);
        SysCmd acSysCmdInitMeter, "Importing " & obj_type_label, 100
        obj_count = 0
        Do Until Len(FileName) = 0
            obj_count = obj_count + ImportObjType(FileName, obj_type_label, obj_type_num, ignoreVCS, src_path, appInstance)
            FileName = Dir$()
            SysCmd acSysCmdUpdateMeter, obj_count
        Loop
        
        SysCmd acSysCmdRemoveMeter
        Debug.Print "[" & obj_count & "]"
    End If
End Sub

Public Sub ImportTableDef(ByVal FileName As String, Optional ByVal src_path As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim obj_name As String
    Dim obj_path As String

    If src_path = vbNullString Then src_path = VCS_SourcePath
    
    obj_path = src_path & "tbldefs\"
    obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)
    If DebugOutput Then
        Debug.Print "  [debug] table " & obj_name;
        Debug.Print
    End If
    
    ImportTable.VCS_ImportTableDef CStr(obj_name), obj_path, appInstance
End Sub

Public Sub ImportAllTableDefs(Optional ByVal src_path As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim obj_path As String
    Dim FileName As String
    Dim obj_count As Integer
    Dim obj_name As String

    If src_path = vbNullString Then src_path = VCS_SourcePath
    
    obj_path = src_path & "tbldefs\"
    FileName = Dir$(obj_path & "*.xml")
    If Len(FileName) > 0 Then
        Debug.Print VCS_String.VCS_PadRight("Importing tabledefs...", 24);
        SysCmd acSysCmdInitMeter, "Importing tabledefs", 100
        obj_count = 0
        If DebugOutput Then Debug.Print
        Do Until Len(FileName) = 0
            ImportTableDef FileName, src_path, appInstance
            obj_count = obj_count + 1
            FileName = Dir$()
            SysCmd acSysCmdUpdateMeter, obj_count
        Loop
        
        SysCmd acSysCmdRemoveMeter
        Debug.Print "[" & obj_count & "]"
    End If

    ' restore linked tables - we must have access to the remote store to import these!
    Dim searchPath As Object
    Set searchPath = FSO.GetFolder(obj_path)
    
    Debug.Print VCS_String.VCS_PadRight("Importing Linked Tables", 24);
    SysCmd acSysCmdInitMeter, "Importing Linked tabledefs", searchPath.Files.Count
    obj_count = 0
    If DebugOutput Then Debug.Print
    Dim foundFile As Object
    For Each foundFile In searchPath.Files
        If Right$(foundFile.Name, 5) = ".LNKD" Then
            obj_name = Mid$(foundFile.Name, 1, InStrRev(foundFile.Name, ".") - 1)
            If DebugOutput Then
                Debug.Print "  [debug] table " & obj_name;
                Debug.Print
            End If
            
            ImportTable.VCS_ImportLinkedTable CStr(obj_name), obj_path, appInstance
            obj_count = obj_count + 1
    
            SysCmd acSysCmdUpdateMeter, obj_count
        End If
    Next foundFile
    
    SysCmd acSysCmdRemoveMeter
    Debug.Print "[" & obj_count & "]"
End Sub

Public Sub ImportAllTableData(Optional ByVal src_path As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim obj_path As String
    Dim FileName As String
    Dim obj_count As Integer
    Dim obj_name As String

    If src_path = vbNullString Then src_path = VCS_SourcePath
    
    obj_path = src_path & "tables\"
    FileName = Dir$(obj_path & "*.txt")
    If Len(FileName) > 0 Then
        Debug.Print VCS_String.VCS_PadRight("Importing tables...", 24);
        SysCmd acSysCmdInitMeter, "Importing tables", 100
        obj_count = 0
        Do Until Len(FileName) = 0
            DoEvents
            obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
            ImportTableData CStr(obj_name), obj_path, appInstance
            obj_count = obj_count + 1
            FileName = Dir$()
            SysCmd acSysCmdUpdateMeter, obj_count
        Loop
        
        SysCmd acSysCmdRemoveMeter
        Debug.Print "[" & obj_count & "]"
    End If
End Sub

Public Sub ImportAllTableDataMacros(Optional ByVal src_path As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim obj_path As String
    Dim FileName As String
    Dim obj_count As Integer
    Dim obj_name As String

    If src_path = vbNullString Then src_path = VCS_SourcePath
    obj_path = src_path & "tbldef\"
    FileName = Dir$(obj_path & "*.dm")
    If Len(FileName) > 0 Then
        Debug.Print VCS_String.VCS_PadRight("Importing Data Macros...", 24);
        SysCmd acSysCmdInitMeter, "Importing Data Macros", 100
        obj_count = 0
        Do Until Len(FileName) = 0
            DoEvents
            obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)
            'VCS_Table.VCS_ImportTableData CStr(obj_name), obj_path
            ImportDataMacro.VCS_ImportDataMacros obj_name, obj_path, appInstance
            obj_count = obj_count + 1
            FileName = Dir$()
            SysCmd acSysCmdUpdateMeter, obj_count
        Loop
        
        SysCmd acSysCmdRemoveMeter
        Debug.Print "[" & obj_count & "]"
    End If
End Sub

' Main entry point for IMPORT. Import all forms, reports, queries,
' macros, modules, and lookup tables from `source` folder under the
' database's folder.
Public Sub ImportAllSource(Optional ByVal ignoreVCS As Boolean = False, Optional source_path As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
        
    Dim obj_path As String
    Dim obj_type As Variant
    Dim obj_count As Integer
    Dim FileName As String
    Dim obj_name As String
    Dim appendOnly As Boolean

    'InitVCS_UsingUcs2

    If source_path = vbNullString Then source_path = VCS_SourcePath
    If Not FSO.FolderExists(source_path) Then
        MsgBox "No source found at:" & vbCrLf & source_path, vbExclamation, "Import failed"
        Exit Sub
    End If

    obj_path = source_path & "queries\"
    FileName = Dir$(obj_path & "*.bas")
    
    If Len(FileName) > 0 Then
        Debug.Print VCS_String.VCS_PadRight("Importing queries...", 24);
        SysCmd acSysCmdInitMeter, "Importing queries", 100
        obj_count = 0
        Do Until Len(FileName) = 0
            DoEvents
            obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)
            'Check for plain sql export/import
            If HandleQueriesAsSQL Then
                ImportQuery.ImportQueryFromSQL obj_name, obj_path & FileName, False, appInstance.CurrentDb
            Else
                Dim tempFilePath As String
                tempFilePath = VCS_File.VCS_TempFile()
                VCS_IE_Functions.VCS_ImportObject acQuery, obj_name, obj_path & FileName, VCS_File.VCS_UsingUcs2, appInstance
                VCS_IE_Functions.VCS_ExportObject acQuery, obj_name, tempFilePath, VCS_File.VCS_UsingUcs2, appInstance
                VCS_IE_Functions.VCS_ImportObject acQuery, obj_name, tempFilePath, VCS_File.VCS_UsingUcs2, appInstance
                VCS_Dir.VCS_DelIfExist tempFilePath
            End If
            
            obj_count = obj_count + 1
            FileName = Dir$()
            SysCmd acSysCmdUpdateMeter, obj_count
        Loop
        
        SysCmd acSysCmdRemoveMeter
        Debug.Print "[" & obj_count & "]"
    End If
    
    ' restore table definitions
    ImportAllTableDefs source_path, appInstance
    
    ' NOW we may load data
    ImportAllTableData source_path, appInstance
    
    'load Data Macros - not DRY!
    ImportAllTableDataMacros source_path, appInstance

    'import access objects
    For Each obj_type In Split( _
        "forms|" & acForm & "," & _
        "reports|" & acReport & "," & _
        "macros|" & acMacro & "," & _
        "modules|" & acModule _
        , "," _
         )
        ImportObjTypeSource obj_type, ignoreVCS, source_path, appInstance
    Next
    
    ' Load Settings found in Options > Current Database
    ImportProperties source_path, appInstance
    
    'import Print Variables
    Debug.Print VCS_String.VCS_PadRight("Importing Print Vars...", 24);
    obj_count = 0
    
    obj_path = source_path & "reports\"
    FileName = Dir$(obj_path & "*.pv")
    Do Until Len(FileName) = 0
        DoEvents
        obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)
        ImportReport.VCS_ImportPrintVars obj_name, obj_path & FileName, appInstance
        obj_count = obj_count + 1
        FileName = Dir$()
    Loop
    
    Debug.Print "[" & obj_count & "]"
    
    'import relations
    Debug.Print VCS_String.VCS_PadRight("Importing Relations...", 24);
    obj_count = 0
    obj_path = source_path & "relations\"
    FileName = Dir$(obj_path & "*.txt")
    Do Until Len(FileName) = 0
        DoEvents
        ImportRelations.VCS_ImportRelation obj_path & FileName, appInstance
        obj_count = obj_count + 1
        FileName = Dir$()
    Loop
    
    Debug.Print "[" & obj_count & "]"
    DoEvents
    Debug.Print "Done."
End Sub