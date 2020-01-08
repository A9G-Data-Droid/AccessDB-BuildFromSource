Option Compare Database
Option Explicit

' -----------------------------------------------------------------
'
' 2019-01-21 Adam Kauffman
'
'   All the Application.FileDialog windows, customized for Access.
'
' -----------------------------------------------------------------
Public Function GetOpenFile(Optional ByVal varDirectory As String = vbNullString, _
                            Optional ByVal varTitleForDialog As String = "Please choose an Access Database...", _
                            Optional ByVal strFilterName As String = "All Files (*.*)", _
                            Optional ByVal strFilterString As String = "*.*") As String
    
    Const msoFileDialogOpen = 1
    Dim filePicker As Object
    Set filePicker = Application.FileDialog(msoFileDialogOpen)
    With filePicker
        .AllowMultiSelect = False
        .Title = varTitleForDialog
        .InitialFileName = varDirectory
        .Filters.Clear
        .Filters.Add strFilterName, strFilterString
        .Filters.Add "ACCDE Files (*.accde)", "*.ACCDE"
        .Filters.Add "MDE Files (*.mde)", "*.MDE"
        .Filters.Add "Access (*.mdb)", "*.MDB"
        .Filters.Add "Access 2007 (*.accdb)", "*.ACCDB"
    End With
    
    If filePicker.Show Then GetOpenFile = filePicker.SelectedItems(1)
End Function

Public Function GetFile( _
       Optional ByVal pblnOpen As Boolean = True, _
       Optional ByVal varDirectory As String = vbNullString, _
       Optional ByVal varTitleForDialog As String = "Please choose a file name ...", _
       Optional ByVal pstrFilter As String = vbNullString) As String
         
    Const msoFileDialogFilePicker = 3
    Dim filePicker As Object
    Set filePicker = Application.FileDialog(msoFileDialogFilePicker)
    With filePicker
        .AllowMultiSelect = False
        .Title = varTitleForDialog
        .InitialFileName = varDirectory
        .Filters.Clear
        .Filters.Add "All Files (*.*)", "*.*"
    End With
    
    If filePicker.Show Then GetFile = filePicker.SelectedItems(1)
End Function

Public Function BrowseFolder(Optional ByVal szDialogTitle As String = "Select a folder.", Optional ByVal startPath As String = "%SYSTEMROOT%") As String
    Const msoFileDialogFolderPicker = 4
    Dim folderPicker As Object
    Set folderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With folderPicker
        .ButtonName = "Select"
        .InitialFileName = startPath
        .Title = szDialogTitle
    End With
    
    If folderPicker.Show Then
        BrowseFolder = folderPicker.SelectedItems(1)
        If Right$(BrowseFolder, 1) <> "\" Then BrowseFolder = BrowseFolder & "\"
    End If
End Function

Public Function GetSaveasFile(Optional ByVal varDirectory As String = vbNullString, _
                              Optional ByVal varTitleForDialog As String = "Please choose an Access Database...", _
                              Optional ByVal strFilterName As String = "All Files (*.*)", _
                              Optional ByVal strFilterString As String = "*.*") As String
    
    Const msoFileDialogSaveAs = 2
    Dim filePicker As Object
    Set filePicker = Application.FileDialog(msoFileDialogSaveAs)
    With filePicker
        .AllowMultiSelect = False
        .Title = varTitleForDialog
        .InitialFileName = varDirectory
    End With
    
    If filePicker.Show Then GetSaveasFile = filePicker.SelectedItems(1)
End Function