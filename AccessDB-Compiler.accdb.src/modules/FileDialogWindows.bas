Option Compare Database
Option Explicit
'-----------------------------------------------------------------
'
' All the Application.FileDialog windows, customized for Access.
'   by Adam Kauffman on 2019-01-21
'
'-----------------------------------------------------------------

Public Function GetOpenFile(Optional ByVal startPath As String = vbNullString, _
                            Optional ByVal titleForDialog As String = "Please choose an Access Database...", _
                            Optional ByVal filterName As String = "All Files (*.*)", _
                            Optional ByVal filterString As String = "*.*") As String
    
    Const msoFileDialogOpen = 1
    Dim filePicker As Object
    Set filePicker = Application.FileDialog(msoFileDialogOpen)
    With filePicker
        .AllowMultiSelect = False
        .Title = titleForDialog
        .InitialFileName = startPath
        .Filters.Clear
        .Filters.Add filterName, filterString
        .Filters.Add "ACCDE Files (*.accde)", "*.ACCDE"
        .Filters.Add "MDE Files (*.mde)", "*.MDE"
        .Filters.Add "Access (*.mdb)", "*.MDB"
        .Filters.Add "Access 2007 (*.accdb)", "*.ACCDB"
    End With
    
    If filePicker.Show Then GetOpenFile = filePicker.SelectedItems(1)
End Function

Public Function GetFile(Optional ByVal startPath As String = vbNullString, _
                        Optional ByVal titleForDialog As String = "Please choose a file name ...", _
                        Optional ByVal filterName As String = "All Files (*.*)", _
                        Optional ByVal filterString As String = "*.*") As String
         
    Const msoFileDialogFilePicker = 3
    Dim filePicker As Object
    Set filePicker = Application.FileDialog(msoFileDialogFilePicker)
    With filePicker
        .AllowMultiSelect = False
        .Title = titleForDialog
        .InitialFileName = startPath
        .Filters.Clear
        .Filters.Add filterName, filterString
    End With
    
    If filePicker.Show Then GetFile = filePicker.SelectedItems(1)
End Function

Public Function BrowseFolder(Optional ByVal startPath As String = vbNullString, _
                             Optional ByVal titleForDialog As String = "Select a folder.") As String
                            
    Const msoFileDialogFolderPicker = 4
    Dim folderPicker As Object
    Set folderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With folderPicker
        .ButtonName = "Select"
        .InitialFileName = startPath
        .Title = titleForDialog
    End With
    
    If folderPicker.Show Then
        BrowseFolder = folderPicker.SelectedItems(1)
        If Right$(BrowseFolder, 1) <> "\" Then BrowseFolder = BrowseFolder & "\"
    End If
End Function

Public Function GetSaveasFile(Optional ByVal startPath As String = vbNullString, _
                              Optional ByVal titleForDialog As String = "Please choose an Access Database...") As String
    
    Const msoFileDialogSaveAs = 2
    Dim filePicker As Object
    Set filePicker = Application.FileDialog(msoFileDialogSaveAs)
    With filePicker
        .AllowMultiSelect = False
        .Title = titleForDialog
        .InitialFileName = startPath
    End With
    
    If filePicker.Show Then GetSaveasFile = filePicker.SelectedItems(1)
End Function