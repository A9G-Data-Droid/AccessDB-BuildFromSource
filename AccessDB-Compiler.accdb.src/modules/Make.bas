Option Compare Database
Option Explicit
'--------------------------------------------------------------------
' Complier to Make and Build Access DB from source files.
'   by Adam Kauffman on 2020-01-07
'
'   Designed to work with code exported by msaccess-vcs-integration
'       https://github.com/joyfullservice/msaccess-vcs-integration
'       https://github.com/timabell/msaccess-vcs-integration
'--------------------------------------------------------------------
Const Version As String = "0.6.0"



' This is the main entry point called by the Compiler GUI
Public Sub Build(ByVal sourceFolder As String, ByVal outputFile As String)
    Const cstrSpacer As String = "-------------------------------"
    Dim startTime As Single
    startTime = Timer
    With Form_LogWindow
        .Visible = True
        .ClearLog
        .WriteLine "<pre>"
        .WriteLine cstrSpacer
        .WriteLine " -= " & CurrentProject.Name & " =-"
        .WriteLine " Version: " & Version
        .WriteLine " Started: " & Now
        .WriteLine
        .WriteLine cstrSpacer
        .WriteLine " Source Folder:"
        .WriteLine sourceFolder
    End With
       
    Dim newApp As Application
    Set newApp = New Access.Application
    newApp.NewCurrentDatabase outputFile
    On Error GoTo ErrorHandler
    With Form_LogWindow
        .WriteLine " Created new DB: "
        .WriteLine newApp.CurrentDb.Name
        .WriteLine cstrSpacer
        .WriteLine
    End With
    
    If DebugOutput Then newApp.Visible = True

    ' This is the main procedure
    ImportAllSource False, sourceFolder, newApp
    
    With Form_LogWindow
        .WriteLine
        .WriteLine " Runtime: " & Round(Timer - startTime, 2) & " seconds"
        .WriteLine "</pre>"
    End With
    
ErrorHandler:
    If Err.Number > 0 Then
        Form_LogWindow.WriteError "Error: " & Err.Number & " " & Err.Description
    End If

    newApp.Quit acQuitSaveAll
End Sub



Public Sub HideAccessGui()
    With DoCmd
        .ShowToolbar "Ribbon", acToolbarNo
        .NavigateTo "acNavigationCategoryObjectType"
        .RunCommand acCmdWindowHide
    End With
End Sub



Public Sub ShowAccessGui()
    With DoCmd
        .ShowToolbar "Ribbon", acToolbarYes
        .SelectObject acForm, "CompilerGUI", True
    End With
End Sub