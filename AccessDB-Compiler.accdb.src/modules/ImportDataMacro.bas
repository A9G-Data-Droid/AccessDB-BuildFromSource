Option Compare Database
Option Private Module
Option Explicit

' For Access 2007 (VBA6) and earlier
#If Not VBA7 Then
    Private Const acTableDataMacro As Integer = 12
#End If

Public Sub VCS_ImportDataMacros(ByVal tableName As String, ByVal directory As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim filePath As String
    filePath = directory & tableName & ".dm"
    
    VCS_IE_Functions.VCS_ImportObject acTableDataMacro, tableName, filePath, VCS_File.VCS_UsingUcs2, appInstance
End Sub