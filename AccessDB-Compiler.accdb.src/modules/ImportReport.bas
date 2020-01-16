Option Compare Database
Option Private Module
Option Explicit

' --------------------------------
' Structures
' --------------------------------
Private Type str_DEVMODE
    RGB As String * 94
End Type

Private Type type_DEVMODE
    strDeviceName(31) As Byte                    'vba strings are encoded in unicode (16 bit) not ascii
    intSpecVersion As Integer
    intDriverVersion As Integer
    intSize As Integer
    intDriverExtra As Integer
    lngFields As Long
    intOrientation As Integer
    intPaperSize As Integer
    intPaperLength As Integer
    intPaperWidth As Integer
    intScale As Integer
    intCopies As Integer
    intDefaultSource As Integer
    intPrintQuality As Integer
    intColor As Integer
    intDuplex As Integer
    intResolution As Integer
    intTTOption As Integer
    intCollate As Integer
    strFormName(31) As Byte
    lngPad As Long
    lngBits As Long
    lngPW As Long
    lngPH As Long
    lngDFI As Long
    lngDFr As Long
End Type

Public Sub VCS_ImportPrintVars(ByVal obj_name As String, ByVal filePath As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
   
    Dim DevModeString As str_DEVMODE
    Dim DevModeExtra As String
  
    Dim DM As type_DEVMODE
    Dim rpt As Report

    'report must be opened in design view to save changes to the print vars
    appInstance.DoCmd.OpenReport obj_name, acViewDesign
  
    Set rpt = appInstance.Reports(obj_name)
  
    'read print vars into struct
    If Not IsNull(rpt.PrtDevMode) Then
        DevModeExtra = rpt.PrtDevMode
        DevModeString.RGB = DevModeExtra
        LSet DM = DevModeString
    Else
        Form_LogWindow.WriteError "Warning: PrtDevMode is null"
        GoTo theEnd
    End If
  
    Dim InFile As Object
    Set InFile = FSO.OpenTextFile(filePath, iomode:=ForReading, create:=False, Format:=TristateFalse)
  
    'print out print var values
    '    DM.intOrientation = InFile.ReadLine
    '    DM.intPaperSize = InFile.ReadLine
    '    DM.intPaperLength = InFile.ReadLine
    '    DM.intPaperWidth = InFile.ReadLine
    '    DM.intScale = InFile.ReadLine
    
    ' Loop through lines instead of assuming order
    Do While Not InFile.AtEndOfStream
        Dim varLine As Variant
        varLine = Split(InFile.ReadLine, "=")
        If UBound(varLine) = 1 Then
            Select Case varLine(0)
            Case "Orientation":     DM.intOrientation = varLine(1)
            Case "PaperSize":       DM.intPaperSize = varLine(1)
            Case "PaperLength":     DM.intPaperLength = varLine(1)
            Case "PaperWidth":      DM.intPaperWidth = varLine(1)
            Case "Scale":           DM.intScale = varLine(1)
            Case Else
                Form_LogWindow.WriteError "* Unknown print var: '" & varLine(0) & "'"
            End Select
        End If
    Loop
    
    InFile.Close
   
    'write print vars back into report
    LSet DevModeString = DM
    Mid(DevModeExtra, 1, 94) = DevModeString.RGB
    rpt.PrtDevMode = DevModeExtra
    
theEnd:
    Set rpt = Nothing
    appInstance.DoCmd.Close acReport, obj_name, acSaveYes
End Sub