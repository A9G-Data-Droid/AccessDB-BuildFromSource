Option Explicit
Option Compare Database
Option Private Module

Private Const UnitSeparator = "?"  ' Chr(31) INFORMATION SEPARATOR ONE

Public Function ThisProjectDB(Optional ByRef appInstance As Application) As Object
    If appInstance Is Nothing Then Set appInstance = Application.Application
    If CurrentProject.ProjectType = acMDB Then
        Set ThisProjectDB = appInstance.CurrentDb
    Else  ' ADP project
        Set ThisProjectDB = appInstance.CurrentProject
    End If
End Function


'---------------------------------------------------------------------------------------
' Module    : ImportProperties
' Author    : Adam Kauffman
' Date      : 2020-01-10
' Purpose   : Import database properties from the exported source
'---------------------------------------------------------------------------------------

' Import database properties from a text file, true=SUCCESS
Public Function ImportProperties(ByVal sourcePath As String, Optional ByRef appInstance As Application) As Boolean
    If appInstance Is Nothing Then Set appInstance = Application.Application
      
    Dim propertiesFile As String
    propertiesFile = Dir(sourcePath & "properties.txt")
    If Len(propertiesFile) = 0 Then ' File not foud
        ImportProperties = False
        Exit Function
    End If
    
    Form_LogWindow.Append VCS_String.VCS_PadRight("Importing Properties...", 24)
    
    Dim thisDb As Object
    Set thisDb = ThisProjectDB(appInstance)
   
    Dim inputFile As Object
    Set inputFile = FSO.OpenTextFile(sourcePath & propertiesFile, ForReading)
    
    Dim propertyCount As Long
    On Error GoTo ErrorHandler
    Do Until inputFile.AtEndOfStream
        Dim recordUnit() As String
        recordUnit = Split(inputFile.ReadLine, UnitSeparator)
        If UBound(recordUnit) > 1 Then ' Looks like a valid entry
            propertyCount = propertyCount + 1
            
            Dim propertyName As String
            Dim propertyValue As Variant
            Dim propertyType As Long
            propertyName = recordUnit(0)
            propertyValue = recordUnit(1)
            propertyType = recordUnit(2)
            
            SetProperty propertyName, propertyValue, thisDb, propertyType
        End If
    Loop
    
ErrorHandler:
    If Err.Number > 0 Then
        If Err.Number = 3001 Then
            ' Invalid argument; means that this property cannot be set by code.
        ElseIf Err.Number = 3032 Then
            ' Cannot perform this operation; means that this property cannot be set by code.
        ElseIf Err.Number = 3259 Then
            ' Invalid field data type; means that the property was not found, use create.
        ElseIf Err.Number = 3251 Then
            ' Operation is not supported for this type of object; means that this property cannot be set by code.
        Else
            Form_LogWindow.WriteError " Error: " & Err.Number & " " & Err.Description
        End If
        
        Err.Clear
        Resume Next
    End If
    
    On Error GoTo 0
    
    Form_LogWindow.WriteLine "[" & propertyCount & "]"
    inputFile.Close
    Set inputFile = Nothing
    ImportProperties = True

End Function

' SetProperty() requires either propertyType is set explicitly OR
'   propertyValue has a valid value and type for a new property to be created.
Public Sub SetProperty(ByVal propertyName As String, ByVal propertyValue As Variant, _
                       Optional ByRef thisDb As Object, _
                       Optional ByVal propertyType As Integer = -1)
                       
    If thisDb Is Nothing Then Set thisDb = ThisProjectDB
    
    Dim newProperty As Property
    Set newProperty = GetProperty(propertyName, thisDb)
    If Not newProperty Is Nothing Then
        If newProperty.Value <> propertyValue Then newProperty.Value = propertyValue
    Else ' Property not found
        If propertyType = -1 Then propertyType = DAOType(varType(propertyValue)) ' Guess the type
        Set newProperty = thisDb.CreateProperty(propertyName, propertyType, propertyValue)
        thisDb.Properties.Append newProperty
    End If
End Sub

' Returns nothing upon Error
Public Function GetProperty(ByVal propertyName As String, _
                            Optional ByRef thisDb As Object) As Property
                            
    Const PropertyNotFound As Integer = 3270
    If thisDb Is Nothing Then Set thisDb = ThisProjectDB
    
    On Error GoTo Err_PropertyExists
    Set GetProperty = thisDb.Properties(propertyName)

    Exit Function
     
Err_PropertyExists:
    If Err.Number <> PropertyNotFound Then
        Form_LogWindow.WriteError "Error getting property: " & propertyName & vbNewLine & Err.Number & " " & Err.Description
    End If
    
    Err.Clear
End Function

' Return DataTypeEnum enumeration (DAO) property type that closely matches VBA varible type passed in
'   https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/datatypeenum-enumeration-dao
Private Function DAOType(ByVal VBAType As Integer) As Integer
    ' Handle arrays
    If VBAType > 8192 Then VBAType = VBAType - 8192

    Select Case VBAType
        Case vbInteger: DAOType = dbInteger
        Case vbLong: DAOType = dbLong
        Case vbSingle: DAOType = dbSingle
        Case vbDouble: DAOType = dbDouble
        Case vbCurrency: DAOType = dbCurrency
        Case vbDate: DAOType = dbDate
        Case vbString: DAOType = dbText
        Case vbObject: DAOType = dbAttachment
        Case vbError: DAOType = dbText
        Case vbBoolean: DAOType = dbBoolean
        Case vbVariant: DAOType = dbText
        Case vbDataObject: DAOType = dbLongBinary
        Case vbDecimal: DAOType = dbDecimal
        Case vbByte: DAOType = dbByte
        Case vbLongLong: DAOType = dbBigInt
        Case Else: DAOType = dbText
    End Select
End Function