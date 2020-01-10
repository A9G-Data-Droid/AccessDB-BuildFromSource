Option Explicit
Option Compare Database
Option Private Module
'---------------------------------------------------------------------------------------
' Module    : ImportProperties
' Author    : Adam Kauffman
' Date      : 2020-01-09
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
    
    Debug.Print VCS_String.VCS_PadRight("Importing Properties...", 24);
    
    ' Save list of properties set in current database.
    Dim thisDB As Object
    If CurrentProject.ProjectType = acMDB Then
        Set thisDB = appInstance.CurrentDb
    Else                                         ' ADP project
        Set thisDB = appInstance.CurrentProject
    End If
    
    Dim inputFile As Object
    Set inputFile = FSO.OpenTextFile(sourcePath & propertiesFile, ForReading)
    
    Dim propertyCount As Long
    Dim fileLine As String
    On Error GoTo ErrorHandler
    Do Until inputFile.AtEndOfStream
        fileLine = inputFile.ReadLine
        Dim Item() As String
        Item = Split(fileLine, "=")
        If UBound(Item) > 0 Then ' Looks like a valid entry
            propertyCount = propertyCount + 1
            
            Dim propertyName As String
            Dim propertyValue As Variant
            Dim propertyType As Long
            propertyName = Item(0)
            propertyValue = Item(1)
            If UBound(Item) > 1 Then
                propertyType = Item(2)
            Else
                propertyType = -1
            End If
            
            SetProperty propertyName, propertyValue, thisDB, propertyType
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
            Debug.Print fileLine & " Error: " & Err.Number & " " & Err.Description
        End If
        
        Err.Clear
        Resume Next
    End If
    
    On Error GoTo 0
    
    Debug.Print "[" & propertyCount & "]"
    inputFile.Close
    Set inputFile = Nothing
    ImportProperties = True

End Function

'SetProperty() requires that either intPType is set explicitly OR that
'              varPVal has a valid value if a new property is to be created.
Public Sub SetProperty(ByVal propertyName As String, ByVal propertyValue As Variant, _
                       Optional ByRef thisDB As Database, _
                       Optional ByVal propertyType As Integer = -1)
                       
    If thisDB Is Nothing Then Set thisDB = CurrentDb
    
    Dim newProperty As Property
    Set newProperty = GetProperty(propertyName, thisDB)
    If Not newProperty Is Nothing Then
        If newProperty.Value <> propertyValue Then newProperty.Value = propertyValue
    Else ' Property not found
        If propertyType = -1 Then propertyType = DBVal(varType(propertyValue))
        Set newProperty = thisDB.CreateProperty(propertyName, propertyType, propertyValue)
        thisDB.Properties.Append newProperty
    End If
End Sub

' Returns nothing upon Error: 3270 Property not found.
Public Function GetProperty(ByVal propertyName As String, _
                            Optional ByRef thisDB As Database) As Property
                            
    Const PropertyNotFound As Integer = 3270
    If thisDB Is Nothing Then Set thisDB = CurrentDb
    
    On Error GoTo Err_PropertyExists
    Set GetProperty = thisDB.Properties(propertyName)

    Exit Function
     
Err_PropertyExists:
    If Err.Number <> PropertyNotFound Then
        Debug.Print "Error getting property: " & propertyName & vbNewLine & Err.Number & " " & Err.Description
    End If
    
    Err.Clear
End Function

'   HERE BE DRAGONS
' Return db property type that closely matches VBA varible type
Private Function DBVal(ByVal intVBVal As Integer) As Integer
    Const TypeVBToDB As String = "\2|3\3|4\4|6\5|7\6|5" & _
                                 "\7|8\8|10\11|1\14|20\17|2"
    Dim intX As Integer
    intX = InStr(1, TypeVBToDB, "\" & intVBVal & "|")
    DBVal = Val(Mid$(TypeVBToDB, intX + Len(intVBVal) + 2))
End Function