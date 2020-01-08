Version =21
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =49
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11700
    ItemSuffix =43
    Left =-27256
    Top =4230
    Right =-15796
    Bottom =7095
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xfd4ee20d2d67e540
    End
    RecordSource ="CompilerSettings"
    Caption ="AccessDB Compiler"
    DatasheetFontName ="Courier"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =2880
            BackColor =3684411
            Name ="Detail"
            BackThemeColorIndex =3
            BackShade =25.0
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =240
                    Top =540
                    Width =10740
                    Name ="SourceFolderPath"
                    ControlSource ="SourcePath"
                    DefaultValue ="=VCS_SourcePath()"
                    FontName ="Segoe UI"
                    ControlTipText ="Path that contains source files exported from an AccessDB"

                    LayoutCachedLeft =240
                    LayoutCachedTop =540
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =780
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =180
                            Width =4020
                            Height =300
                            FontSize =11
                            ForeColor =16777215
                            Name ="Label0"
                            Caption ="Path to source files"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =240
                            LayoutCachedTop =180
                            LayoutCachedWidth =4260
                            LayoutCachedHeight =480
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11040
                    Top =540
                    Width =300
                    Height =300
                    TabIndex =1
                    Name ="SelectSourceFolder"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =11040
                    LayoutCachedTop =540
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =840
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =240
                    Top =1260
                    Width =10740
                    TabIndex =3
                    Name ="OutputFilePath"
                    ControlSource ="OutputPath"
                    DefaultValue ="=VCS_ProjectPath() & \"*.accdb\""
                    FontName ="Segoe UI"
                    ControlTipText ="Full path and filename of the binary you want to create."

                    LayoutCachedLeft =240
                    LayoutCachedTop =1260
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =1500
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =900
                            Width =4260
                            Height =300
                            FontSize =11
                            ForeColor =16777215
                            Name ="Label10"
                            Caption ="Destination full file path"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =240
                            LayoutCachedTop =900
                            LayoutCachedWidth =4500
                            LayoutCachedHeight =1200
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =79
                    Left =10200
                    Top =2160
                    Width =1140
                    TabIndex =4
                    Name ="cmdOK"
                    Caption ="&OK"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =10200
                    LayoutCachedTop =2160
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =2520
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    OverlapFlags =85
                    Left =120
                    Top =2040
                    Width =11220
                    BorderColor =-2147483628
                    Name ="Line143"
                    LayoutCachedLeft =120
                    LayoutCachedTop =2040
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =2040
                End
                Begin CheckBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =240
                    Top =1710
                    TabIndex =2
                    Name ="OverwriteFlag"
                    ControlSource ="OverwriteDB"
                    DefaultValue ="0"

                    LayoutCachedLeft =240
                    LayoutCachedTop =1710
                    LayoutCachedWidth =500
                    LayoutCachedHeight =1950
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =465
                            Top =1620
                            Width =6315
                            Height =300
                            FontSize =11
                            ForeColor =16777215
                            Name ="Label21"
                            Caption ="Overwrite if destination DB exists"
                            FontName ="Segoe UI"
                            LayoutCachedLeft =465
                            LayoutCachedTop =1620
                            LayoutCachedWidth =6780
                            LayoutCachedHeight =1920
                            ForeThemeColorIndex =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =11040
                    Top =1260
                    Width =300
                    Height =300
                    TabIndex =5
                    Name ="SelectOutputFile"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"

                    LayoutCachedLeft =11040
                    LayoutCachedTop =1260
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =1560
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub SelectOutputFile_Click()
    On Error GoTo Err_Section
    
    Dim selection As String
    If IsNull(Me.OutputFilePath.Value) Or Me.OutputFilePath.Value = vbNullString Then
        selection = GetSaveasFile(VCS_ProjectPath & "*.accdb")
    Else
        selection = GetSaveasFile(Me.OutputFilePath.Value & "*.accdb")
    End If
       
    If selection <> vbNullString Then Me.OutputFilePath.Value = selection
       
Err_Section:
    If Err.Number <> 0 Then MsgBox "Error in cmdBuild_Click - " & Err.Number & " - " & Err.Description

End Sub

Private Sub SelectSourceFolder_Click()
    Dim strRtn As String

    On Error Resume Next

    strRtn = BrowseFolder("Select a folder containing the source files.", Me.SourceFolderPath.Value)
    If strRtn <> vbNullString Then Me.SourceFolderPath.Value = strRtn
        
    On Error GoTo 0
    If Err.Number <> 0 Then MsgBox "Error in cmdBuild2_Click (9): " & Err.Number & " - " & Err.Description

End Sub

Private Sub cmdOK_Click()
    On Error GoTo Err_cmdOK_Click

    If Me.Dirty Then Me.Dirty = False
    
    Make.Build Me.SourceFolderPath.Value, Me.OutputFilePath.Value, Me.OverwriteFlag.Value

Err_cmdOK_Click:
    If Err.Number <> 0 Then MsgBox Err.Description
    
End Sub
