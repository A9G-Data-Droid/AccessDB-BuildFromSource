Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15780
    DatasheetFontHeight =11
    ItemSuffix =5
    Right =25575
    Bottom =14625
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0xb31cfbeccd68e540
    End
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =10200
            BackColor =3684411
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =3
            BackShade =25.0
            Begin
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    BorderWidth =3
                    OverlapFlags =215
                    TextFontFamily =49
                    IMESentenceMode =3
                    Left =180
                    Top =480
                    Width =15420
                    Height =8940
                    FontSize =12
                    BackColor =0
                    BorderColor =10921638
                    ForeColor =16777215
                    Name ="LogView"
                    FontName ="Consolas"
                    GridlineColor =10921638
                    TextFormat =1
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =180
                    LayoutCachedTop =480
                    LayoutCachedWidth =15600
                    LayoutCachedHeight =9420
                    BackThemeColorIndex =0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            SpecialEffect =3
                            BackStyle =1
                            OldBorderStyle =1
                            BorderWidth =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =180
                            Top =180
                            Width =15420
                            Height =315
                            BackColor =5855577
                            ForeColor =16777215
                            Name ="Label1"
                            Caption ="LOG WINDOW"
                            FontName ="Segoe UI"
                            GridlineColor =10921638
                            HorizontalAnchor =2
                            LayoutCachedLeft =180
                            LayoutCachedTop =180
                            LayoutCachedWidth =15600
                            LayoutCachedHeight =495
                            ThemeFontIndex =-1
                            BackThemeColorIndex =0
                            BackTint =65.0
                            BorderTint =100.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14160
                    Top =9600
                    TabIndex =1
                    ForeColor =16777215
                    Name ="SaveLog"
                    Caption ="Save Log File"
                    OnClick ="[Event Procedure]"
                    LeftPadding =105
                    TopPadding =60
                    RightPadding =120
                    BottomPadding =165
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    VerticalAnchor =1

                    LayoutCachedLeft =14160
                    LayoutCachedTop =9600
                    LayoutCachedWidth =15600
                    LayoutCachedHeight =9960
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    Gradient =2
                    BackColor =10855845
                    BackThemeColorIndex =6
                    BackTint =100.0
                    BorderColor =10855845
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    HoverColor =12040119
                    HoverThemeColorIndex =6
                    HoverTint =80.0
                    PressedColor =8684676
                    PressedThemeColorIndex =6
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =39
                    QuickStyleMask =-1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =180
                    Top =9600
                    Width =1800
                    FontSize =8
                    TabIndex =2
                    ForeColor =16777215
                    Name ="ShowAccess"
                    Caption ="Show Access GUI"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Requires 'Version Control.accda'"
                    LeftPadding =105
                    TopPadding =60
                    RightPadding =120
                    BottomPadding =165
                    VerticalAnchor =1

                    LayoutCachedLeft =180
                    LayoutCachedTop =9600
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =9960
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =2
                    BackColor =10855845
                    BackThemeColorIndex =6
                    BackTint =100.0
                    BorderColor =10855845
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =12040119
                    HoverThemeColorIndex =6
                    HoverTint =80.0
                    PressedColor =8684676
                    PressedThemeColorIndex =6
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =39
                    QuickStyleMask =-1
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =87
                    Left =2160
                    Top =9600
                    Width =1800
                    FontSize =8
                    TabIndex =3
                    ForeColor =16777215
                    Name ="LoadVCS"
                    Caption ="Load Version Control"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    ControlTipText ="Requires 'Version Control.accda'"
                    LeftPadding =105
                    TopPadding =60
                    RightPadding =120
                    BottomPadding =165
                    VerticalAnchor =1

                    LayoutCachedLeft =2160
                    LayoutCachedTop =9600
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =9960
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    Gradient =2
                    BackColor =10855845
                    BackThemeColorIndex =6
                    BackTint =100.0
                    BorderColor =10855845
                    BorderThemeColorIndex =6
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =12040119
                    HoverThemeColorIndex =6
                    HoverTint =80.0
                    PressedColor =8684676
                    PressedThemeColorIndex =6
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =1
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =39
                    QuickStyleMask =-1
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

Private VCSLoaded As Boolean
Private logWindowText As String

Private Sub SaveLog_Click()
    SaveLogFile
End Sub

Private Sub Form_Load()
    Application.SetOption "Behavior Entering Field", 2 ' 1 = Start 2= End 0 = all
    DoCmd.OpenForm "CompilerGUI", acNormal, , , acFormEdit, acWindowNormal
    LoadVCS.Visible = False
End Sub

Public Sub Append(ByVal newText As String)
    logWindowText = logWindowText & newText
    
    With Me.LogView
        .SetFocus
        .Value = logWindowText
        .Requery
        .SelStart = Len(.Text)
        .SelLength = 0
    End With
    
    Debug.Print Replace(newText, "<br/>", vbNullString);
End Sub

Public Sub WriteLine(Optional ByVal newText As String = vbNullString)
    Append newText & "<br/>" & vbNewLine
End Sub

Public Sub WriteError(ByVal newText As String)
    WriteLine "<font color=red>" & newText & "</font>"
End Sub

Public Sub ClearLog()
    logWindowText = vbNullString
    Me.LogView.Value = logWindowText
End Sub

Public Sub SaveLogFile()
    Dim saveAsPath As String
    saveAsPath = GetSaveasFile(VCS_Dir.VCS_ProjectPath() & "CompilerLog.htm", "Select name for log file:")
    If Not saveAsPath = vbNullString Then WriteFile Me.LogView.Value, saveAsPath, True
End Sub

Private Sub ShowAccess_Click()
    ShowAccessGui
    LoadVCS.Visible = True
End Sub

' Load or Unload Version Control Add-in
Private Sub LoadVCS_Click()
    VCSLoaded = Not VCSLoaded
    InitializeVersionControlSystem VCSLoaded
    VBE.ActiveCodePane.Show
End Sub
