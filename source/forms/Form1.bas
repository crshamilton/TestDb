Version =20
VersionRequired =20
PublishOption =1
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =14
    Right =10650
    Bottom =7830
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0xd55613367a44e440
    End
    RecordSource ="Query1"
    Caption ="Query1"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
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
        Begin FormHeader
            Height =1080
            BackColor =15064278
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =360
                    Top =720
                    Width =840
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Table1.ID_Label"
                    Caption ="Table1.ID"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Table1_ID_Label"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =720
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =1260
                    Top =720
                    Width =4200
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Field1_Label"
                    Caption ="Field1"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =1260
                    LayoutCachedTop =720
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =5520
                    Top =720
                    Width =900
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Field2_Label"
                    Caption ="Field2"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5520
                    LayoutCachedTop =720
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =6480
                    Top =720
                    Width =4200
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Name1_Label"
                    Caption ="Name1"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =6480
                    LayoutCachedTop =720
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =10740
                    Top =720
                    Width =720
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Name2_Label"
                    Caption ="Name2"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10740
                    LayoutCachedTop =720
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =215
                    Left =60
                    Top =60
                    Width =1548
                    Height =1020
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label10"
                    Caption ="Query1"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1608
                    LayoutCachedHeight =1080
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =3780
                    Top =300
                    Width =576
                    Height =576
                    ForeColor =4210752
                    Name ="Command13"
                    Caption ="Command13"
                    ControlTipText ="Find Next"
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="FindNext"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command13\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"FindNextRecord\"/></"
                                "Statements></UserInterfaceMacro"
                        End
                        Begin
                            Comment ="_AXL:>"
                        End
                    End
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a39b17d4af3b17d4aff ,
                        0xb17d4a8a00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a39b17d4af3 ,
                        0xb17d4affb17d4a7e000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff727272ffffffffff727272ff727272ff727272ff727272ff ,
                        0x000000000000000000000000000000000000000000000000b17d4a3cb58353ff ,
                        0xb17d4affb17d4a7b727272ffffffffff727272ff727272ff727272ff727272ff ,
                        0x000000000000000000000000727272ff00000000b17d4a3cb17d4af3b17d4aff ,
                        0xb17d4a8a00000000727272ffffffffff727272ff727272ff727272ff727272ff ,
                        0x000000000000000000000000727272ff727272ff0000000000000000ffffffff ,
                        0x0000000000000000727272ffffffffff727272ff727272ff727272ff727272ff ,
                        0x000000000000000000000000727272ff727272ff727272ff727272ffffffffff ,
                        0x727272ff00000000727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff727272ff ,
                        0x727272ff0000000000000000727272ffffffffff727272ff727272ff727272ff ,
                        0x727272ffffffffff727272ff727272ff727272ff727272ffffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ffffffffff727272fc ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff727272ff727272ff727272ff727272ff727272ff727272ff00000000 ,
                        0x00000000000000000000000000000000727272ffffffffff727272ff727272ff ,
                        0x727272ff00000000727272ff727272ff727272ffffffffff727272ff00000000 ,
                        0x00000000000000000000000000000000727272ff727272ff727272ff727272ff ,
                        0x727272ff00000000727272ff727272ff727272ff727272ff727272ff00000000 ,
                        0x0000000000000000000000000000000000000000727272ffffffffff727272ff ,
                        0x000000000000000000000000727272ffffffffff727272ff0000000000000000 ,
                        0x0000000000000000000000000000000000000000727272ff727272ff727272ff ,
                        0x000000000000000000000000727272ff727272ff727272fc0000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =3780
                    LayoutCachedTop =300
                    LayoutCachedWidth =4356
                    LayoutCachedHeight =876
                    BackColor =15123357
                    BorderColor =15123357
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =2856
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =360
                    Top =60
                    Width =840
                    Height =315
                    ColumnWidth =1440
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Table1.ID"
                    ControlSource ="Table1.ID"
                    EventProcPrefix ="Table1_ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1260
                    Top =60
                    Width =4200
                    Height =600
                    ColumnWidth =3000
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Field1"
                    ControlSource ="Field1"
                    GridlineColor =10921638

                    LayoutCachedLeft =1260
                    LayoutCachedTop =60
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =660
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5520
                    Top =60
                    Width =900
                    Height =330
                    ColumnWidth =1530
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Field2"
                    ControlSource ="Field2"
                    GridlineColor =10921638

                    LayoutCachedLeft =5520
                    LayoutCachedTop =60
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6480
                    Top =60
                    Width =4200
                    Height =600
                    ColumnWidth =3000
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Name1"
                    ControlSource ="Name1"
                    GridlineColor =10921638

                    LayoutCachedLeft =6480
                    LayoutCachedTop =60
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =660
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10740
                    Top =60
                    Width =720
                    Height =330
                    ColumnWidth =1530
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Name2"
                    ControlSource ="Name2"
                    GridlineColor =10921638

                    LayoutCachedLeft =10740
                    LayoutCachedTop =60
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =390
                End
            End
        End
        Begin FormFooter
            Height =1380
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Width =576
                    Height =576
                    ForeColor =4210752
                    Name ="Command11"
                    Caption ="Command11"
                    ControlTipText ="Close Form"
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Close"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command11\" xmlns=\"http://schemas.microsoft.com/office/acce"
                                "ssservices/2009/11/application\"><Statements><Action Name=\"CloseWindow\"/></Sta"
                                "tements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000082c2ea0982c2ea4b82c2ea90 ,
                        0x82c2eade00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000082c2ea2182c2ea7582c2eab782c2eaf982c2eaff82c2eaff ,
                        0x82c2eaff00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff000000000000000000000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffffffffffffffffffff82c2eaff000000000000000000000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffffffffffffffffffa500000000000000000000000000000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffffffffffc000000000b17d4a90b17d4affb17d4af0b17d4a36 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffedffffff30b17d4a87b17d4affb17d4af0b17d4a3600000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaffffffffffd7ecf8ff82c2eaff ,
                        0x82c2eaffffffff30b17d4a81b17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff82c2eaff82c2eaff82c2eaffdceef9ffc4e2f5ff82c2eaff ,
                        0x82c2eaffffffff27b17d4a7eb17d4affb17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4aff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffe4ffffff27b17d4a84b17d4affb17d4af0b17d4a3900000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffffffffffbd00000000b17d4a8db17d4affb17d4af0b17d4a39 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffffffffffffffffffa500000000000000000000000000000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaffffffffffffffffffffffffff82c2eaff000000000000000000000000 ,
                        0x000000000000000082c2eaff82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff ,
                        0x82c2eaff82c2eaff82c2eaff82c2eaff82c2eaff000000000000000000000000 ,
                        0x000000000000000082c2ea2182c2ea6f82c2eab782c2eaf982c2eaff82c2eaff ,
                        0x82c2eaff00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000082c2ea0982c2ea4e82c2ea96 ,
                        0x82c2eae400000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedWidth =576
                    LayoutCachedHeight =576
                    BackColor =15123357
                    BorderColor =15123357
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =300
                    TabIndex =1
                    ForeColor =4210752
                    Name ="Command12"
                    Caption ="Command12"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1440
                    LayoutCachedTop =300
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =660
                    BackColor =15123357
                    BorderColor =15123357
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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

Private Sub Command12_Click()
MsgBox ("you clicked this button - change #3")
End Sub
