Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularCharSet =204
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10080
    DatasheetFontHeight =11
    ItemSuffix =20
    Left =3225
    Top =2490
    Right =28545
    Bottom =15015
    BeforeDelConfirm ="[Event Procedure]"
    Filter ="[IndividualID]>0"
    OrderBy ="[IndividualName]"
    RecSrcDt = Begin
        0x8d1b915ab47ce440
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Individuals"
    Caption ="Individuals"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =255
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    PictureSizeMode =4
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =18
            FontSize =10
            BorderColor =11644565
            ForeColor =7233610
            FontName ="Constantia"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Rectangle
            BackStyle =0
            BorderColor =11644565
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Line
            BorderColor =11644565
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Image
            BackStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =11644565
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CommandButton
            TextFontFamily =18
            FontWeight =400
            ForeColor =7233610
            FontName ="Constantia"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =16118511
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin OptionButton
            LabelX =230
            LabelY =-30
            BorderColor =11644565
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin OptionGroup
            BackStyle =1
            BorderColor =11644565
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin BoundObjectFrame
            SizeMode =3
            BackStyle =0
            LabelX =-1800
            BorderColor =11644565
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =204
            LabelX =-1800
            FontSize =11
            BorderColor =11644565
            FontName ="Calibri"
            AsianLineBreak =1
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin ListBox
            TextFontCharSet =204
            LabelX =-1800
            FontSize =11
            BorderColor =11644565
            FontName ="Calibri"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =204
            LabelX =-1800
            FontSize =11
            BorderColor =11644565
            FontName ="Calibri"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderColor =11644565
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =11644565
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CustomControl
            OldBorderStyle =1
            BorderColor =11644565
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin ToggleButton
            TextFontFamily =18
            FontWeight =400
            ForeColor =7233610
            FontName ="Constantia"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =16118511
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Tab
            TextFontCharSet =204
            FontSize =11
            FontName ="Calibri"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Attachment
            BackStyle =0
            PictureSizeMode =3
            Width =4800
            Height =3840
            BorderColor =11644565
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineColor =7233610
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            LabelX =-1800
        End
        Begin FormHeader
            Height =1027
            BackColor =6108695
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =34
                    Left =2760
                    Top =720
                    Width =2040
                    Height =285
                    BackColor =6108695
                    BorderColor =5029044
                    ForeColor =16252927
                    Name ="IndividualName_Label"
                    Caption ="Individual Name"
                    FontName ="Franklin Gothic Book"
                    Tag ="DetachedLabel"
                    ColumnGroup =1
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =2760
                    LayoutCachedTop =720
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =34
                    Left =360
                    Top =720
                    Width =2100
                    Height =285
                    BackColor =6108695
                    BorderColor =5029044
                    ForeColor =16252927
                    Name ="SpeciesID_Label"
                    Caption ="Species"
                    FontName ="Franklin Gothic Book"
                    Tag ="DetachedLabel"
                    ColumnGroup =2
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =360
                    LayoutCachedTop =720
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =34
                    Left =5100
                    Top =720
                    Width =1245
                    Height =285
                    BackColor =6108695
                    BorderColor =5029044
                    ForeColor =16252927
                    Name ="GenderID_Label"
                    Caption ="Gender"
                    FontName ="Franklin Gothic Book"
                    Tag ="DetachedLabel"
                    ColumnGroup =3
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =5100
                    LayoutCachedTop =720
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =34
                    Left =6420
                    Top =720
                    Width =2880
                    Height =285
                    BackColor =6108695
                    BorderColor =5029044
                    ForeColor =16252927
                    Name ="Comments_Label"
                    Caption ="Comments"
                    FontName ="Franklin Gothic Book"
                    Tag ="DetachedLabel"
                    ColumnGroup =4
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =6420
                    LayoutCachedTop =720
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =60
                    Top =60
                    Width =2805
                    Height =510
                    FontSize =20
                    BackColor =6108695
                    BorderColor =5029044
                    ForeColor =-2147483608
                    Name ="Label14"
                    Caption ="Individuals List"
                    FontName ="Trebuchet MS"
                    GridlineColor =0
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =2865
                    LayoutCachedHeight =570
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =9600
                    Top =120
                    Width =366
                    Height =366
                    Name ="cmdHelp"
                    Caption ="Help"
                    StatusBarText ="Help"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Help"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000e0e8e000e0c8b000 ,
                        0xe0d8d000e0d0c010e0d0c010d0d0c010d0d0c000d0d0d000e0e0e00000000000 ,
                        0x0000000000000000000000000000000000000000f0e8e0009068303080582080 ,
                        0x905010c0804820e0804820c0804810b06040108050381030d0c8c01000000000 ,
                        0x000000000000000000000000e0780000e0a05010a0683070c08860f0e0c8b0ff ,
                        0xf0f0f0fffffffffffffffffff0f0f0ffe0c8c0ffa07850c040301060d0c8c010 ,
                        0xe0d8d0000000000000000000e0882000b0703070e0a880fffff0e0ffe0b8a0ff ,
                        0xd08050ffc05820ffc05820ffd08050ffe0b8a0fff0e8e0ffb09070f050301060 ,
                        0xd0c8c000e0e0e00000000000b0783030d09870f0fff0e0ffe0a890ffc05010ff ,
                        0xc05010ffe0a890ffffffffffb04810ffb04810ffd0a080fff0f0e0ffa07050d0 ,
                        0x50381030d0d0d000f0f0f000b0784080f0d8c0fff0c8b0ffe05820ffd05810ff ,
                        0xd05010ffe08050ffe0a880ffc05010ffb04810ffb04810ffe0b8a0ffe0c8c0ff ,
                        0x50401080d0d0d010f0f0f000d08040e0fff8f0fff09870fff06020ffe05820ff ,
                        0xe05820fff0a890ffffffffffd05010ffc05010ffb05010ffc07850fff0f0f0ff ,
                        0x804020c0e0d0c000f0f0f000d08040f0ffffffffff7840ffff6830fff06820ff ,
                        0xf06020fff08850fffffffffff0c0b0ffc05820ffb05010ffb05820ffffffffff ,
                        0x804820e0e0d0c010f0f0f000d08850f0ffffffffff8050ffff7030ffff6830ff ,
                        0xff6830ffff6820fff09060fffff8f0fff0d8c0ffc05020ffc05820ffffffffff ,
                        0x804820e0e0d8d010f0f0f000d08050c0fff8f0ffffa880ffff7040ffff8850ff ,
                        0xffb090ffff7030fff06820fff09070fffffffffff08050ffd08860fffff0f0ff ,
                        0x805820b0e0d8d010f0f0f000c0804070f0d8c0ffffd0c0ffff7840ffff9870ff ,
                        0xffffffffffc8b0ffff9060ffffc8b0fffff8f0fff07840fff0c8b0ffe0c8b0ff ,
                        0x90602070e0c8b00000000000c0884030e0a070f0fff8f0ffffc0a0ffff7840ff ,
                        0xffb8a0fffff8f0fffffffffffff0e0ffff9870fff0b8a0fffff0e0ffc08850e0 ,
                        0xa0682030f0e8e0000000000000000000c0884060e0b8a0f0fff8f0ffffd0c0ff ,
                        0xffa880ffff8850ffff8850ffffa880fff0d0c0fffff0e0ffd0a880f0a0683060 ,
                        0xe0c0a00000000000000000000000000000000000c0884060e0a070f0f0d8c0ff ,
                        0xfff8f0fffffffffffffffffffff8f0fff0d8c0ffc09060e0a0703050f0b89000 ,
                        0x0000000000000000000000000000000000000000f0f0f000c0884030c0804070 ,
                        0xe0a070c0d09870e0d09860f0d09870d0b0784070b0784020f0e8f00000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0f0f000f0f0f000f0f0f000f0f0f000f0f0f00000000000f0f0f00000000000 ,
                        0x0000000000000000
                    End
                    BackStyle =0

                    LayoutCachedLeft =9600
                    LayoutCachedTop =120
                    LayoutCachedWidth =9966
                    LayoutCachedHeight =486
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =34
                    Left =9368
                    Top =720
                    Width =660
                    Height =285
                    BackColor =6108695
                    BorderColor =5029044
                    ForeColor =16252927
                    Name ="Label15"
                    Caption =" "
                    FontName ="Franklin Gothic Book"
                    ColumnGroup =5
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =9368
                    LayoutCachedTop =720
                    LayoutCachedWidth =10028
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =34
                    Left =2528
                    Top =720
                    Width =165
                    Height =285
                    BackColor =6108695
                    BorderColor =5029044
                    ForeColor =16252927
                    Name ="Label17"
                    FontName ="Franklin Gothic Book"
                    ColumnGroup =6
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =2528
                    LayoutCachedTop =720
                    LayoutCachedWidth =2693
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =34
                    Left =4868
                    Top =720
                    Width =165
                    Height =285
                    BackColor =6108695
                    BorderColor =5029044
                    ForeColor =16252927
                    Name ="Label19"
                    FontName ="Franklin Gothic Book"
                    ColumnGroup =7
                    GroupTable =1
                    GridlineColor =0
                    LayoutCachedLeft =4868
                    LayoutCachedTop =720
                    LayoutCachedWidth =5033
                    LayoutCachedHeight =1005
                End
            End
        End
        Begin Section
            Height =428
            BackColor =16252927
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2760
                    Top =60
                    Width =2040
                    Height =330
                    ColumnWidth =3000
                    TabIndex =2
                    BorderColor =7169897
                    Name ="IndividualName"
                    ControlSource ="IndividualName"
                    ColumnGroup =1
                    GroupTable =1

                    LayoutCachedLeft =2760
                    LayoutCachedTop =60
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6420
                    Top =60
                    Width =2880
                    Height =330
                    ColumnWidth =3000
                    TabIndex =5
                    BorderColor =7169897
                    Name ="Comments"
                    ControlSource ="Comments"
                    ColumnGroup =4
                    GroupTable =1

                    LayoutCachedLeft =6420
                    LayoutCachedTop =60
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =390
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =360
                    Top =60
                    Width =2100
                    Height =330
                    ColumnWidth =1530
                    TabIndex =1
                    BorderColor =7169897
                    Name ="SpeciesID"
                    ControlSource ="SpeciesID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Species.SpeciesID, Species.ShortName FROM Species ORDER BY Species.ShortN"
                        "ame;"
                    ColumnWidths ="0;1440"
                    ColumnGroup =2
                    GroupTable =1

                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =390
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =9368
                    Top =60
                    Width =660
                    Height =330
                    FontSize =11
                    TabIndex =6
                    ForeColor =0
                    Name ="cmdDelete"
                    Caption ="Command17"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ColumnGroup =5
                    GroupTable =1
                    GridlineColor =7233610
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000020202030000800900000004000080010 ,
                        0x0000000000000000000000000000000000000000000000000000000000000010 ,
                        0x0000000000000000000000001010100010181080101010ff202020c010081020 ,
                        0x1010100000000000000000000000000000000000101010004038309010101030 ,
                        0x3030200000000000000000000000000020202050201810ff302820ff30282080 ,
                        0x100810100000000000000000000000000008000030282020202820e010101020 ,
                        0x3030200000000000000000000000000040303010202020a0201820f0303030e0 ,
                        0x3030303000000000000000001010101010081010303030d02020209020202000 ,
                        0x000000000000000000000000000000000000000030283020202020d0202020ff ,
                        0x303030c020182010101010101010101020202090303030f03028202000000000 ,
                        0x00000000000000000000000000000000000000000000000030303020302820d0 ,
                        0x302820ff302830b020181030202020b0303030ff302820700000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000030383020 ,
                        0x302820d0303030ff303030ff403830ff302830e0000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x30303040303030ff403840ff303030e030202010000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000004040400040404030 ,
                        0x404040d0403830ff303030f0303030f0302820b0302820100000000000000000 ,
                        0x00000000000000000000000000000000000000004040400040404070404040ff ,
                        0x404040ff505050a04038403040384090404040ff303030903028201000000000 ,
                        0x0000000000000000000000004040402040404040404040b0404040ff404040ff ,
                        0x505050b0505050100000000040383010504840b0404040ff3038309030303010 ,
                        0x00000000000000004040400040404060404040b0403830ff404040ff50485080 ,
                        0x504850000000000000000000000000003028201040404040504850f0404040a0 ,
                        0x00000000000000004040400040484060404040ff404040f05048505050505000 ,
                        0x0000000000000000000000000000000000000000404040005050502040484070 ,
                        0x00000000000000000000000050485030504850c0504850405048500000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =9368
                    LayoutCachedTop =60
                    LayoutCachedWidth =10028
                    LayoutCachedHeight =390
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5100
                    Top =60
                    Width =1245
                    Height =330
                    ColumnWidth =1530
                    TabIndex =4
                    BorderColor =7169897
                    Name ="GenderID"
                    ControlSource ="GenderID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lkupGender.GenderID, lkupGender.GenderText FROM lkupGender ORDER BY lkupG"
                        "ender.GenderText;"
                    ColumnWidths ="0;1440"
                    DefaultValue ="3"
                    ColumnGroup =3
                    GroupTable =1

                    LayoutCachedLeft =5100
                    LayoutCachedTop =60
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2528
                    Top =60
                    Width =165
                    Height =330
                    ForeColor =2366701
                    Name ="txtSpeciesReqd"
                    ControlSource ="=IIf([SpeciesID] Is Null,\"*\",\"\")"
                    OnEnter ="[Event Procedure]"
                    ColumnGroup =6
                    GroupTable =1

                    LayoutCachedLeft =2528
                    LayoutCachedTop =60
                    LayoutCachedWidth =2693
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4868
                    Top =60
                    Width =165
                    Height =330
                    TabIndex =3
                    ForeColor =2366701
                    Name ="txtIndNameReqd"
                    ControlSource ="=IIf([IndividualName] Is Null,\"*\",\"\")"
                    OnEnter ="[Event Procedure]"
                    ColumnGroup =7
                    GroupTable =1

                    LayoutCachedLeft =4868
                    LayoutCachedTop =60
                    LayoutCachedWidth =5033
                    LayoutCachedHeight =390
                End
            End
        End
        Begin FormFooter
            Height =1020
            BackColor =6108695
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3660
                    Top =60
                    Width =2760
                    Height =480
                    Name ="cmdDone"
                    Caption ="Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3660
                    LayoutCachedTop =60
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3960
                    Top =600
                    Width =2160
                    TabIndex =1
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3960
                    LayoutCachedTop =600
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =960
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

Private Const strGenericError As String = _
    "The Individuals form encountered an error."

Private Sub cmdCancel_Click()
'Quit out of adding new individual

On Error GoTo ErrHandler

Me.Undo
DoCmd.Close
Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description

End Sub

Private Sub cmdDelete_Click()
'Prompt to make sure the user wants to delete,
'and enable merging two individuals

Dim lngDeleteID As Long
Dim lngNextID As Long

On Error GoTo ErrHandler

If Me.Dirty = True Then
    Me.Undo
    Exit Sub
End If

'Prompt to confirm
If MsgBox("Are you sure you want to delete this record?", _
    vbOKCancel) = vbOK Then
    'Suppress warnings
    DoCmd.SetWarnings False
    'Get current record's ID
    lngDeleteID = Me.IndividualID
    'Check for IDs
    If DCount("*", "IndividualDetections", "IndividualID=" & _
        lngDeleteID) > 0 Then
        'Prompt to merge
        If MsgBox("Do you want to switch IDs for this individual to a " & _
            "remaining individual in the list?", vbYesNo) = vbYes Then
            'Open merge form
            DoCmd.OpenForm "MergeInd", , , "IndividualID=" & _
                lngDeleteID, , acDialog
            'Suppress warnings again - MergeInd restored them
            DoCmd.SetWarnings False
            'Check for IDs again
            If DCount("*", "IndividualDetections", _
                "IndividualID=" & lngDeleteID) > 0 Then
                'Prompt again if there are still IDs for this individual
                If MsgBox("You canceled the merge operation." & vbNewLine & _
                    vbNewLine & "Do you still want to delete this " & _
                    "individual from the list?", vbOKCancel) = vbCancel Then
                    GoTo DeleteExit
                End If
            End If
        Else
            'Delete IDs
            DoCmd.RunSQL "DELETE * FROM IndividualDetections WHERE " & _
                "IndividualID=" & lngDeleteID
        End If
    End If
    'Find appropriate record to move to after delete
    Me.RecordsetClone.FindFirst "IndividualID=" & lngDeleteID
    Me.RecordsetClone.MoveNext
    If Me.RecordsetClone.EOF Then
        Me.RecordsetClone.MovePrevious
        Me.RecordsetClone.MovePrevious
    End If
    lngNextID = Me.RecordsetClone!IndividualID
    'Delete record
    DoCmd.RunSQL "DELETE * FROM Individuals WHERE IndividualID=" & _
        lngDeleteID
    'Scroll to appropriate record
    Me.Requery
    Me.RecordsetClone.FindFirst "IndividualID=" & lngNextID
    Me.Bookmark = Me.RecordsetClone.Bookmark
End If

DeleteExit:
    DoCmd.SetWarnings True
    Exit Sub

ErrHandler:
    ErrorMsg "There was an error deleting the record.", Err.Number, _
        Err.Description
    Resume DeleteExit

End Sub

Private Sub cmdDone_Click()
'Close the form

On Error GoTo ErrHandler

DoCmd.Close
Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description

End Sub

Private Sub cmdHelp_Click()
'Display help message

HelpMsg Me.Name, "Individuals"

End Sub

Private Sub Form_AfterUpdate()
'Reset cancel button

Me.cmdCancel.Enabled = False

End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
'Suppress system delete message

On Error GoTo ErrHandler

Cancel = False

Response = acDataErrContinue
Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'Avoid saving records with no species/name

If IsNull(Me.SpeciesID) Or IsNull(Me.IndividualName) Then
    Cancel = True
    MsgBox "Species and name required."
Else
    Cancel = False
End If

End Sub

Private Sub Form_Current()
'Disable cancel button unless coming from
'IndividualID form

On Error GoTo ErrHandler

If Not IsNull(Me.OpenArgs) Then
    If Not Me.IndividualName = Left(Me.OpenArgs, _
            InStr(1, Me.OpenArgs, ";") - 1) Then
        Me.cmdCancel.Enabled = False
    End If
End If
Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description

End Sub

Private Sub Form_Dirty(Cancel As Integer)
'Enable cancel button to undo edit/add

On Error GoTo ErrHandler

Me.cmdCancel.Enabled = True
Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description

End Sub

Private Sub Form_Open(Cancel As Integer)
'Populates new record based on OpenArgs property
'and sets window size

On Error GoTo ErrHandler

'Add species from IndividualID form and display message
If Len(Me.OpenArgs) > 0 Then
    DoCmd.GoToRecord , , acNewRec
    Me.IndividualName = Left(Me.OpenArgs, _
        InStr(1, Me.OpenArgs, ";") - 1)
    Me.SpeciesID = Mid(Me.OpenArgs, _
        InStr(1, Me.OpenArgs, ";") + 1)
    Me.cmdCancel.Enabled = True
    MsgBox "The individual you added appears at the bottom " & _
        "of the form." & vbNewLine & vbNewLine & _
        "If you don't want to add this individual, " & _
        "press cancel."
End If

'Set form size
Me.InsideHeight = Me.Section(acHeader).Height + Me.Section(acFooter).Height + _
    18 * Me.Detail.Height

Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description

End Sub

Private Sub txtIndNameReqd_Enter()
'Prevent clicking/tabbing into textbox

Me.GenderID.SetFocus

End Sub

Private Sub txtSpeciesReqd_Enter()
'Prevent clicking/tabbing into textbox

Me.IndividualName.SetFocus

End Sub
