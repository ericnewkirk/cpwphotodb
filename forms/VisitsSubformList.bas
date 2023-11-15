Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularCharSet =186
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =10620
    DatasheetFontHeight =11
    ItemSuffix =54
    Left =4200
    Top =4515
    Right =15090
    Bottom =9705
    OrderBy ="[VisitDate], [VisitTypeID] Mod 3"
    RecSrcDt = Begin
        0x17490adeccbce440
    End
    RecordSource ="qryVisitsRecSource"
    Caption ="Visits"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
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
            TextFontCharSet =186
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
            TextFontCharSet =186
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
            TextFontCharSet =186
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
            TextFontCharSet =186
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
            Height =352
            BackColor =6108695
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextFontFamily =34
                    Left =120
                    Top =60
                    Width =1500
                    Height =270
                    BackColor =6108695
                    ForeColor =16252927
                    Name ="VisitTypeID_Label"
                    Caption ="Visit Type"
                    FontName ="Franklin Gothic Book"
                    ColumnGroup =1
                    GroupTable =14
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =330
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    TextFontFamily =34
                    Left =1695
                    Top =60
                    Width =1905
                    Height =270
                    BackColor =6108695
                    ForeColor =16252927
                    Name ="VisitDate_Label"
                    Caption ="Date"
                    FontName ="Franklin Gothic Book"
                    ColumnGroup =2
                    GroupTable =14
                    LayoutCachedLeft =1695
                    LayoutCachedTop =60
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =330
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextFontFamily =34
                    Left =3675
                    Top =60
                    Width =4245
                    Height =270
                    BackColor =6108695
                    ForeColor =16252927
                    Name ="Comments_Label"
                    Caption ="Comments"
                    FontName ="Franklin Gothic Book"
                    ColumnGroup =3
                    GroupTable =14
                    LayoutCachedLeft =3675
                    LayoutCachedTop =60
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =330
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    TextFontFamily =34
                    Left =7995
                    Top =60
                    Width =825
                    Height =270
                    BackColor =6108695
                    ForeColor =16252927
                    Name ="PhotoCount_Label"
                    Caption ="Photos"
                    FontName ="Franklin Gothic Book"
                    ColumnGroup =4
                    GroupTable =14
                    LayoutCachedLeft =7995
                    LayoutCachedTop =60
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =330
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =420
            BackColor =16252927
            Name ="Detail"
            Begin
                Begin TextBox
                    FontUnderline = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9720
                    Top =60
                    Width =780
                    Height =299
                    FontSize =10
                    TabIndex =6
                    ForeColor =12349952
                    Name ="lnkModule"
                    ControlSource ="=IIf([VisitTypeID]<3,\"Module\",\"\")"
                    FontName ="Trebuchet MS"
                    GridlineColor =16118511

                    LayoutCachedLeft =9720
                    LayoutCachedTop =60
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =359
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1695
                    Top =60
                    Width =1905
                    Height =314
                    TabIndex =1
                    BorderColor =7169897
                    Name ="VisitDate"
                    ControlSource ="VisitDate"
                    Format ="Short Date"
                    ConditionalFormat = Begin
                        0x01000000b4000000010000000100000000000000000000002900000001000000 ,
                        0xff000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00560069007300690074005400790070006500490044005d003c0033002000 ,
                        0x41006e00640020005b0053006500740056006900730069007400490044005d00 ,
                        0x20004900730020004e0075006c006c0000000000
                    End
                    ColumnGroup =2
                    GroupTable =14

                    LayoutCachedLeft =1695
                    LayoutCachedTop =60
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =374
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Left =3675
                    Top =60
                    Width =4245
                    Height =314
                    TabIndex =2
                    BorderColor =7169897
                    Name ="Comments"
                    ControlSource ="Comments"
                    ConditionalFormat = Begin
                        0x01000000b4000000010000000100000000000000000000002900000001000000 ,
                        0xff000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00560069007300690074005400790070006500490044005d003c0033002000 ,
                        0x41006e00640020005b0053006500740056006900730069007400490044005d00 ,
                        0x20004900730020004e0075006c006c0000000000
                    End
                    ColumnGroup =3
                    GroupTable =14

                    LayoutCachedLeft =3675
                    LayoutCachedTop =60
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =374
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7995
                    Top =60
                    Width =825
                    Height =314
                    TabIndex =3
                    BorderColor =7169897
                    Name ="PhotoCount"
                    ControlSource ="PhotoCount"
                    ConditionalFormat = Begin
                        0x01000000b4000000010000000100000000000000000000002900000001000000 ,
                        0xff000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00560069007300690074005400790070006500490044005d003c0033002000 ,
                        0x41006e00640020005b0053006500740056006900730069007400490044005d00 ,
                        0x20004900730020004e0075006c006c0000000000
                    End
                    ColumnGroup =4
                    GroupTable =14

                    LayoutCachedLeft =7995
                    LayoutCachedTop =60
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =374
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextFontCharSet =204
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =120
                    Top =60
                    Width =1500
                    Height =314
                    BorderColor =7169897
                    ConditionalFormat = Begin
                        0x01000000b4000000010000000100000000000000000000002900000001000000 ,
                        0xed1c2400ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00560069007300690074005400790070006500490044005d003c0033002000 ,
                        0x41006e00640020005b0053006500740056006900730069007400490044005d00 ,
                        0x20004900730020004e0075006c006c0000000000
                    End
                    Name ="VisitTypeID"
                    ControlSource ="VisitTypeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT lkupVisitTypes.ID, lkupVisitTypes.VisitType FROM lkupVisitTypes ORDER BY "
                        "lkupVisitTypes.VisitType;"
                    ColumnWidths ="0;1440"
                    ColumnGroup =1
                    GroupTable =14

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =374
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8940
                    Top =60
                    Width =600
                    Height =300
                    TabIndex =4
                    ForeColor =12349952
                    Name ="cmdEditVisit"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    BackStyle =0

                    LayoutCachedLeft =8940
                    LayoutCachedTop =60
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =247
                    TextFontFamily =34
                    Left =9720
                    Top =60
                    Width =780
                    Height =300
                    TabIndex =5
                    ForeColor =12349952
                    Name ="cmdModule"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ImageData = Begin
                        0x00000000
                    End
                    BackStyle =0

                    LayoutCachedLeft =9720
                    LayoutCachedTop =60
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            Height =420
            BackColor =6108695
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4560
                    Top =60
                    Width =1560
                    Height =270
                    ForeColor =16252927
                    Name ="cmdNewVisit"
                    Caption ="Add New Visit"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    BackStyle =0

                    LayoutCachedLeft =4560
                    LayoutCachedTop =60
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =330
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
    "The Visits form encountered an error."

Private Sub cmdEditVisit_Click()

On Error GoTo ErrHandler

DoCmd.OpenForm "SingleVisitPopup", , , "VisitID=" & Me.VisitID, , acDialog
UpdateVisitDependencies Me.Parent.LocationID
Me.Requery

Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description, Me.Name

End Sub

Private Sub cmdModule_Click()

Dim strSQL As String

On Error GoTo ErrHandler

If Me.VisitTypeID = 3 Then
    Exit Sub
End If

If Me.PhotoCount = 0 Then
    MsgBox "No photos have been imported for this visit."
    Exit Sub
End If

If IsLoaded("PhotoViewer") Then
    DoCmd.Close acForm, "PhotoViewer", acSaveNo
End If

If QueryIsOpen("PhotoViewerQuery") Then
    DoCmd.Close acQuery, "PhotoViewerQuery", acSaveNo
End If

strSQL = CurrentDb.QueryDefs("FilmStripRecSource").SQL
strSQL = Replace(strSQL, "ORDER BY", "WHERE (((Visits.VisitID) = " & _
    Me.VisitID & ")) ORDER BY")
SaveQuery "PhotoViewerQuery", strSQL

DoCmd.OpenForm "CopyPhotos", , , , , acDialog

Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description, Me.Name

End Sub

Private Sub cmdNewVisit_Click()

On Error GoTo ErrHandler

DoCmd.OpenForm "AddVisitPopup", , , "LocationID=" & _
    Me.Parent.LocationID, , acDialog
UpdateVisitDependencies Me.Parent.LocationID
Me.Requery

Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description, Me.Name

End Sub
