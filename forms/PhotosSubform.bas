Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =186
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8640
    DatasheetFontHeight =11
    ItemSuffix =24
    Right =25050
    Bottom =12525
    RecSrcDt = Begin
        0x339a6faf2e38e440
    End
    RecordSource ="SELECT Photos.*, [FilePath] & [FileName] AS ImgPath FROM Photos;"
    Caption ="CamPhotos"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    NavigationCaption ="Photo:"
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
            Height =0
            BackColor =6108695
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =3420
            BackColor =16252927
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin Image
                    Left =2400
                    Top =120
                    Width =3840
                    Height =3180
                    Name ="Image13"
                    ControlSource ="ImgPath"

                    LayoutCachedLeft =2400
                    LayoutCachedTop =120
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =3300
                    TabIndex =3
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =255
                    TextFontFamily =34
                    Left =3060
                    Top =1200
                    Width =2460
                    Height =960
                    Name ="cmdLoadPhotos"
                    Caption ="Load Photos"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Import photos for this visit"

                    LayoutCachedLeft =3060
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =2160
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =119
                    TextFontFamily =34
                    Left =3060
                    Top =1440
                    Width =0
                    Height =0
                    FontSize =16
                    BackColor =16252927
                    ForeColor =-2147483608
                    Name ="NotFound_Label"
                    Caption ="Image not found"
                    FontName ="Trebuchet MS"
                    LayoutCachedLeft =3060
                    LayoutCachedTop =1440
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =1440
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =204
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2520
                    Top =2400
                    Width =3600
                    Height =660
                    FontSize =12
                    TabIndex =1
                    BorderColor =7169897
                    Name ="txtLoading"

                    LayoutCachedLeft =2520
                    LayoutCachedTop =2400
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =3060
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4560
                    Top =120
                    Width =0
                    Height =0
                    TabIndex =2
                    Name ="cmd0"
                    Caption ="cmd0"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =120
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =120
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =6108695
            Name ="FormFooter"
            AutoHeight =1
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

Private Sub cmdLoadPhotos_Click()
'Open ImportOptions with current visit

On Error GoTo ErrHandler

DoCmd.OpenForm "ImportOptions", , , , , acDialog, Me.Parent.VisitID
Me.Requery
Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
'Prevent adding a new record

On Error GoTo ErrHandler

MsgBox "Photo records cannot be added manually through this form."
Me.Undo
Cancel = True
Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description

End Sub

Public Sub Form_Current()
'Show/hide load photos button and image not found label

On Error GoTo ErrHandler

Me.cmd0.SetFocus

With Me.cmdLoadPhotos
    If Me.NewRecord Then
        'Show button
        .Visible = True
        .Width = 1.7083 * 1440
        .Height = 0.6667 * 1440
        .Top = 0.8333 * 1440
        .Left = 2.125 * 1440
    Else
        'Hide button
        .Visible = False
        .Width = 0
        .Height = 0
        .Top = 0
        .Left = 0
    End If
End With

With Me.NotFound_Label
    If IsNull(Me.ImgPath) Or FileExists(Nz(Me.ImgPath, "x.xxx")) Then
        'Hide label
        .Visible = False
        .Height = 0
        .Width = 0
        .Top = 0
        .Left = 0
    Else
        'Show label
        .Visible = True
        .Height = 0.3021 * 1440
        .Width = 1.66667 * 1440
        .Top = 1440
        .Left = 2.125 * 1440
    End If
End With

Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description

End Sub
