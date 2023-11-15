Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8640
    DatasheetFontHeight =11
    ItemSuffix =1
    Right =25320
    Bottom =12525
    RecSrcDt = Begin
        0xf19063204f55e440
    End
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
        Begin Section
            Height =5760
            BackColor =16252927
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =360
                    Top =240
                    Width =7500
                    Height =2100
                    BackColor =16252927
                    ForeColor =7169897
                    Name ="Label0"
                    Caption ="VBA is currently disabled.  To use the Database, you need to enable it by clicki"
                        "ng the options button and selecting enable.\015\012\015\012To avoid this in the "
                        "future, create a Trusted Location through Access Options and save the database t"
                        "here."
                    FontName ="Franklin Gothic Book"
                    LayoutCachedLeft =360
                    LayoutCachedTop =240
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =2340
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

'This is an empty form that is only used to open
'the switchboard when the database is opened
'If VBA is disabled this code won't run, and this
'form will remain visible disaplying instructions
'on how to enable the code

Private Sub Form_Open(Cancel As Integer)

DoCmd.Close
DoCmd.OpenForm "Switchboard"

End Sub
