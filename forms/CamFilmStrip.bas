Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3960
    DatasheetFontHeight =11
    ItemSuffix =24
    Left =345
    Top =1470
    Right =4215
    Bottom =11460
    Filter ="[ImgID]=0"
    RecSrcDt = Begin
        0x6f418eaed3c6e440
    End
    RecordSource ="FilmStripRecSource"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    FilterOnLoad =255
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            SizeMode =3
            PictureAlignment =2
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
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
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
        Begin TextBox
            FELineBreak = NotDefault
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
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
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
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
        Begin FormHeader
            Height =0
            BackColor =16252927
            Name ="FormHeader"
        End
        Begin Section
            Height =2700
            BackColor =6108695
            Name ="Detail"
            Begin
                Begin Image
                    PictureType =1
                    Left =300
                    Top =120
                    Width =3360
                    Height =2460
                    Name ="Image0"
                    ControlSource ="ImgPath"

                    LayoutCachedLeft =300
                    LayoutCachedTop =120
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =2580
                    TabIndex =1
                End
                Begin CommandButton
                    Transparent = NotDefault
                    OverlapFlags =247
                    Left =300
                    Top =120
                    Width =3360
                    Height =2460
                    Name ="Command9"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =300
                    LayoutCachedTop =120
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =2580
                End
            End
        End
        Begin FormFooter
            Height =375
            BackColor =6108695
            Name ="FormFooter"
            AutoHeight =1
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =3720
                    Height =315
                    FontWeight =700
                    ForeColor =16252927
                    Name ="txtPhotoCount"
                    ControlSource ="=Format(Nz(Count(*),0),\"#,##0\") & \" Photo\" & IIf(Count(*)=1,\"\",\"s\")"
                    FontName ="Franklin Gothic Book"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =375
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

Public Sub Form_Current()

On Error GoTo ErrHandler

'Look for jpg file
ShowPhoto

'Display photo data in main form
Me.Parent.lblFileName.Caption = Me.FileName
Me.Parent.lblDate.Caption = "Date: " & Format(Me.ImageDate, "short date")
Me.Parent.lblTime.Caption = "Time: " & Format(Me.ImageDate, "medium time")
Me.Parent.lblEasting.Caption = "UTM E: " & Round(Me.UTM_E)
Me.Parent.lblNorthing.Caption = "UTM N: " & Round(Me.UTM_N)

'Run through detection table to generate captions
Me.Parent.lblSpecies.Caption = GetSpeciesCaption
Me.Parent.lblComments.Caption = GetCommentCaption

Exit Sub

ErrHandler:
    ErrorMsg "An error occurred in the PhotoViewer form.", _
        Err.Number, Err.Description

End Sub

'------------------------------
'Form-specific custom functions
'------------------------------

Private Function GetCommentCaption() As String
'Concatenate comments for non-deleted detections

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim strComments As String

On Error GoTo ErrHandler

strComments = ""

Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT Detections.Comments FROM " & _
    "Detections WHERE (((Detections.ImageID) = " & Me.ImgID & ") AND " & _
    "((Detections.StatusID)<3))")
If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        If Not IsNull(rs!Comments) Then
            strComments = strComments & rs!Comments & ", "
        End If
        rs.MoveNext
    Loop
End If

'Strip trailing comma and space
If Len(strComments) > 0 Then
    strComments = Left(strComments, Len(strComments) - 2)
    strComments = "Comments: " & strComments
End If

GCCExit:
    GetCommentCaption = strComments
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function

ErrHandler:
    strComments = "<Error>"
    Resume GCCExit

End Function

Private Function GetSpeciesCaption() As String
'Concatenate species/details for non-deleted detections

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim strSpecies As String
Dim strDetail As String
Dim strCaption As String

On Error GoTo ErrHandler

strCaption = ""

'Get detection records for current photo
Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT Species.CommonName, " & _
    "DetectionDetails.DetailText, Max(Detections.Individuals) AS Ind " & _
    "FROM (Detections INNER JOIN Species ON " & _
    "Detections.SpeciesID = Species.SpeciesID) LEFT JOIN DetectionDetails " & _
    "ON Detections.DetailID = DetectionDetails.DetailID " & _
    "WHERE (((Species.CommonName) Is Not Null) AND ((Detections.ImageID) = " & Me.ImgID & ") AND " & _
    "((Detections.StatusID)<3)) GROUP BY Species.CommonName, DetectionDetails.DetailText " & _
    "ORDER BY Species.CommonName")
If Not rs.EOF Then
    'Get first species
    rs.MoveFirst
    strSpecies = rs!CommonName
    strDetail = ""
    Do Until rs.EOF
        If rs!CommonName = strSpecies Then
            'Concatenate details for this species
            If Not IsNull(rs!DetailText) Then
                strDetail = strDetail & rs!DetailText & "/"
            End If
            rs.MoveNext
        Else
            'Add species string to caption
            strCaption = strCaption & strSpecies
            If Len(strDetail) > 1 Then
                strDetail = Left(strDetail, Len(strDetail) - 1)
                strCaption = strCaption & " - " & strDetail
            End If
            strCaption = strCaption & "; "
            strSpecies = rs!CommonName
            strDetail = ""
        End If
    Loop
    'Add last species
    strCaption = strCaption & strSpecies
    If Len(strDetail) > 1 Then
        strDetail = Left(strDetail, Len(strDetail) - 1)
        strCaption = strCaption & " - " & strDetail
    End If
End If

'Add label and strip last semicolon
If Len(strCaption) > 0 Then
    strCaption = "Species: " & strCaption
End If

GSCExit:
    GetSpeciesCaption = strCaption
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function

ErrHandler:
    strCaption = "<Error>"
    Resume GSCExit

End Function

Private Sub ShowPhoto()

With Me.Parent.lblNotFound
    If FileExists(Me.ImgPath) Then
        'File is there, hide "Not Found"
        .Visible = False
        .Height = 0
        .Width = 0
        .Top = 0
        .Left = 0
    Else
        'File is missing, show "Not Found"
        .Visible = True
        .Height = 0.3021 * 1440
        .Width = 2.5 * 1440
        .Top = 2.5 * 1440
        .Left = 5.5 * 1440
    End If
End With

End Sub
