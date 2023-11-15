Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularCharSet =186
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8310
    DatasheetFontHeight =11
    ItemSuffix =25
    Right =21465
    Bottom =12525
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0xe92f57586f7be440
    End
    RecordSource ="SELECT Detections.* FROM Detections WHERE (((Detections.StatusID)<3));"
    Caption ="Detections"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin FormHeader
            Height =405
            BackColor =6108695
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =34
                    Left =120
                    Top =60
                    Width =1440
                    Height =285
                    FontSize =9
                    BackColor =6108695
                    ForeColor =16252927
                    Name ="SpeciesID_Label"
                    Caption ="Species"
                    FontName ="Franklin Gothic Book"
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =345
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =34
                    Left =4140
                    Top =60
                    Width =3330
                    Height =285
                    FontSize =9
                    BackColor =6108695
                    ForeColor =16252927
                    Name ="Comments_Label"
                    Caption ="Comments"
                    FontName ="Franklin Gothic Book"
                    LayoutCachedLeft =4140
                    LayoutCachedTop =60
                    LayoutCachedWidth =7470
                    LayoutCachedHeight =345
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =34
                    Left =3120
                    Top =60
                    Width =960
                    Height =285
                    FontSize =9
                    BackColor =6108695
                    ForeColor =16252927
                    Name ="Individuals_Label"
                    Caption ="Individuals"
                    FontName ="Franklin Gothic Book"
                    LayoutCachedLeft =3120
                    LayoutCachedTop =60
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Width =0
                    Height =0
                    BorderColor =7169897
                    Name ="txtFocus"

                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =6570
                    Top =60
                    Width =1620
                    Height =285
                    FontSize =9
                    TabIndex =1
                    Name ="cmdVerify"
                    Caption ="Verify All Species"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =6570
                    LayoutCachedTop =60
                    LayoutCachedWidth =8190
                    LayoutCachedHeight =345
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =34
                    Left =1620
                    Top =60
                    Width =1440
                    Height =300
                    FontSize =9
                    BackColor =6108695
                    ForeColor =16252927
                    Name ="DetailID_Label"
                    Caption ="Detail"
                    FontName ="Franklin Gothic Book"
                    LayoutCachedLeft =1620
                    LayoutCachedTop =60
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =5910
                    Top =60
                    Width =600
                    Height =285
                    FontSize =9
                    TabIndex =2
                    Name ="cmdUndo"
                    Caption ="Reset"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =5910
                    LayoutCachedTop =60
                    LayoutCachedWidth =6510
                    LayoutCachedHeight =345
                End
            End
        End
        Begin Section
            Height =420
            BackColor =16252927
            Name ="Detail"
            AlternateBackColor =16777215
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =120
                    Top =60
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    BorderColor =7169897
                    Name ="SpeciesID"
                    ControlSource ="SpeciesID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Species.SpeciesID, Species.ShortName FROM Species ORDER BY Species.ShortN"
                        "ame;"
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    ScrollBars =2
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4140
                    Top =60
                    Width =3330
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =3
                    BorderColor =7169897
                    Name ="Comments"
                    ControlSource ="Comments"

                    LayoutCachedLeft =4140
                    LayoutCachedTop =60
                    LayoutCachedWidth =7470
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3120
                    Top =60
                    Width =960
                    Height =300
                    FontSize =10
                    TabIndex =2
                    BorderColor =7169897
                    Name ="Individuals"
                    ControlSource ="Individuals"
                    StatusBarText ="Number of individuals of this species present in the photo"
                    DefaultValue ="1"
                    ControlTipText ="Number of individuals of this species present in the photo"

                    LayoutCachedLeft =3120
                    LayoutCachedTop =60
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =7590
                    Top =60
                    Width =600
                    Height =300
                    TabIndex =4
                    Name ="cmdDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
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

                    LayoutCachedLeft =7590
                    LayoutCachedTop =60
                    LayoutCachedWidth =8190
                    LayoutCachedHeight =360
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1620
                    Top =60
                    Height =300
                    FontSize =10
                    TabIndex =1
                    BorderColor =7169897
                    ConditionalFormat = Begin
                        0x010000008a000000010000000100000000000000000000001400000000000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0053007000650063006900650073004900 ,
                        0x44005d00290000000000
                    End
                    Name ="DetailID"
                    ControlSource ="DetailID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DetectionDetails.DetailID, DetectionDetails.DetailText FROM DetectionDeta"
                        "ils ORDER BY DetectionDetails.DetailText;"
                    ColumnWidths ="0;1440"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =60
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    IMESentenceMode =3
                    Width =0
                    Height =0
                    TabIndex =5
                    BorderColor =7169897
                    Name ="ObsID"
                    ControlSource ="ObsID"
                    DefaultValue ="6"

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =6108695
            Name ="FormFooter"
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
    "The PhotoID form encountered an error."

Dim bCompare As Boolean

Private Sub cmdDelete_Click()
'Delete the current record

On Error GoTo ErrHandler

If Me.NewRecord Then
    Me.Undo
Else
    If bCompare Then
        Me.StatusID = 3
        Me.Dirty = False
        Call cmdVerify_Click
        Me.cmdUndo.Enabled = True
    Else
        DoCmd.RunCommand acCmdDeleteRecord
    End If
End If

Me.Requery
Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description

End Sub

Private Sub cmdUndo_Click()
'Reset all detections to 'pending'

Dim strSQL As String

On Error GoTo ErrHandler

strSQL = "UPDATE Detections SET StatusID = 1 " & _
    "WHERE ImageID=" & Me.Parent.ImageID
DoCmd.SetWarnings False
DoCmd.RunSQL strSQL
Me.Requery
Me.txtFocus.SetFocus
Me.cmdUndo.Enabled = False

UndoExit:
    DoCmd.SetWarnings True
    Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description
    Resume UndoExit

End Sub

Private Sub cmdVerify_Click()
'Verify all detections for current photo

Dim db As Database
Dim rs As DAO.Recordset

On Error GoTo ErrHandler

'Get detections for this photo
Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM Detections WHERE (((ImageID)=" & _
    Me.ImageID & ") AND ((StatusID)<3))")

With rs
    .MoveFirst
    'Set all status fields to 2 (verified)
    Do Until .EOF
        .Edit
        rs!StatusID = 2
        .Update
        .MoveNext
    Loop
End With
Me.txtFocus.SetFocus
Me.Requery
Me.cmdVerify.Enabled = False
Me.Parent.txtFocus.SetFocus

VerifyExit:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Sub

ErrHandler:
    ErrorMsg "An error occurred during verification.", _
        Err.Number, Err.Description
    Resume VerifyExit

End Sub

Private Sub DetailID_Enter()
'Set the rowsource based on the SpeciesID

If IsNull(Me.SpeciesID) Then
    Me.DetailID.RowSource = ""
Else
    Me.DetailID.RowSource = "SELECT DetailID, DetailText " & _
        "FROM DetectionDetails WHERE (((DetectionDetails.SpeciesID)=" & _
        Me.SpeciesID & "))"
End If

End Sub

Private Sub DetailID_Exit(Cancel As Integer)
'Revert to unfiltered row source

Me.DetailID.RowSource = "SELECT DetailID, DetailText " & _
    "FROM DetectionDetails"

End Sub

Private Sub DetailID_NotInList(NewData As String, Response As Integer)
'Handles values in combo box that don't match details list

On Error GoTo ErrHandler

'Prompt to add new detail to lookup table
If MsgBox(NewData & " is not in the details list." & vbNewLine & _
    "Do you want to add it?", vbYesNo, "Unknown Detail") = vbYes Then
    'Allow new data in list and open form to add to species table
    Response = acDataErrAdded
    DoCmd.OpenForm "DetectionDetails", acNormal, , , , acDialog, _
        Me.SpeciesID & ";" & NewData
Else
    'User chose not to add, undo typed entry
    Response = acDataErrContinue
    Me.Undo
End If
Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description

End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
'Suppress delete record warning

On Error GoTo ErrHandler

Cancel = False

Response = acDataErrContinue
Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
'Verify new detection record in compare mode

If bCompare Then
    Me.StatusID = 2
End If

End Sub

Private Sub Form_Current()
'Show the verify button and prevent changes in compare mode

If bCompare Then
    EnableVerify
    EnableUndo
    LockFields
End If

End Sub

Private Sub Form_Open(Cancel As Integer)

CheckCompareMode

End Sub

Private Sub SpeciesID_AfterUpdate()
'Update Detail combo box

On Error GoTo ErrHandler

If IsNull(Me.SpeciesID) Then
    'Disable if no species selected
    Me.DetailID = Null
Else
    If Not IsNull(Me.DetailID) Then
        'Clear detail if it doesn't match species
        If Me.SpeciesID <> DLookup("SpeciesID", "DetectionDetails", _
            "DetailID=" & Me.DetailID) Then
            Me.DetailID = Null
        End If
    End If
End If

Exit Sub

ErrHandler:
    ErrorMsg "An error occurred retrieving details for this species.", _
        Err.Number, Err.Description

End Sub

Private Sub SpeciesID_NotInList(NewData As String, Response As Integer)
'Handle values in combo box that don't match species list

On Error GoTo ErrHandler

'Prompt to add new species to lookup table
If MsgBox(NewData & " is not in the species list." & vbNewLine & _
    "Do you want to add it?", vbYesNo, "Unknown Species") = vbYes Then
    'Allow new data in list and open form to add to species table
    Response = acDataErrAdded
    DoCmd.OpenForm "Species", acNormal, , , , acDialog, NewData
Else
    'User chose not to add, undo typed entry
    Response = acDataErrContinue
    Me.Undo
End If
Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description

End Sub

'------------------------------
'Form-specific custom functions
'------------------------------

Public Sub CheckCompareMode()
'Format form for compare IDs mode

bCompare = False
If (Nz(Me.Parent.RecordSource, "") = "PhotoIDCompare") Then
    bCompare = True
Else
    If (InStr(1, Me.Parent.Filter, "ImageID") > 0) Then
        bCompare = (DLookup("ObsCount", "Photos", _
            Me.Parent.Filter) > 1)
    End If
End If

If bCompare Then
    'Show verify button
    Me.cmdVerify.Visible = True
    Me.cmdUndo.Visible = True
    'Set up conditional formatting for species combo box
    Me.SpeciesID.ForeColor = vbBlack      'Val("&H" & "ED1C24")
    Me.SpeciesID.FormatConditions.Delete
    Me.SpeciesID.FormatConditions.Add acExpression, , "[StatusID]=1"
    Me.SpeciesID.FormatConditions(0).ForeColor = vbRed    'Val("&H" & "22B14C")
Else
    'Not in compare mode - opposite of above
    Me.cmdVerify.Visible = False
    Me.cmdUndo.Visible = False
    Me.SpeciesID.ForeColor = vbBlack
    Me.SpeciesID.FormatConditions.Delete
End If

End Sub

Private Sub EnableUndo()

Me.cmdUndo.Enabled = (DCount("*", "Detections", _
    "ImageID=" & Me.Parent.ImageID & " AND StatusID=3") > 0)

End Sub

Private Sub EnableVerify()
'Enable the verify button when in compare mode with > 1 species/detail

Dim bShowVerify As Boolean
Dim db As Database
Dim rs As DAO.Recordset
Dim lngSpecies As Long
Dim lngDetail As Long

On Error GoTo ErrHandler

bShowVerify = False

If Me.RecordsetClone.RecordCount > 1 Then
    'Compare mode & multiple detections
    Set db = CurrentDb
    Set rs = Me.RecordsetClone
    With rs
        'Get first species other than none
        .MoveFirst
        Do Until rs.EOF
            If Not rs!SpeciesID = 0 Then
                lngSpecies = rs!SpeciesID
                lngDetail = Nz(rs!DetailID, 0)
                Exit Do
            End If
            .MoveNext
        Loop
        If lngSpecies > 0 Then
            'Check additional detections
            .MoveFirst
            Do Until .EOF
                If rs!SpeciesID > 0 And ((rs!SpeciesID <> _
                    lngSpecies) Or (Nz(rs!DetailID, 0) <> lngDetail)) Then
                    'Different species or detail, show verify button
                    bShowVerify = True
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
End If

Me.cmdVerify.Enabled = bShowVerify

SHVExit:
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrHandler:
    ErrorMsg strGenericError, Err.Number, Err.Description
    Resume SHVExit

End Sub

Private Sub LockFields()

Dim bLock As Boolean

bLock = Not Me.NewRecord

Me.SpeciesID.Locked = bLock
Me.DetailID.Locked = bLock
Me.Individuals.Locked = bLock

End Sub
