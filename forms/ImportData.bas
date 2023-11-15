Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =163
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =4560
    DatasheetFontHeight =11
    ItemSuffix =2
    Left =3495
    Top =2490
    Right =28545
    Bottom =15015
    RecSrcDt = Begin
        0xd85ee3ab4bcce440
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
        Begin Section
            Height =540
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =120
                    Width =4320
                    Height =315
                    Name ="txtSourceFile"

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =435
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

Private strSource As String
Private dblTimer As Double
Private lngCount As Long

Private Sub Form_Open(Cancel As Integer)
'Import data from previous version of database

Dim iVersion As Integer

On Error GoTo ErrHandler

dblTimer = Timer

If IsNull(Me.OpenArgs) Then
    Cancel = True
    Exit Sub
End If

strSource = Me.OpenArgs
Me.txtSourceFile = strSource

If Not TestSource(strSource) Then
    MsgBox "Invalid file"
    Cancel = True
    Exit Sub
End If

ShowCaption "Determining source database version..."

iVersion = GetVersion(strSource)
Select Case iVersion
    Case 0
        MsgBox "Unrecognized file"
        Cancel = True
        Exit Sub
    Case 2
        If ReformatDetections(strSource) Then
            If ReformatTables(strSource) Then
                ImportData strSource
            End If
        Else
            lngCount = -1
        End If
    Case 3, 4
        ImportData strSource
End Select

RestoreSwitchboard True
DoCmd.SelectObject acForm, "Maintenance", False
DoCmd.Restore
ShowCaption "Data import complete."

If lngCount >= 0 Then
    MsgBox lngCount & " records imported from " & strSource & _
        vbNewLine & vbNewLine & Timer - dblTimer & " seconds elapsed"
End If

Exit Sub

ErrHandler:
    ErrorMsg "An error occurred importing the data." & vbNewLine & _
        vbNewLine & "Download a fresh copy of the latest version, " & _
        "compact and repair your old copy, and try again.", _
        Err.Number, Err.Description

End Sub

'------------------------------
'Form-specific custom functions
'------------------------------

Private Function AddDeletedDetection(strInput As String, _
    lngImg As Long, rs As DAO.Recordset) As String
'Add deleted detection and return observer ID

Dim vObs As Variant
Dim vSp As Variant

vObs = GetObserverID(strInput, "sObservers")
vSp = GetSpeciesID(strInput, "sSpecies")

If IsNull(vObs) Or IsNull(vSp) Then
    AddDeletedDetection = ""
Else
    rs.AddNew
    rs!ObsID = vObs
    rs!SpeciesID = vSp
    rs!Individuals = 1
    rs!StatusID = 3
    rs!ImageID = lngImg
    rs.Update
    AddDeletedDetection = CStr(vObs)
End If

End Function

Private Sub AddOtherDetection(lngImg As Long, _
    lngSp As Long, vDetail As Variant, lngObs As Long, _
    lngInd As Long, strComments As String, _
    rs As DAO.Recordset, bVer As Boolean)

rs.AddNew
rs!ObsID = lngObs
rs!SpeciesID = lngSp
rs!DetailID = vDetail
rs!Individuals = lngInd
rs!ImageID = lngImg
If bVer Then
    rs!StatusID = 2
End If
If Len(strComments) > 0 Then
    rs!Comments = strComments
End If
rs.Update

End Sub

Private Function GetCommonName(strCPart As String) As String
'Returns common name from a section of a detection comment

Dim iPos As Integer             'Position of character
Dim strCN As String             'Common Name

strCN = ""

'Get : position
iPos = InStr(1, strCPart, ":")
'Test for validity
If iPos > 0 Then
    'Take everything after :
    strCN = Mid(strCPart, iPos + 2)
    'Get - position
    iPos = InStr(1, strCN, " -")
    If iPos = 0 Then
        'Get ) position
        iPos = InStr(1, strCN, ")")
    End If
    If iPos > 0 Then
        'Take everything before end character
        strCN = Left(strCN, iPos - 1)
    End If
End If

'Check for apostrophe
iPos = InStr(1, strCN, "'")
If iPos > 0 Then
    'Do Until iPos = 0
        strCN = Replace(strCN, "'", "' & Chr(39) & '")
        'iPos = InStr(1, strCN, "'")
    'Loop
End If

'Return string
GetCommonName = strCN

End Function

Private Function GetObserverID(strIn As String, strTbl As String) _
    As Variant
'Translate intitials into observer ID

GetObserverID = DLookup("ObserverID", strTbl, "Initials='" & _
    Replace(GetObserverInitials(strIn), "'", "''") & "'")

End Function

Private Function GetObserverInitials(strCPart As String) As String
'Returns observer initials from a section of a detection comment

Dim iPos As Integer             'Position of character
Dim strInitials As String       'Observer Initials

strInitials = ""

'Get : position
iPos = InStr(1, strCPart, ":")
'Test for validity
If iPos > 0 Then
    'Take everything left of :
    strInitials = Left(strCPart, iPos - 1)
    'Get ( position
    iPos = InStrRev(strInitials, "(")
    'Take everything after (
    strInitials = Mid(strInitials, iPos + 1)
End If

'Return string
GetObserverInitials = strInitials

End Function

Private Function GetOriginalComments(strIn As String) As String

Dim strResult As String
Dim iPos As String

strResult = strIn
iPos = InStr(1, strResult, "(")
If iPos > 0 Then
    strResult = Left(strResult, iPos - 1)
End If

GetOriginalComments = strResult

End Function

Private Function GetSpeciesID(strIn As String, strTbl As String) _
    As Variant
'Translate common name into species ID

GetSpeciesID = DLookup("SpeciesID", strTbl, "CommonName='" & _
    Replace(GetCommonName(strIn), "'", "''") & "'")

End Function

Private Function GetTables() As String()

Dim strResult(0 To 16) As String

strResult(0) = "StudyAreas"
strResult(1) = "CameraLocations"
strResult(2) = "Visits"
strResult(3) = "Photos"
strResult(4) = "Observers"
strResult(5) = "ObsPhone"
strResult(6) = "SpeciesGroups"
strResult(7) = "Species"
strResult(8) = "SpeciesShortcuts"
strResult(9) = "DetectionDetails"
strResult(10) = "DetailShortcuts"
strResult(11) = "Detections"
strResult(12) = "PhotoTags"
strResult(13) = "Individuals"
strResult(14) = "IndependentDetections"
strResult(15) = "IndDetGroups"
strResult(16) = "IndividualDetections"

GetTables = strResult

Erase strResult

End Function

Private Function GetVersion(strFile As String) As Integer

Dim db As Database
Dim tdf As TableDef
Dim fld As Field

On Error GoTo ErrHandler

If LinkTable(strFile, "Detections") Then
    Set db = CurrentDb
    Set tdf = db.TableDefs("sDetections")
    For Each fld In tdf.Fields
        If fld.Name = "ObsID2" Then
            GetVersion = 2
            Exit For
        Else
            If fld.Name = "StatusID" Then
                GetVersion = 3
                Exit For
            End If
        End If
    Next
    Set tdf = Nothing
    Set db = Nothing
End If

If GetVersion = 3 Then
    If LinkTable(strFile, "Visits") Then
        Set db = CurrentDb
        Set tdf = db.TableDefs("sVisits")
        For Each fld In tdf.Fields
            If fld.Name = "SetVisitID" Then
                GetVersion = 4
                Exit For
            End If
        Next
    End If
End If

GVExit:
    On Error Resume Next
    Set tdf = Nothing
    Set db = Nothing
    DoCmd.DeleteObject acTable, "sDetections"
    DoCmd.DeleteObject acTable, "sVisits"
    Exit Function

ErrHandler:
    GetVersion = 0
    Resume GVExit

End Function

Private Sub ImportData(strFile As String)

Dim strTables() As String
Dim i As Integer

strTables = GetTables()
For i = 0 To UBound(strTables)
    ShowCaption "Importing " & strTables(i) & "..."
    ImportTable strFile, strTables(i)
Next

Erase strTables

End Sub

Private Sub ImportTable(strFile As String, strTable As String)

Dim db As Database
Dim strSQL As String

If LinkTable(strFile, strTable) Then
    Set db = CurrentDb
    strSQL = "INSERT INTO " & strTable & " SELECT * FROM s" & _
        strTable & ";"
    db.Execute strSQL
    lngCount = lngCount + db.RecordsAffected
    If strTable = "Visits" Then
        db.Execute strSQL
        lngCount = lngCount + db.RecordsAffected
    End If
    DoCmd.DeleteObject acTable, "s" & strTable
    Set db = Nothing
End If

End Sub

Private Function LinkTable(strFile As String, strTable As String) As Boolean

On Error Resume Next

If TableExists("s" & strTable) Then
    LinkTable = True
    Exit Function
End If

DoCmd.TransferDatabase acLink, "Microsoft Access", strFile, acTable, _
    strTable, "s" & strTable
LinkTable = TableExists("s" & strTable)

End Function

Private Function ReformatDetections(strFile As String) As Boolean

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim rsDest As DAO.Recordset
Dim strComments As String
Dim strObs As String

On Error GoTo ErrHandler

ShowCaption "Converting detections to new format..."

Set db = CurrentDb

If LinkTable(strFile, "Observers") And LinkTable(strFile, "Species") Then
    DoCmd.TransferDatabase acLink, "Microsoft Access", strFile, acTable, _
        "Detections", "xDetections"
    DoCmd.CopyObject , "sDetections", acTable, "Detections"
    db.Execute "DELETE * FROM sDetections"
    Set rs = db.OpenRecordset("xDetections")
    Set rsDest = db.OpenRecordset("sDetections")
    If Not (rs.EOF And rs.BOF) Then
        'Start with first record
        rs.MoveFirst
        Do Until rs.EOF
            'Reset strings
            strComments = Nz(rs!Comments, "")
            strObs = ""
            'Add records from comments
            Do Until InStr(1, strComments, ":") = 0
                strObs = strObs & AddDeletedDetection(strComments, _
                    rs!ImageID, rsDest)
                strComments = Replace(strComments, ":", "", 1, 1)
            Loop
            'Get new comment string
            strComments = GetOriginalComments(Nz(rs!Comments, ""))
            'Check second observer
            If Not IsNull(rs!ObsID2) Then
                If InStr(1, strObs, CStr(rs!ObsID2)) = 0 Then
                    'Add verified record for second observer
                    AddOtherDetection rs!ImageID, rs!SpeciesID, _
                        rs!DetailID, rs!ObsID2, rs!Individuals, _
                        strComments, rsDest, True
                End If
                'Add verified record for first observer
                AddOtherDetection rs!ImageID, rs!SpeciesID, _
                    rs!DetailID, rs!ObsID, rs!Individuals, _
                    strComments, rsDest, True
            Else
                'Add pending record for first observer
                AddOtherDetection rs!ImageID, rs!SpeciesID, _
                    rs!DetailID, rs!ObsID, rs!Individuals, _
                    strComments, rsDest, False
            End If
            'Move to next detection
            rs.MoveNext
        Loop
    End If
End If

ReformatDetections = True

RDExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not rsDest Is Nothing Then
        rsDest.Close
        Set rsDest = Nothing
    End If
    Set db = Nothing
    DoCmd.DeleteObject acTable, "xDetections"
    If Err.Number > 0 Then
        DoCmd.DeleteObject acTable, "sDetections"
        DoCmd.DeleteObject acTable, "sSpecies"
        DoCmd.DeleteObject acTable, "sObservers"
    End If
    Exit Function

ErrHandler:
    MsgBox "Converting IDs failed - import canceled."
    Resume RDExit

End Function

Private Function ReformatTables(strFile As String) As Boolean

Dim db As Database
Dim tdf As TableDef

On Error Resume Next

DoCmd.DeleteObject acTable, "sObsPhone"
DoCmd.DeleteObject acTable, "sPhotos"

On Error GoTo ErrHandler

DoCmd.TransferDatabase acImport, "Microsoft Access", strFile, acTable, _
    "ObsPhone", "sObsPhone"
DoCmd.TransferDatabase acImport, "Microsoft Access", strFile, acTable, _
    "Photos", "sPhotos"

Set db = CurrentDb

Set tdf = db.TableDefs("sObsPhone")

On Error Resume Next

tdf.Fields.Delete "upsize_ts"
If Err.Number = 3265 Then
    Err.Clear
End If

On Error GoTo ErrHandler

Set tdf = db.TableDefs("sPhotos")

On Error Resume Next

tdf.Fields.Delete "NeedsID"
If Err.Number = 3265 Then
    Err.Clear
End If

ReformatTables = True

RTExit:
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    ErrorMsg "Source tables could not be modified to current format.", _
        Err.Number, Err.Description, "Import Cancelled"
    Resume RTExit

End Function

Private Sub ShowCaption(strText As String)

Dim frm As Form

If IsLoaded("Maintenance") Then
    Set frm = Forms("Maintenance")
    frm.lblStatus.Caption = strText
    frm.Repaint
End If

Set frm = Nothing

End Sub

Private Function TableExists(strTable As String) As Boolean

Dim db As Database
Dim tdf As TableDef

Set db = CurrentDb

For Each tdf In db.TableDefs
    If tdf.Name = strTable Then
        TableExists = True
        Exit For
    End If
Next

Set tdf = Nothing
Set db = Nothing

End Function

Private Function TestSource(strFile As String) As Boolean

TestSource = Len(Dir(strFile)) > 0

If TestSource Then
    TestSource = Right(strFile, 6) = ".accdb"
End If

End Function
