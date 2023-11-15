Attribute VB_Name = "NewFunctions2016"
Option Compare Database
Option Explicit

Public Function ExportKML(strQueryName As String, _
    strFileName As String, strLatField As String, _
    strLongField As String, Optional bStayOpen As Boolean, _
    Optional strNameField As String = "", _
    Optional strDescField As String = "") As Boolean

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim i As Integer
Dim j As Integer
Dim iFile As Integer
Dim strName As String
Dim strDesc As String

On Error GoTo ErrHandler

Set db = CurrentDb
Set rs = db.OpenRecordset(strQueryName)

'Test for valid recordset with lat long fields
If rs.BOF And rs.EOF Then
    GoTo ExportExit
Else
    If TestRSField(rs, strLatField) = False Or _
        TestRSField(rs, strLongField) = False Then
        GoTo ExportExit
    Else
        rs.MoveFirst
    End If
End If

'Test name and description fields
If Len(strNameField) > 0 Then
    If TestRSField(rs, strNameField) = False Then
        strNameField = ""
    End If
End If
If Len(strDescField) > 0 Then
    If TestRSField(rs, strDescField) = False Then
        strDescField = ""
    End If
End If

'Delete existing file if present
On Error Resume Next
Kill strFileName
On Error GoTo ErrHandler
'Get text file name and open it for writing
iFile = FreeFile
Open strFileName For Output As iFile

Print #iFile, "<?xml version=""1.0"" encoding=""UTF-8""?>"
Print #iFile, "<kml xmlns=""http://earth.google.com/kml/2.1"">"
Print #iFile, "<Document>"

With rs
    'Loop through recordset
    Do Until .EOF
        i = i + 1
        'Create new placemark
        Print #iFile, "   <Placemark>"
        'Get name
        strName = ""
        If Len(strNameField) > 0 Then
            If Not IsNull(.Fields(strNameField)) Then
                strName = .Fields(strNameField)
            End If
        End If
        If Len(strName) = 0 Then
            strName = "Location" & Format(i, "000")
        End If
        Print #iFile, "      <name>" & strName & "</name>"
        'Get description
        If Len(strDescField) > 0 Then
            If IsNull(.Fields(strDescField)) Then
                strDesc = "No Description"
            Else
                strDesc = XMLString(.Fields(strDescField))
            End If
        Else
            strDesc = ""
            For j = 0 To .Fields.Count - 1
                strDesc = strDesc & .Fields(j).Name & ": " & _
                    XMLString(.Fields(j)) & "<br/>"
            Next
            strDesc = Left(strDesc, Len(strDesc) - 5)
        End If
        Print #iFile, "      <description>" & strDesc & "</description>"
        'Create point with coordinates
        Print #iFile, "      <Point>"
        Print #iFile, "         <coordinates>" & .Fields(strLongField) & "," & _
            .Fields(strLatField) & "</coordinates>"
        Print #iFile, "      </Point>"
        Print #iFile, "   </Placemark>"
        .MoveNext
    Loop
End With

Print #iFile, "</Document>"
Print #iFile, "</kml>"

ExportKML = True

ExportExit:
    On Error Resume Next
    Close #iFile
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    'Open the file in google earth
    If bStayOpen Then
        Shell "RUNDLL32.EXE URL.DLL,FileProtocolHandler " & _
            strFileName, vbNormalFocus
    End If
    Exit Function

ErrHandler:
    ErrorMsg "An error occurred creating the kml file.", _
        Err.Number, Err.Description
    Resume ExportExit

End Function

Private Function TestRSField(rs As DAO.Recordset, strField As String) As Boolean

Dim i As Integer

For i = 0 To rs.Fields.Count - 1
    If rs.Fields(i).Name = strField Then
        TestRSField = True
        Exit For
    End If
Next

End Function

Public Sub UpdateFlags()

Dim db As DAO.Database
Dim rs As DAO.Recordset

Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT ImageID FROM Photos " & _
    "WHERE (((NeedsUpdate)=True))")

With rs
    If Not (.BOF And .EOF) Then
        .MoveFirst
        Do Until .EOF
            UpdateSinglePhoto !ImageID
            .MoveNext
        Loop
    End If
    .Close
End With

UFExit:
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrHandler:
    ErrorMsg "An error occurred processing IDs.", _
        Err.Number, Err.Description
    Resume UFExit

End Sub

Public Sub UpdateSinglePhoto(lngImageID As Long)

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim bCompare As Boolean
Dim iObs As Integer
Dim bVer As Boolean
Dim iPObs As Integer
Dim bNotNone As Boolean
Dim bMultiSp As Boolean

'Get observer count and status and compare flags
Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT ObsID, Sum(IIf(StatusID=1,1,0)) " & _
    "AS PCount, Sum(IIf(StatusID=2,1,0)) As VCount, " & _
    "Max(IIf(StatusID<3,SpeciesID,0)) As MaxS, " & _
    "Min(IIf(StatusID<3 And SpeciesID>0,SpeciesID,999)) As MinS " & _
    "FROM Detections WHERE ImageID=" & lngImageID & _
    " GROUP BY ObsID")
If Not (rs.BOF And rs.EOF) Then
    rs.MoveFirst
    Do Until rs.EOF
        If rs!MaxS > 0 Then
            bNotNone = True
            If rs!MinS < 999 And rs!MaxS <> rs!MinS Then
                bMultiSp = True
            End If
        End If
        If rs!VCount > 0 Then
            bVer = True
        End If
        If rs!PCount > 0 Then
            iPObs = iPObs + 1
        End If
        rs.MoveNext
    Loop
    iObs = rs.RecordCount
    If iPObs > 0 Then
        bCompare = bVer Or (iPObs > 1)
    End If
End If
rs.Close

'Check for multiple verified species
If Not bMultiSp Then
    Set rs = db.OpenRecordset("SELECT SpeciesID FROM Detections " & _
        "WHERE ImageID= " & lngImageID & " AND StatusID=2 " & _
        "GROUP BY SpeciesID")
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
        bMultiSp = (rs.RecordCount > 1)
    End If
    rs.Close
End If

Set rs = db.OpenRecordset("SELECT * FROM Photos " & _
    "WHERE ImageID=" & lngImageID)
rs.MoveFirst
rs.Edit
rs!Compare = bCompare
rs!ObsCount = iObs
rs!Verified = bVer
rs!Pending = (iPObs > 0)
rs!MultiSp = bMultiSp
rs!NotNone = bNotNone
rs!NeedsUpdate = False
rs.Update
rs.Close

Set rs = Nothing
Set db = Nothing

End Sub

Public Sub UpdateVisitDependencies( _
    Optional vLocationID As Variant = Null)

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim strSQL As String
Dim vSet As Variant
Dim vLoc As Variant

On Error GoTo ErrHandler

Set db = CurrentDb

strSQL = "SELECT * FROM Visits "

If Not IsNull(vLocationID) Then
    strSQL = strSQL & "WHERE (((LocationID) = " & _
        vLocationID & ")) "
End If

strSQL = strSQL & "ORDER BY LocationID, VisitDate, " & _
    "[VisitTypeID] Mod 3;"

Set rs = db.OpenRecordset(strSQL)

With rs
    If Not (.BOF And .EOF) Then
        .MoveFirst
        Do Until .EOF
            If IsNull(vLoc) Or (vLoc <> !LocationID) Then
                vLoc = !LocationID
                vSet = Null
            End If
            .Edit
            Select Case !VisitTypeID
                Case 1
                    !SetVisitID = vSet
                Case 2
                    !SetVisitID = vSet
                    vSet = Null
                Case 3
                    !SetVisitID = Null
                    vSet = !VisitID
            End Select
            .Update
            .MoveNext
        Loop
    End If
End With

UVDExit:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Sub

ErrHandler:
    ErrorMsg "An error occurred associating check " & _
        "and pull visits with sets.", Err.Number, Err.Description
    Resume UVDExit

End Sub

Public Function XMLString(vStringIn As Variant) As String

Dim strResult As String

If Not IsNull(vStringIn) Then
    strResult = Replace(vStringIn, chr(38), chr(38) & "amp;")
    strResult = Replace(strResult, chr(34), chr(38) & "quot;")
    strResult = Replace(strResult, chr(39), chr(38) & "apos;")
    strResult = Replace(strResult, chr(60), chr(38) & "lt;")
    strResult = Replace(strResult, chr(62), chr(38) & "gt;")
    XMLString = strResult
Else
    XMLString = ""
End If

End Function
