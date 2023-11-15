Attribute VB_Name = "CustomFunctions"
Option Compare Database
Option Explicit

Dim strSubfolders() As String
Dim iFolder As Integer
Dim iCount As Long
Dim strLocalTables(0 To 14, 0 To 1) As String

Public Const msoFileDialogSaveAs As Integer = 2
Public Const msoFileDialogFilePicker As Integer = 3
Public Const msoFileDialogFolderPicker As Integer = 4

'Code in this module written by Eric Newkirk for Colorado
    'Parks and Wildlife, copyright 2015

'This program is free software: you can redistribute it and/or modify
'it under the terms of the included license.  To view the license
'click the credits link on the startup form then click license
'agreement.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'included license agreement for more details.

Public Sub AddModuleIndexes(strTargetDB As String)
'Adds indexes to the tables in the module DB

'Create PK Indexes in new DB
createIndex strTargetDB, "Detections", "DetectionID", True
createIndex strTargetDB, "PhotoTags", "TagID", True
createIndex strTargetDB, "Photos", "ImageID", True
createIndex strTargetDB, "Visits", "VisitID", True
createIndex strTargetDB, "CameraLocations", "LocationID", True
createIndex strTargetDB, "Observers", "ObserverID", True
createIndex strTargetDB, "Species", "SpeciesID", True
createIndex strTargetDB, "CurrentObserver", "ObserverID", True
createIndex strTargetDB, "Help", "HelpID", True
createIndex strTargetDB, "lkupPhoneTypes", "PhoneTypeID", True
createIndex strTargetDB, "ObsPhone", "ID", True
createIndex strTargetDB, "SpeciesShortcuts", "SpeciesID", True
createIndex strTargetDB, "StudyAreas", "StudyAreaID", True
createIndex strTargetDB, "DetectionDetails", "DetailID", True
createIndex strTargetDB, "DetailShortcuts", "DetailID", True

'Create other indexes
createIndex strTargetDB, "Detections", "ImageID"
createIndex strTargetDB, "Detections", "SpeciesID"
createIndex strTargetDB, "Detections", "DetailID"
createIndex strTargetDB, "PhotoTags", "ImageID"
createIndex strTargetDB, "Photos", "VisitID"
createIndex strTargetDB, "Visits", "LocationID"
createIndex strTargetDB, "CameraLocations", "StudyAreaID"
createIndex strTargetDB, "ObsPhone", "ObserverID"
createIndex strTargetDB, "ObsPhone", "TypeID"

End Sub

Public Sub AddModuleRelations(strTargetDB As String)
'Relates tables in module DB

CreateRelation "Photos", "ImageID", "Detections", "ImageID", strTargetDB
CreateRelation "Photos", "ImageID", "PhotoTags", "ImageID", strTargetDB
CreateRelation "Visits", "VisitID", "Photos", "VisitID", strTargetDB
CreateRelation "CameraLocations", "LocationID", "Visits", "LocationID", _
    strTargetDB
CreateRelation "Observers", "ObserverID", "Detections", "ObsID", strTargetDB
CreateRelation "Observers", "ObserverID", "PhotoTags", "ObsID", strTargetDB
CreateRelation "Observers", "ObserverID", "CurrentObserver", _
    "ObserverID", strTargetDB
CreateRelation "Observers", "ObserverID", "ObsPhone", "ObserverID", strTargetDB
CreateRelation "lkupPhoneTypes", "PhoneTypeID", "ObsPhone", "TypeID", _
    strTargetDB
CreateRelation "StudyAreas", "StudyAreaID", "CameraLocations", _
    "StudyAreaID", strTargetDB
CreateRelation "Species", "SpeciesID", "Detections", "SpeciesID", strTargetDB
CreateRelation "Species", "SpeciesID", "SpeciesShortcuts", "SpeciesID", _
    strTargetDB
CreateRelation "Species", "SpeciesID", "DetectionDetails", "SpeciesID", _
    strTargetDB
CreateRelation "DetectionDetails", "DetailID", "Detections", "DetailID", _
    strTargetDB
CreateRelation "DetectionDetails", "DetailID", "DetailShortcuts", "DetailID", _
    strTargetDB

End Sub

Public Function AddSlash(strFolder As String) As String
'Add slash to end of folder path

If Right(strFolder, 1) = "\" Then
    AddSlash = strFolder
Else
    AddSlash = strFolder & "\"
End If

End Function

Public Function BuildRuntimeDB(strPath As String) As Boolean
'Creates a new Photo ID module database

Dim db As Database
Dim strTempDB As String
Dim strDestDB As String
Dim db2 As Database
Dim appAccess As Access.Application
Dim prop As Property

On Error GoTo ErrHandler

BuildRuntimeDB = False

'Create blank database for new module
Set db = CurrentDb
strTempDB = Replace(db.Name, ".accdb", "temp.accdb")
If Len(Dir(strTempDB)) > 0 Then
    Kill strTempDB
End If
Set appAccess = New Access.Application
appAccess.NewCurrentDatabase strTempDB
'Add reference to allow opening jpg files
appAccess.References.AddFromFile "C:\WINDOWS\system32\scrrun.dll"
appAccess.CloseCurrentDatabase
appAccess.Quit
Set appAccess = Nothing

'Set startup form and name autocorrect properties for new database
Set db2 = DBEngine.OpenDatabase(strTempDB)
Set prop = db2.CreateProperty("StartupForm", dbText, "PhotoID")
db2.Properties.Append prop
Set prop = db2.CreateProperty("Track Name AutoCorrect Info", _
    dbInteger, 0)
db2.Properties.Append prop
db2.Close
Set db2 = Nothing

'Add tables to new database
TransferLocalTables strTempDB

'Add other objects
TransferModuleObjects strTempDB

'Index tables
AddModuleIndexes strTempDB

'Create relationships in new DB
AddModuleRelations strTempDB

'Compact new database and change extension
strDestDB = strPath & "\PhotoID.accdb"
If Dir(strDestDB) <> "" Then
    Kill strDestDB
End If
Application.CompactRepair strTempDB, strDestDB
Kill strTempDB
strTempDB = strDestDB
strDestDB = Replace(strDestDB, "accdb", "accdr")
If Dir(strDestDB) <> "" Then
    Kill strDestDB
End If
Name strTempDB As strDestDB

'Remove temporary tables
DeleteLocalTables

BuildRuntimeDB = True

ModuleExit:
    Set db = Nothing
    Set db2 = Nothing
    Set appAccess = Nothing
    Erase strLocalTables
    Exit Function

ErrHandler:
    ErrorMsg "Module creation failed due to an error.", _
        Err.Number, Err.Description
    Resume ModuleExit

End Function

Public Function CameraIsActive(lngLocation As Long, dCheckStart As Date, _
    dCheckEnd As Date) As Boolean
'Checks if a camera was active for an entire occasion

Dim db As Database
Dim rsVisits As DAO.Recordset
Dim dStart As Date
Dim dEnd As Date

CameraIsActive = False
If DCount("*", "Visits", "LocationID=" & lngLocation & _
    " And ActiveStart<=#" & USDate(dCheckStart) & "# And ActiveEnd>=#" & _
    USDate(dCheckEnd) & "#") > 0 Then
    CameraIsActive = True
    Exit Function
End If

Set db = CurrentDb
Set rsVisits = db.OpenRecordset("SELECT * FROM Visits WHERE (((" & _
    "Visits.LocationID)=" & lngLocation & ") AND ((" & _
    "Visits.VisitTypeID)<3)) ORDER BY Visits.VisitDate")
If rsVisits.EOF And rsVisits.BOF Then
    GoTo CAExit
Else
    rsVisits.MoveFirst
    Do Until rsVisits.EOF
        If Not (IsNull(rsVisits!ActiveStart) Or _
            IsNull(rsVisits!ActiveEnd)) Then
            dStart = rsVisits!ActiveStart
            dEnd = rsVisits!ActiveEnd
            rsVisits.MoveNext
            Do Until rsVisits.EOF
                If rsVisits!ActiveStart - dEnd < 0.5 Then
                    dEnd = rsVisits!ActiveEnd
                Else
                    Exit Do
                End If
                rsVisits.MoveNext
            Loop
            If dStart <= dCheckStart And dEnd >= dCheckEnd Then
                CameraIsActive = True
                Exit Do
            End If
        Else
            rsVisits.MoveNext
        End If
    Loop
End If

CAExit:
    If Not rsVisits Is Nothing Then
        rsVisits.Close
        Set rsVisits = Nothing
    End If
    Set db = Nothing

End Function

Public Function CheckImageFile(strFile As String) As String
'Returns a blank string if file is found, X if not

On Error Resume Next

CheckImageFile = ""

If Not FileExists(strFile) Then
    CheckImageFile = "X"
End If

End Function

Public Sub CombineDetections()
'Compare detections and apply status when available

On Error Resume Next

DoCmd.SetWarnings False
DeleteDuplicateDetections
DoCmd.OpenQuery "qryVerifyMatchingDetections"
DoCmd.OpenQuery "qryVerifyMatchingDetectionsNoDetail"
DoCmd.SetWarnings True

End Sub

Public Sub CombineDetectionsSingle(lngImageID As Long)
'Compare detections and apply status when available

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim vBookmark As Variant

Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM Detections " & _
    "WHERE ImageID=" & lngImageID & " AND StatusID<3;")

With rs
    If Not (.BOF And .EOF) Then
        .MoveFirst
        Do Until .EOF
            If !StatusID = 1 Then
                vBookmark = .Bookmark
                .FindFirst "SpeciesID=" & !SpeciesID & _
                    " AND Nz(DetailID,0)=" & Nz(!DetailID, 0) & _
                    " AND Individuals=" & !Individuals & _
                    " AND ObsID<>" & !ObsID
                If Not .NoMatch Then
                    .Edit
                    !StatusID = 2
                    .Update
                    .Bookmark = vBookmark
                    .Edit
                    !StatusID = 2
                    .Update
                Else
                    .Bookmark = vBookmark
                End If
            End If
            .MoveNext
        Loop
    End If
    .Close
End With

Set rs = Nothing
Set db = Nothing

End Sub

Public Function CombineStrings(vStr1 As Variant, vStr2 As Variant) _
    As String
'Concatenates two strings

Dim strResult As String

strResult = ""
If Len(vStr1) > 0 Then
    strResult = vStr1
    If Len(vStr2) > 0 Then
        strResult = strResult & "; " & vStr2
    End If
Else
    If Len(vStr2) > 0 Then
        strResult = vStr2
    End If
End If

CombineStrings = strResult

End Function

Public Function CountImages(MainFolder As String, _
    IncludeSub As Boolean) As Long
'Returns number of photos in main folder and subfolders

iCount = 0
CountImages = CountImagesSubRoutine(MainFolder, IncludeSub)

End Function

Public Function CountImagesSubRoutine(MainFolder As String, _
    IncludeSub As Boolean) As Long
'Returns number of photos

Dim fso As Object
Dim pFolder As Object
Dim pSubfolders As Object
Dim pSubFolder As Object
Dim strFile As String

strFile = Dir(MainFolder & "\*jpg")
If Not strFile = "" Then
    Do Until strFile = ""
        iCount = iCount + 1
        strFile = Dir()
    Loop
End If

If IncludeSub Then
    Set fso = CreateObject("scripting.FileSystemObject")
    Set pFolder = fso.GetFolder(MainFolder)
    Set pSubfolders = pFolder.SubFolders
    For Each pSubFolder In pSubfolders
        Call CountImagesSubRoutine(pSubFolder.Path, True)
    Next
End If

Set fso = Nothing
Set pFolder = Nothing
Set pSubfolders = Nothing
Set pSubFolder = Nothing

CountImagesSubRoutine = iCount

End Function

Public Sub CreateLocalTables(strDBPath As String)
'Creates local copies of all the tables needed for a
'Photo ID module

Dim strSQL As String
Dim i As Integer
Dim db As DAO.Database

On Error GoTo ErrHandler

DoCmd.SetWarnings False
'Create tables with special conditions first
'COLocal is empty
strSQL = "SELECT CurrentObserver.* INTO COLocal FROM " & _
    "CurrentObserver WHERE (((CurrentObserver.ObserverID)=0));"
DoCmd.RunSQL strSQL

'CPDLocal is empty
strSQL = "SELECT Detections.* INTO CPDLocal " & _
    "FROM Detections WHERE (((Detections.DetectionID) Is Null));"
DoCmd.RunSQL strSQL
'Add default for individuals and status
CurrentDb.TableDefs("CPDLocal").Fields("Individuals").DefaultValue = 1
CurrentDb.TableDefs("CPDLocal").Fields("StatusID").DefaultValue = 1

'PTLocal is empty
strSQL = "SELECT PhotoTags.* INTO PTLocal " & _
    "FROM PhotoTags WHERE (((PhotoTags.TagID) Is Null));"
DoCmd.RunSQL strSQL

'Photos limited by PhotoViewerQuery
strSQL = "SELECT Photos.* INTO CPLocal " & _
    "FROM PhotoViewerQuery INNER JOIN Photos ON " & _
    "PhotoViewerQuery.ImgID = Photos.ImageID;"
DoCmd.RunSQL strSQL
'Update FilePath field
strSQL = "UPDATE CPLocal SET CPLocal.FilePath = '" & strDBPath & _
    "\Photos\';"
DoCmd.RunSQL strSQL

'Visits limited by Photos
strSQL = "SELECT DISTINCT Visits.* INTO CVLocal " & _
    "FROM Visits INNER JOIN CPLocal ON Visits.VisitID " & _
    "= CPLocal.VisitID;"
DoCmd.RunSQL strSQL
'Add set visits for season filter
strSQL = "INSERT INTO CVLocal SELECT DISTINCT SetVisits.* FROM CVLocal " & _
    "INNER JOIN Visits As SetVisits ON " & _
    "CVLocal.SetVisitID = SetVisits.VisitID"
DoCmd.RunSQL strSQL

'Locations limited by Visits
strSQL = "SELECT DISTINCT CameraLocations.* INTO CLLocal " & _
    "FROM CameraLocations INNER JOIN CVLocal ON " & _
    "CameraLocations.LocationID = CVLocal.LocationID;"
DoCmd.RunSQL strSQL

'Create the rest of the tables
For i = 6 To 14
    strSQL = "SELECT " & strLocalTables(i, 0) & ".* INTO " & _
        strLocalTables(i, 1) & " FROM " & strLocalTables(i, 0) & ";"
    DoCmd.RunSQL strSQL
Next

'Set required fields
Set db = CurrentDb
db.TableDefs("ObsLocal").Fields("LastName").Required = True
db.TableDefs("ObsLocal").Fields("FirstName").Required = True
db.TableDefs("ObsLocal").Fields("Initials").Required = True
db.TableDefs("SpLocal").Fields("CommonName").Required = True
db.TableDefs("SpLocal").Fields("ShortName").Required = True
db.TableDefs("DDLocal").Fields("DetailText").Required = True
db.TableDefs("DDLocal").Fields("SpeciesID").Required = True

CLTExit:
    DoCmd.SetWarnings True
    Set db = Nothing
    Exit Sub

ErrHandler:
    ErrorMsg "An error occurred while creating local tables.", _
        Err.Number, Err.Description
    Resume CLTExit

End Sub

Public Sub CreateModuleTables(strPath As String)

'Create local copies of data tables
GetTableArray
CreateLocalTables strPath

End Sub

Public Function CreateOverflowFolder(strInput As String) As String
'Creates a new folder when the file limit for a single directory
'causes an error

Dim fso As Object
Dim i As Integer
Dim strNewFolder As String

Set fso = CreateObject("scripting.FileSystemObject")

'Check for existing overflow folder
If InStr(1, strInput, "_Overflow") > 0 Then
    i = CInt(Mid(strInput, InStrRev(strInput, "_Overflow") + 9)) + 1
    strNewFolder = Left(strInput, InStrRev(strInput, "_Overflow") + 8)
    'Add sequential number to folder name
    strNewFolder = strNewFolder & i
Else
    i = 1
    'Get name based on original folder
    strNewFolder = strInput & "_Overflow" & i
End If

If fso.FolderExists(strNewFolder) Then
    CreateOverflowFolder = ""
Else
    'Create folder and return path
    fso.CreateFolder (strNewFolder)
    CreateOverflowFolder = strNewFolder
End If

Set fso = Nothing

End Function

Public Sub CreateTempDatabase(strDBName As String)
'Creates a temporary database for activity pattern analysis

Dim appAccess As Access.Application
Dim db As DAO.Database
Dim prp As DAO.Property

On Error GoTo ErrHandler

'Check for existing activity database
If Dir(strDBName) <> "" Then
    Exit Sub
End If

'Create blank database for new module
Set appAccess = New Access.Application
appAccess.NewCurrentDatabase strDBName
appAccess.CloseCurrentDatabase
appAccess.Quit

'Set name autocorrect property in new database.
Set db = DBEngine.OpenDatabase(strDBName)
Set prp = db.CreateProperty("Track Name AutoCorrect Info", _
    dbLong, 0)
db.Properties.Append prp
db.Close

CAPDExit:
    Set appAccess = Nothing
    Set prp = Nothing
    Set db = Nothing
    Exit Sub

ErrHandler:
    Dim iError As Integer
    iError = Err
    Err.Clear
    Err.Raise iError
    Resume CAPDExit

End Sub

Public Sub DeleteDuplicateDetections()
'Removes duplicates from the detections table

Dim db As Database
Dim rs As Recordset
Dim strSQL As String

On Error Resume Next

'Get one ID for each set of duplicates
strSQL = "SELECT First(DetectionID) As DupID FROM Photos " & _
    "INNER JOIN Detections ON Photos.ImageID = Detections.ImageID " & _
    "WHERE (((Photos.NeedsUpdate) = True " & _
    "GROUP BY Photos.ImageID, Detections.SpeciesID, " & _
    "Detections.DetailID, Detections.Individuals, " & _
    "Detections.ObsID HAVING (((Count(Detections.DetectionID))>1))"
Set db = CurrentDb
Set rs = db.OpenRecordset(strSQL)

'Check for results
If Not (rs.BOF And rs.EOF) Then
    'There are duplicates, loop through results
    rs.MoveFirst
    Do Until rs.EOF
        'Prevent infinite loop on error
        If Err.Number > 0 Then
            Exit Do
        End If
        'Delete specified record
        strSQL = "DELETE * FROM Detections WHERE (((DetectionID)=" & _
            rs!DupID & "))"
        db.Execute strSQL
        'Move to next set of duplicates
        rs.MoveNext
    Loop
    'Call procedure again in case there are duplicates with > 2 records
    DeleteDuplicateDetections
End If

'Clean up
rs.Close
Set rs = Nothing
Set db = Nothing

End Sub

Public Sub DeleteLocalTables()
'Gets rid of the tables created with the
'CreateLocalTables function

Dim i As Integer

On Error GoTo ErrHandler

DoCmd.SetWarnings False
For i = 0 To 14
    DoCmd.DeleteObject acTable, strLocalTables(i, 1)
Next i

DLTExit:
    DoCmd.SetWarnings True
    Exit Sub

ErrHandler:
    If Err.Number = 7874 Then
        Err.Clear
        Resume Next
    Else
        ErrorMsg "An error occurred while deleting " & _
            "temporary tables.", Err.Number, Err.Description
        Resume DLTExit
    End If

End Sub

Public Function DiskSpace(strDriveLetter As String) As Double
'Returns the number of megabytes available on a storage drive

Dim fso As Object

On Error Resume Next

Set fso = CreateObject("scripting.FileSystemObject")
DiskSpace = fso.GetDrive(strDriveLetter).FreeSpace / 1000000
Set fso = Nothing

End Function

Public Function ExportCSV(strQueryName As String, _
    strFileName As String, Optional bStayOpen As Boolean) As Boolean
'Exports query results to csv
'Returns true if export is successful
'Field names included in header
'strFileName should be the destination file name with
'.csv as an extension

Dim db As Database
Dim rs As DAO.Recordset
Dim iFields As Integer
Dim strFieldNames As String
Dim i As Integer
Dim iFile As Integer
Dim strData As String
Dim strError As String

On Error GoTo ErrHandler

ExportCSV = False

'Open the query
Set db = CurrentDb
Set rs = db.OpenRecordset(strQueryName)
'Get the number of fields
iFields = rs.Fields.Count
i = 0
'Add field names
Do Until i = iFields
    strFieldNames = strFieldNames & _
        rs(i).Name & ","
    i = i + 1
Loop
'Remove last comma
strFieldNames = Left(strFieldNames, Len(strFieldNames) - 1)

'Delete existing file if present
On Error Resume Next
Kill strFileName
On Error GoTo ErrHandler
'Get file name and open it for writing
iFile = FreeFile
Open strFileName For Output As iFile

'Add field names to header
Print #iFile, strFieldNames

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        'Get data string for current record
        strData = ""
        i = 0
        Do Until i = iFields
            strData = strData & CStr(Nz(rs(i), "")) & ","
            i = i + 1
        Loop
        'Remove last comma
        strData = Left(strData, Len(strData) - 1)
        'Print data string to file and go to next record
        Print #iFile, strData
        rs.MoveNext
    Loop
End If

ExportCSV = True

'Open the file in notepad
If bStayOpen Then
    Close #iFile
    Shell "Notepad.exe " & strFileName, vbMaximizedFocus
End If

'Clean up
CSVExit:
    Close #iFile
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function

ErrHandler:
    'Create error message
    If ExportCSV Then
        strError = "Access encountered an error after creating the file."
    Else
        strError = "The output file " & strFileName & " could not be created."
    End If
    'Display error
    ErrorMsg strError, Err.Number, Err.Description, "Export Error"
    Resume CSVExit

End Function

Public Function ExportDetectors(strFileName As String, _
    strSettings As String, Optional bStayOpen As Boolean, _
    Optional bXL As Boolean) As Boolean
'Export trap files for SECR

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim qdf As QueryDef
Dim strTrapQuery As String
Dim strSQL As String
Dim strTrapFile As String

'Get sessions
ExportDetectors = False
Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT Session FROM " & _
    "qrySECRDetSubquery GROUP BY Session")

'Test for at least one session
If Not (rs.EOF And rs.BOF) Then
    'Get parameters
    strTrapQuery = "SECRDetectorQuery"
    strSQL = "SELECT LocationID, UTM_E, UTM_N, Effort, Location " & _
        "FROM qrySECRDetSubquery WHERE (((Session)=''))"
    Set qdf = db.QueryDefs(strTrapQuery)
    rs.MoveFirst
    'Loop through sessions
    Do Until rs.EOF
        qdf.SQL = Replace(strSQL, "''", "'" & rs!Session & "'")
        strTrapFile = Replace(strFileName, ".", _
            "Traps" & rs!Session & ".")
        'Create trap file or add sheet to workbook
        If bXL Then
            DoCmd.TransferSpreadsheet acExport, , _
                strTrapQuery, strFileName, , "Traps" & rs!Session
        Else
            ExportForSECR strTrapQuery, strSettings, _
                strTrapFile, bStayOpen
        End If
        rs.MoveNext
    Loop
    'Reset detector query
    strSQL = "SELECT LocationID, UTM_E, UTM_N, Effort, Location, " & _
        "Session FROM qrySECRDetSubquery"
    qdf.SQL = strSQL
End If

ExportDetectors = True

EDExit:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set qdf = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    ErrorMsg "An error occurred while createing trap files.", _
        Err.Number, Err.Description
    Resume EDExit

End Function

Public Function ExportForMARK(strQueryName As String, _
    strSettings As String, strFileName As String, _
    Optional bStayOpen As Boolean, _
    Optional bSkipLastField As Boolean) As Boolean
'Exports query results in MARK format
'Returns true if export is successful
'Field names and settings included in header
'First query field should be sample identifier and will be
'commented out in output
'Following fields should be capture history fields named
'Occasion1, Occasion2, etc
'Capture history fields are concatenated in output
'Count of 1 is inserted automatically for all capture histories
'Do not include a count field in the query
'Subsequent fields are class variables and covariates
'strSettings is added to header without modification
'strFileName should be the destination file name with
'.inp as an extension

Dim db As Database
Dim rs As DAO.Recordset
Dim iFields As Integer
Dim strHeaderLine As String
Dim i As Integer
Dim strTextFile As String
Dim iFile As Integer
Dim strData As String
Dim strOcc As String
Dim strError As String

On Error GoTo ErrHandler

ExportForMARK = False

'Open the query
Set db = CurrentDb
Set rs = db.OpenRecordset(strQueryName)
'Get the number of fields
iFields = rs.Fields.Count
If bSkipLastField Then
    iFields = iFields - 1
End If

'Get text file name and open it for writing
strTextFile = Replace(strFileName, ".inp", ".txt")
iFile = FreeFile
Open strTextFile For Output As iFile

'Add settings and field names to header
strHeaderLine = "/*" & strSettings & "*/"
i = InStr(1, strHeaderLine, "*//*")
Do Until i = 0
    Print #iFile, Left(strHeaderLine, i + 1)
    strHeaderLine = Mid(strHeaderLine, i + 2)
    i = InStr(1, strHeaderLine, "*//*")
Loop
Print #iFile, strHeaderLine

'Start the field name string for the header
strHeaderLine = FixedWidthString("/*" & rs(0).Name, 39) & " "
strHeaderLine = strHeaderLine & _
    FixedWidthString("CaptureHistory", 29) & " "
strHeaderLine = strHeaderLine & _
    FixedWidthString("Group", 29) & " "
i = 1
'Skip Occasion and Group fields
Do Until Left(rs(i).Name, 8) <> "Occasion"
    i = i + 1
Loop
Do Until i = iFields
    If Left(rs(i).Name, 5) = "Group" Then
        i = i + 1
    Else
        Exit Do
    End If
Loop
'Add remaining fields
Do Until i = iFields
    strHeaderLine = strHeaderLine & _
        FixedWidthString(rs(i).Name, 29) & " "
    i = i + 1
Loop
strHeaderLine = Trim(strHeaderLine) & "*/"
Print #iFile, strHeaderLine

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        'Get data string for current record
        'Start with sample identifier commented out
        strData = FixedWidthString("/*" & rs(0) & "*/", 39) & " "
        i = 1
        strOcc = ""
        'Add occasion fields without spaces for capture history
        Do Until Left(rs(i).Name, 8) <> "Occasion"
            strOcc = strOcc & rs(i)
            i = i + 1
        Loop
        strData = strData & FixedWidthString(strOcc, 29) & " "
        'Add group fields with spaces
        strOcc = ""
        Do Until i = iFields
            If Left(rs(i).Name, 5) = "Group" Then
                strOcc = strOcc & rs(i) & " "
                i = i + 1
            Else
                Exit Do
            End If
        Loop
        strData = strData & FixedWidthString(strOcc, 29) & " "
        'Add subsequent fields until the last one
        Do Until i = iFields
            strData = strData & FixedWidthString(rs(i), 29) & " "
            i = i + 1
        Loop
        'Remove trailing space and add semicolon
        strData = Trim(strData) & ";"
        'Print data string to file and go to next record
        Print #iFile, strData
        rs.MoveNext
    Loop
End If
Close #iFile

ExportForMARK = True

'Delete existing file if present
On Error Resume Next
Kill strFileName
On Error GoTo ErrHandler

'Change extension of text file to .inp
Name strTextFile As strFileName

'Open the file in notepad
If bStayOpen Then
    Shell "Notepad.exe " & strFileName, vbMaximizedFocus
End If

'Clean up
MARKExit:
    On Error Resume Next
    Close #iFile
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function

ErrHandler:
    If ExportForMARK Then
        strError = "Access encountered an error after creating the file."
    Else
        strError = "The output file " & strFileName & " could not be created."
    End If
    ErrorMsg strError, Err.Number, Err.Description, "Export Error"
    Reset
    Resume MARKExit

End Function

Public Function ExportForOverlap(strQueryName As String, _
    strSettings As String, strFileName As String, _
    Optional bStayOpen As Boolean) As Boolean
'Exports query results in Overlap format
'Returns true if export is successful
'Field names and settings included in header
'strSettings is added to header without modification
'strFileName should be the destination file name with
'.csv as an extension

Dim db As Database
Dim rs As DAO.Recordset
Dim iFields As Integer
Dim strFieldNames As String
Dim i As Integer
Dim iFile As Integer
Dim strData As String
Dim strError As String

On Error GoTo ErrHandler

ExportForOverlap = False

'Open the query
Set db = CurrentDb
Set rs = db.OpenRecordset(strQueryName)
'Get the number of fields
iFields = rs.Fields.Count
i = 0
'Add field names
Do Until i = iFields
    strFieldNames = strFieldNames & _
        rs(i).Name & ","
    i = i + 1
Loop
'Remove last comma
strFieldNames = Left(strFieldNames, Len(strFieldNames) - 1)

'Delete existing file if present
On Error Resume Next
Kill strFileName
On Error GoTo ErrHandler
'Get text file name and open it for writing
iFile = FreeFile
Open strFileName For Output As iFile

'Add settings and field names to header
Print #iFile, "#" & strSettings
Print #iFile, "#Make sure to identify comment character when importing:"
Print #iFile, "#read.csv(file, comment.char =  " & _
    chr(34) & "#" & chr(34) & ")"
Print #iFile, strFieldNames

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        'Get data string for current record
        strData = ""
        i = 0
        Do Until i = iFields
            strData = strData & CStr(rs(i)) & ","
            i = i + 1
        Loop
        'Remove last comma
        strData = Left(strData, Len(strData) - 1)
        'Print data string to file and go to next record
        Print #iFile, strData
        rs.MoveNext
    Loop
End If

ExportForOverlap = True

'Open the file in notepad
If bStayOpen Then
    Close #iFile
    Shell "Notepad.exe " & strFileName, vbMaximizedFocus
End If

'Clean up
OverlapExit:
    Close #iFile
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function

ErrHandler:
    'Create error message
    If ExportForOverlap Then
        strError = "Access encountered an error after creating the file."
    Else
        strError = "The output file " & strFileName & " could not be created."
    End If
    'Display error
    ErrorMsg strError, Err.Number, Err.Description, "Export Error"
    Resume OverlapExit

End Function

Public Function ExportForSECR(strQueryName As String, _
    strSettings As String, strFileName As String, _
    Optional bStayOpen As Boolean) As Boolean
'Exports query results in SECR format
'Returns true if export is successful
'Field names and settings included in header
'strSettings is added to header without modification
'strFileName should be the destination file name with
'.txt as an extension

Dim db As Database
Dim rs As DAO.Recordset
Dim iFields As Integer
Dim strFieldNames As String
Dim i As Integer
Dim iFile As Integer
Dim strData As String
Dim strError As String

On Error GoTo ErrHandler

ExportForSECR = False

'Open the query
Set db = CurrentDb
Set rs = db.OpenRecordset(strQueryName)
'Get the number of fields
iFields = rs.Fields.Count
'Start the field name string for the header
strFieldNames = FixedWidthString("#" & CStr(rs(0).Name), 19) & " "
i = 1
'Add remaining fields, except StartDateTime
Do Until i = iFields
    strFieldNames = strFieldNames & _
        FixedWidthString(CStr(rs(i).Name), 19) & " "
    i = i + 1
Loop

'Delete existing file if present
On Error Resume Next
Kill strFileName
On Error GoTo ErrHandler
'Get text file name and open it for writing
iFile = FreeFile
Open strFileName For Output As iFile

'Add settings and field names to header
Print #iFile, "#" & strSettings
Print #iFile, "#Make sure to set usage when importing:"
Print #iFile, "#read.capthist(captfile, trapfile, " & _
    "detector = " & chr(34) & "proximity" & chr(34) & _
    ", binary.usage = FALSE)"
Print #iFile, strFieldNames

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        'Get data string for current record
        strData = ""
        i = 0
        Do Until i = iFields
            strData = strData & FixedWidthString(CStr(Nz(rs(i), "")), 19) & " "
            i = i + 1
        Loop
        'Remove last space
        strData = Trim(strData)
        'Print data string to file and go to next record
        Print #iFile, strData
        rs.MoveNext
    Loop
End If

ExportForSECR = True

'Open the file in notepad
If bStayOpen Then
    Close #iFile
    Shell "Notepad.exe " & strFileName, vbMaximizedFocus
End If

'Clean up
SECRExit:
    Close #iFile
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Exit Function

ErrHandler:
    'Create error message
    If ExportForSECR Then
        strError = "Access encountered an error after creating the file."
    Else
        strError = "The output file " & strFileName & " could not be created."
    End If
    'Display error
    ErrorMsg strError, Err.Number, Err.Description, "Export Error"
    Close #iFile
    Resume SECRExit

End Function

Public Function ExportToExcel(strQueryName As String, _
    strSettingsIn As String, strFileName As String, _
    Optional bStayOpen As Boolean) As Boolean
'Exports query results to xlsx and inserts settings
'in a separate sheet within the workbook
'Returns true if export is successful
'Check filename for .xlsx extension and verify path before
'calling this procedure

Dim appXL As Object
Dim xlWorkbook As Object
Dim xlQuerySheet As Object
Dim xlSettingsSheet As Object
Dim strSettings As String
Dim i As Integer
Dim strBackupFile As String
Dim strError As String

'Delete existing results file
On Error Resume Next
Kill strFileName

On Error GoTo ErrHandler

ExportToExcel = False

'Create new results file
DoCmd.TransferSpreadsheet acExport, , strQueryName, strFileName
If strQueryName = "SECRQuery" Then
    ExportDetectors strFileName, "", , True
End If
'Start excel
Set appXL = CreateObject("Excel.Application")
'Open the results workbook
Set xlWorkbook = appXL.Workbooks.Open(strFileName, , False)
'Get the existing (query results) sheet
Set xlQuerySheet = xlWorkbook.Worksheets(strQueryName)
'Insert a new sheet after results and rename
Set xlSettingsSheet = xlWorkbook.Worksheets.Add(, xlQuerySheet)
xlSettingsSheet.Name = "Settings"
'Add settings string to first cell of new sheet, save, and close
xlSettingsSheet.Activate
strSettings = strSettingsIn
i = 0
Do Until InStr(1, strSettings, "*//*") = 0
    xlSettingsSheet.Range("A" & i + 1).Select
    appXL.ActiveCell = Left(strSettings, _
        InStr(1, strSettings, "*//*") - 1)
    strSettings = Mid(strSettings, _
        InStr(1, strSettings, "*//*") + 4)
    i = i + 1
Loop
xlSettingsSheet.Range("A" & i + 1).Select
appXL.ActiveCell = strSettings
xlQuerySheet.Activate
xlQuerySheet.Range("A1").Select
appXL.ActiveWorkbook.Save
ExportToExcel = True

If bStayOpen Then
    appXL.Visible = True
Else
    appXL.ActiveWorkbook.Close
    appXL.Quit
End If

strBackupFile = Left(strFileName, InStrRev(strFileName, "\")) & _
    "Backup of " & Mid(strFileName, InStrRev(strFileName, "\") + 1)
strBackupFile = Replace(strBackupFile, ".xlsx", ".xlk")
On Error Resume Next
Kill strBackupFile
On Error GoTo ErrHandler

ExcelExit:
    Set xlSettingsSheet = Nothing
    Set xlQuerySheet = Nothing
    Set xlWorkbook = Nothing
    Set appXL = Nothing
    Exit Function

ErrHandler:
    If ExportToExcel Then
        strError = "Access encountered an error after creating the file."
    Else
        strError = "The output file " & strFileName & " could not be created."
    End If
    ErrorMsg strError, Err.Number, Err.Description, "Export Error"
    Resume ExcelExit

End Function

Public Function ExtractExtension(strFile As String) As String
'Return file extension

ExtractExtension = Mid(strFile, InStrRev(strFile, ".") + 1)

End Function

Public Function ExtractFilename(strPath As String) As String
'Remove path from filename

Dim strResult As String

If InStr(1, strPath, "\") > 0 Then
    strResult = Mid(strPath, InStrRev(strPath, "\") + 1)
Else
    strResult = strPath
End If

ExtractFilename = strResult

End Function

Public Function FixedWidthString(strIn As String, iLength As Integer, _
    Optional bRightAlign As Boolean = False) As String
'Pads the left or right of a string to create columns in text output files

Dim strResult As String

'Check input length
If Len(strIn) > iLength Then
    'Too long for column - keep as is
    strResult = strIn
Else
    'Create a string of spaces
    strResult = Space(iLength)
    If bRightAlign Then
        'Concatenate input at end of string
        strResult = strResult & strIn
        'Take characters from right
        strResult = Right(strResult, iLength)
    Else
        'Concatenate input at beginning of string
        strResult = strIn & strResult
        'Take characters from left
        strResult = Left(strResult, iLength)
    End If
End If

FixedWidthString = strResult

End Function

Public Function FolderArray(MainFolder As String) As String()
'Return a string array of all subfolders within main folder

iFolder = 0
FolderArray = FolderArraySubRoutine(MainFolder)
Erase strSubfolders

End Function

Public Function FolderArraySubRoutine(MainFolder As String) As String()
'Populate subfolder array

Dim fso As Object
Dim pFolder As Object
Dim pSubfolders As Object
Dim pSubFolder As Object

Set fso = CreateObject("scripting.FileSystemObject")
Set pFolder = fso.GetFolder(MainFolder)
Set pSubfolders = pFolder.SubFolders
For Each pSubFolder In pSubfolders
    iFolder = iFolder + 1
    ReDim Preserve strSubfolders(1 To iFolder)
    strSubfolders(iFolder) = pSubFolder.Path
    Call FolderArraySubRoutine(pSubFolder.Path)
Next

Set fso = Nothing
Set pFolder = Nothing
Set pSubfolders = Nothing
Set pSubFolder = Nothing

FolderArraySubRoutine = strSubfolders

End Function

Public Function FolderSize(strFolder As String) As Double
'Returns the size of a folder in megabytes

On Error Resume Next

Dim fso As Object

Set fso = CreateObject("scripting.FileSystemObject")
FolderSize = fso.GetFolder(strFolder).Size / 1000000
Set fso = Nothing

End Function

Public Function GetDigitString(vInput As Variant) As String
'Return formatting string for the specified number of digits

Dim strResults As String

strResults = ""

If IsNumeric(vInput) Then
    strResults = Space(vInput)
    strResults = Replace(strResults, " ", "0")
End If

GetDigitString = strResults

End Function

Public Function GetImageDate(strFile As String) As Variant
'Reads "date taken" from metadata if possible

Dim img As Object
Dim strDate As String

On Error GoTo ErrHandler

'Start with date modified
GetImageDate = FileDateTime(strFile)

'Attempt to get date taken
Set img = CreateObject("WIA.ImageFile")
img.LoadFile strFile
If img.Properties.Exists("36867") Then
    strDate = Replace(img.Properties("36867").Value, ":", "/", 1, 2)
    GetImageDate = CDate(strDate)
End If

GIDExit:
    Set img = Nothing
    Exit Function

ErrHandler:
    Resume GIDExit

End Function

Public Function GetLongInt(InputStr As String) As Long
'Strips non-digits from a string & converts to a long integer

Dim i As Integer
Dim strChar As String
Dim lngAsc As Long
Dim strReturn As String

For i = 1 To Len(InputStr)

    strChar = Mid(InputStr, i, 1)
    lngAsc = Asc(strChar)
    If lngAsc >= 48 And lngAsc <= 57 Then
        strReturn = strReturn & strChar
    End If
Next

If strReturn = "" Then
    'No digits were found in the string, return 0
    GetLongInt = 0
Else
    GetLongInt = CLng(strReturn)
End If

End Function

Public Function GetNewFileName(strPrefix As String, lngFileNum As Long, _
    iDigits As Integer, strExt As String, Optional vDate As Variant, _
    Optional bIncludeTime As Boolean = False) As String
'Generate new file name for copying photos

Dim strNew As String

If Len(strPrefix) > 0 Then
    strNew = strPrefix & "_"
Else
    strNew = ""
End If

strNew = strNew & Format(lngFileNum, GetDigitString(iDigits))

If Not IsNull(vDate) Then
    strNew = strNew & "_" & Replace(Format(vDate, "short date"), "/", "-")
    If bIncludeTime Then
        strNew = strNew & "_" & Replace(Format(vDate, "short time"), ":", "-")
    End If
End If

strNew = strNew & "." & Replace(strExt, ".", "")

GetNewFileName = strNew

End Function

Public Function GetNextFilename(strFile As String) As String
'Adds (1), (2), etc to filenames to prevent overwrite

Dim strNewFile As String
Dim i As Integer

strNewFile = strFile
i = 1

Do Until FileExists(strNewFile) = False
    strNewFile = Replace(strFile, ".jpg", "(" & i & ").jpg")
    i = i + 1
Loop

GetNextFilename = strNewFile

End Function

Public Function GetNextImageNumber(strFolder As String, _
    strPrefix As String) As Long
'Finds the max image number in a folder by prefix
'and returns the next number

Dim i As Long
Dim j As Long
Dim strFile As String

i = 0
'Get first matching filename
strFile = Dir(AddSlash(strFolder) & strPrefix & "_*.jpg")
If Len(strFile) > 0 Then
    'At least one file exists
    Do Until strFile = ""
        'Remove prefix from filename
        strFile = Mid(strFile, Len(strPrefix) + 2)
        'Use j to store position of first non-numeric character
        j = 1
        Do Until Not IsNumeric(Mid(strFile, j, 1))
            j = j + 1
        Loop
        'Remove characters after image number
        strFile = Left(strFile, j - 1)
        'Use j to store image number from current file
        j = CLng(strFile)
        If j > i Then
            i = j
        End If
        'Next matching filename
        strFile = Dir()
    Loop
End If

GetNextImageNumber = i + 1

End Function

Public Function GetOccasion(lngLocation As Long, _
    dStart As Date, dEnd As Date, lngSpecies As Long, _
    bGroup As Boolean, Optional bVerified As Boolean = True) As Variant
'Returns ., 1, or 0 for a particular occasion, location, species

Dim strWhere As String

If CameraIsActive(lngLocation, dStart, dEnd) Then
    If bVerified Then
        strWhere = "StatusID=2 AND "
    End If
    If bGroup Then
        strWhere = strWhere & "GroupID=" & lngSpecies & " AND "
    Else
        strWhere = strWhere & "SpeciesID=" & lngSpecies & " AND "
    End If
    strWhere = strWhere & "LocationID=" & _
            lngLocation & " And ImageDate Between #" & _
            USDate(dStart) & "# And #" & USDate(dEnd) & "#"
    If DCount("*", "qryValidDetections", strWhere) > 0 Then
        GetOccasion = 1
    Else
        GetOccasion = 0
    End If
Else
    GetOccasion = "."
End If

End Function

Public Sub GetTableArray()
'Populates table name array for creating runtime DB

strLocalTables(0, 0) = "CurrentObserver"
strLocalTables(0, 1) = "COLocal"
strLocalTables(1, 0) = "Detections"
strLocalTables(1, 1) = "CPDLocal"
strLocalTables(2, 0) = "PhotoTags"
strLocalTables(2, 1) = "PTLocal"
strLocalTables(3, 0) = "Photos"
strLocalTables(3, 1) = "CPLocal"
strLocalTables(4, 0) = "Visits"
strLocalTables(4, 1) = "CVLocal"
strLocalTables(5, 0) = "CameraLocations"
strLocalTables(5, 1) = "CLLocal"
strLocalTables(6, 0) = "Observers"
strLocalTables(6, 1) = "ObsLocal"
strLocalTables(7, 0) = "ObsPhone"
strLocalTables(7, 1) = "ObsPLocal"
strLocalTables(8, 0) = "Species"
strLocalTables(8, 1) = "SpLocal"
strLocalTables(9, 0) = "SpeciesShortcuts"
strLocalTables(9, 1) = "SpShLocal"
strLocalTables(10, 0) = "StudyAreas"
strLocalTables(10, 1) = "SALocal"
strLocalTables(11, 0) = "Help"
strLocalTables(11, 1) = "HelpLocal"
strLocalTables(12, 0) = "lkupPhoneTypes"
strLocalTables(12, 1) = "lPTLocal"
strLocalTables(13, 0) = "DetectionDetails"
strLocalTables(13, 1) = "DDLocal"
strLocalTables(14, 0) = "DetailShortcuts"
strLocalTables(14, 1) = "DShLocal"

End Sub

Public Function HideSplash() As Integer
'Close the SplashScreen form

On Error Resume Next
DoCmd.Close acForm, "SplashScreen", acSaveNo
Application.Echo True
DoCmd.Hourglass False
DoEvents

HideSplash = 1

End Function

Public Sub HideSwitchboard(ObjType As Integer, strObjName As String, _
    Optional bShowRibbon As Boolean = False, _
    Optional bShowNav As Boolean = False)
'Minimizes the switchboard if it's open

If IsLoaded("Switchboard") Then
    DoCmd.SelectObject acForm, "Switchboard", False
    DoCmd.Minimize
    If bShowRibbon Then
        DoCmd.ShowToolbar "Ribbon", acToolbarYes
    End If
    If bShowNav Then
        DoCmd.SelectObject acTable, , True
    End If
    If ObjType = acForm Then
        If IsLoaded(strObjName) Then
            DoCmd.SelectObject ObjType, strObjName, False
            DoCmd.Restore
        End If
    End If
End If

End Sub

Public Function NoSlash(strFolder As String) As String
'Remove slash from end of folder path

If Right(strFolder, 1) = "\" Then
    NoSlash = Left(strFolder, Len(strFolder) - 1)
Else
    NoSlash = strFolder
End If

End Function

Public Function OpenPhotoWithDialog(vImgPath As Variant, _
    lngImageID As Long) As Boolean
'Open photo in default program or provide dialog if file is missing
'Returns true if location was changed, indicating refresh is needed

Dim bResult As Boolean
Dim dlg As Object
Dim strFilePath As String
Dim strSQL As String

On Error GoTo ErrHandler

bResult = False

'Test for file
If OpenPhoto(vImgPath) Then
    Exit Function
End If

'File must be missing, set dialog parameters
Set dlg = Application.FileDialog(msoFileDialogFilePicker)
strFilePath = ""
dlg.Filters.Add "Images", "*.gif; *.jpg; *.jpeg; *.bmp, " & _
    "*.tif; *.tiff; *.png"
dlg.AllowMultiSelect = False
dlg.Title = "Select Missing Photo"

If dlg.Show = -1 Then
    'File selected - update table
    strFilePath = dlg.SelectedItems(1)
    strSQL = "UPDATE Photos SET FilePath = '" & Left(strFilePath, _
        InStrRev(strFilePath, "\")) & "', FileName = '" & _
        Mid(strFilePath, InStrRev(strFilePath, "\") + 1) & _
        "' WHERE (((ImageID)=" & lngImageID & "))"
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
End If

bResult = True

OpenPhotoWithDialog = bResult

OPDExit:
    DoCmd.SetWarnings True
    Set dlg = Nothing
    Exit Function

ErrHandler:
    ErrorMsg "Could not update file location.", Err.Number, Err.Description
    Resume OPDExit

End Function

Public Function PhotosSize() As Double
'Returns the total disk space occupied by the photos
'selected in the photoviewer form in megabytes

Dim fso As Object
Dim db As Database
Dim rs As DAO.Recordset
Dim dblMB As Double

On Error GoTo ErrHandler

Set fso = CreateObject("scripting.FileSystemObject")
Set db = CurrentDb
Set rs = db.OpenRecordset("PhotoViewerQuery", dbOpenDynaset, dbSeeChanges)

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        dblMB = dblMB + fso.GetFile(rs!ImgPath).Size / 1000000
        rs.MoveNext
    Loop
End If

PhotosSize = dblMB

PSExit:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Set fso = Nothing
    Exit Function

ErrHandler:
    PhotosSize = 0
    Resume PSExit

End Function

Public Function QueryIsOpen(strQueryName As String) As Boolean
'Check if query is currently open in DB window

On Error GoTo ErrHandler

If CurrentData.AllQueries(strQueryName).CurrentView = 2 Then
    QueryIsOpen = True
End If
Exit Function

ErrHandler:
    QueryIsOpen = False

End Function

Public Function ReplaceYear(dDateIn As Date, iYear As Integer) As Date
'Returns the same month and day with a different year
'If date is Feb 29 and iYear is not a leap year, returns Feb 28
'of that year

Dim i As Integer

i = iYear - Year(dDateIn)
ReplaceYear = DateAdd("yyyy", i, dDateIn)

End Function

Public Function ResetTimer(frm As Form, Optional lngTimer As Long = 2000)

frm.TimerInterval = lngTimer
ResetTimer = 1

End Function

Public Sub RestoreSwitchboard(Optional bNavIsOpen As Boolean = False)
'Brings back the switchboard

If IsLoaded("Switchboard") Then
    If bNavIsOpen Then
        DoCmd.NavigateTo "Custom"
        DoCmd.RunCommand acCmdWindowHide
        DoEvents
    End If
    DoCmd.SelectObject acForm, "Switchboard", False
    DoCmd.Restore
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
End If

End Sub

Public Sub SafeSetFocus(ctlTarget As Control, _
    ctlAlwaysOn As Control)
'Set focus to a control without errors

If ctlTarget.Enabled And ctlTarget.Visible Then
    ctlTarget.SetFocus
Else
    ctlAlwaysOn.SetFocus
End If

End Sub

Public Sub SaveQuery(strQueryName As String, strSQL As String)
'Save or modify querydef

Dim db As DAO.Database
Dim qdf As DAO.QueryDef

Set db = CurrentDb

'Test for query
If QueryExists(strQueryName) Then
    'Close if open
    If QueryIsOpen(strQueryName) Then
        DoCmd.Close acQuery, strQueryName
    End If
    'Overwrite SQL
    Set qdf = db.QueryDefs(strQueryName)
    qdf.SQL = strSQL
Else
    'Create from scratch
    db.CreateQueryDef strQueryName, strSQL
End If

Set qdf = Nothing
Set db = Nothing

End Sub

Public Function SECRDetectorEffort(lngLocation As Long, _
    dStart As Date, dEnd As Date, iHoursBetween As Integer) As Double
'Calculates effort from 0 to 1 for a particular camera and occasion

Dim db As Database
Dim rs As DAO.Recordset
Dim dblEffort As Double
Dim dLastEndDate As Date

'Initialize
dblEffort = 0
dLastEndDate = CDate(0)
Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM Visits WHERE (((LocationID)=" & _
    lngLocation & ") AND ((VisitTypeID)<3) AND ((ActiveStart) Is Not " & _
    "Null) AND ((ActiveEnd) Is Not Null)) ORDER BY VisitDate")

'Check for records
If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    Do Until rs.EOF
        If (rs!ActiveStart > dEnd Or rs!ActiveEnd < dStart) Then
            'Visit is entirely outside occasion - do nothing
        Else
            If rs!ActiveStart <= dStart Then
                'Visit starts before occasion
                If rs!ActiveEnd >= dEnd Then
                    'Occasion totally within visit
                    dblEffort = 1
                    Exit Do
                Else
                    'Visit ends during occasion - calculate effort
                    dblEffort = dblEffort + ((rs!ActiveEnd - dStart) / _
                        (dEnd - dStart))
                    'Save end date to check next visit
                    dLastEndDate = rs!ActiveEnd
                End If
            Else
                'Visit starts during occasion - check last end date
                If rs!ActiveStart - dLastEndDate > _
                    iHoursBetween / 24 Then
                    'Not within time specified - use active start
                    dLastEndDate = rs!ActiveStart
                End If
                If rs!ActiveEnd >= dEnd Then
                    'Visit ends after occasion - calculate effort
                    dblEffort = dblEffort + ((dEnd - dLastEndDate) / _
                        (dEnd - dStart))
                Else
                    'Visit starts and ends during occasion - calculate effort
                    dblEffort = dblEffort + ((rs!ActiveEnd - dLastEndDate) / _
                        (dEnd - dStart))
                    dLastEndDate = rs!ActiveEnd
                End If
            End If
        End If
        rs.MoveNext
    Loop
End If

'Clean up
rs.Close
Set rs = Nothing
Set db = Nothing

SECRDetectorEffort = dblEffort

End Function

Public Function SECREffortString(lngLocation As Long, dStart As Date, _
    iDays As Integer, iOccasions As Integer, iHrs As Integer) As String
'Generates string of effort values ranging from 0 to 1 for SECR detector file

Dim i As Integer
Dim dOccStart As Date
Dim dOccEnd As Date
Dim strResult As String

'Initialize
i = 0
dOccStart = dStart
strResult = ""

Do Until i = iOccasions
    i = i + 1
    'Get occasion end date
    dOccEnd = DateAdd("d", iDays, dOccStart)
    'Calculate effort and add to string
    strResult = strResult & Round(SECRDetectorEffort(lngLocation, _
        dOccStart, dOccEnd, iHrs), 2) & " "
    dOccStart = dOccEnd
Loop

SECREffortString = Trim(strResult)

End Function

Public Function SECROccNumber(dStart As Date, iDays As Integer, _
    dDetDate As Date) As Integer
'Calculates occasion number for each individual detection

Dim dCheck As Date
Dim i As Integer

'Initialize
i = 0
dCheck = dStart

'Increment until end date falls after detection date
Do Until dCheck > dDetDate
    i = i + 1
    dCheck = DateAdd("d", iDays, dCheck)
Loop

SECROccNumber = i

End Function

Public Function SECRSplitEffortString(strIn As String, _
    iOcc As Integer) As Double
'Get detector effort for a particular occasion from effort string

Dim i As Integer
Dim j As Integer
Dim strEffort As String

i = 0
j = 0
'Find space before target value
Do Until i = iOcc - 1
    j = InStr(j + 1, strIn, " ")
    i = i + 1
Loop
'Remove preceding values
strEffort = Mid(strIn, j + 1)
'Find next space
j = InStr(1, strEffort, " ")
'Remove trailing values
If j > 0 Then
    strEffort = Left(strEffort, j - 1)
End If

SECRSplitEffortString = CDbl(strEffort)

End Function

Public Sub SpSelectTable()
'Creates a temporary copy of the species table for filtering
'forms to multiple species

Dim strSQL As String

On Error GoTo ErrHandler

'Create temporary species table with selection checkboxes
strSQL = "SELECT Species.SpeciesID, Species.CommonName, No AS [Select] " & _
    "INTO SpeciesTemp FROM Species;"
DoCmd.SetWarnings False
DoCmd.RunSQL strSQL

SSTExit:
    DoCmd.SetWarnings True
    Exit Sub

ErrHandler:
    ErrorMsg "An error occurred when creating the table 'SpeciesTemp.'", _
        Err.Number, Err.Description
    Resume SSTExit

End Sub

Public Function TimeBetween(TimeCheck As Variant, TimeStart As Date, _
    TimeEnd As Date) As Boolean
'Returns true if TimeCheck is between TimeStart and TimeEnd

Dim strTimeCheck As String
Dim TimeCheckTrim As Date

If IsNull(TimeCheck) Then
    TimeBetween = False
    Exit Function
End If

'Remove the date from TimeCheck
strTimeCheck = DatePart("h", TimeCheck) & ":" & _
    Format(DatePart("n", TimeCheck), "00")
TimeCheckTrim = CDate(strTimeCheck)

If TimeStart < TimeEnd Then
    If TimeCheckTrim < TimeStart Or TimeCheckTrim > TimeEnd Then
        TimeBetween = False
        Exit Function
    End If
Else
    If TimeCheckTrim < TimeStart And TimeCheckTrim > TimeEnd Then
        TimeBetween = False
        Exit Function
    End If
End If

TimeBetween = True

End Function

Public Sub TransferLocalTables(strTargetDB As String)
'Moves tables from CreateLocalTables function to a new database

Dim i As Integer

On Error GoTo ErrHandler

For i = 0 To 14
    DoCmd.TransferDatabase acExport, "Microsoft Access", _
        strTargetDB, acTable, strLocalTables(i, 1), _
        strLocalTables(i, 0), False
Next i

Exit Sub

ErrHandler:
    ErrorMsg "An error occurred transferring tables to the " & _
        "new database.", Err.Number, Err.Description

End Sub

Public Sub TransferModuleObjects(strTargetDB As String)
'Copies forms, reports, queries, and modules to new DB

'Transfer queries
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acQuery, _
    "PhotoIDRecSource", "PhotoIDRecSource"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acQuery, _
    "SpeciesRecSource", "SpeciesRecSource"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acQuery, _
    "DetailsRecSource", "DetailsRecSource"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acQuery, _
    "qryCurrentObserverName", "qryCurrentObserverName"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acQuery, _
    "qryBatchIDRecSource", "qryBatchIDRecSource"

'Transfer forms
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acForm, _
    "PhotoIDDSModule", "PhotoIDDetectionsSubform"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acForm, _
    "PhotoIDModule", "PhotoID"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acForm, _
    "EditShortcuts", "EditShortcuts"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acForm, _
    "Observers", "Observers"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acForm, _
    "ObsPhoneSubform", "ObsPhoneSubform"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acForm, _
    "PhotoIDLogin", "PhotoIDLogin"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acForm, _
    "Species", "Species"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acForm, _
    "SpeciesShortcutSubform", "SpeciesShortcutSubform"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acForm, _
    "DetectionDetails", "DetectionDetails"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acForm, _
    "DetailsShortcutSubform", "DetailsShortcutSubform"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acForm, _
    "CurrentObserverSubform", "CurrentObserverSubform"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acForm, _
    "BatchPhotoID", "BatchPhotoID"

'Transfer Reports
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acReport, _
    "Shortcuts", "Shortcuts"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acReport, _
    "SpShortcutList", "SpShortcutList"
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acReport, _
    "DetailShortcutList", "DetailShortcutList"

'Transfer modules
DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acModule, _
    "FunctionsForPhotoID", "FunctionsForPhotoID"

End Sub

Public Function USDate(dDateIn As Date) As String
'Return the date in US format for building SQL strings in VBA

USDate = Format(dDateIn, "m/d/yyyy hh:mm:ss")

End Function
