dbMemo "SQL" ="SELECT CameraLocations.LocationID, CameraLocations.StudyAreaID, CameraLocations."
    "UTM_E, CameraLocations.UTM_N, CameraLocations.UTMZone, Year([SetVisits].[VisitDa"
    "te]) AS FieldSeason, Photos.FileName, Visits.VisitID, Photos.ImageID AS ImgID, P"
    "hotos.ImageNum, Photos.ImageDate, Photos.Highlight, [Photos].[FilePath] & [Photo"
    "s].[FileName] AS ImgPath\015\012FROM ((CameraLocations INNER JOIN Visits ON Came"
    "raLocations.LocationID = Visits.LocationID) INNER JOIN Visits AS SetVisits ON Vi"
    "sits.SetVisitID = SetVisits.VisitID) INNER JOIN Photos ON Visits.VisitID = Photo"
    "s.VisitID\015\012ORDER BY Photos.ImageDate, Photos.ImageID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbText "Description" ="Record source for the CamFilmStrip subform"
Begin
    Begin
        dbText "Name" ="ImgPath"
        dbInteger "ColumnWidth" ="7800"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImgID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.ImageDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.Highlight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FieldSeason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.FileName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.ImageNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.VisitID"
        dbLong "AggregateType" ="-1"
    End
End
