dbMemo "SQL" ="SELECT Visits.LocationID, Detections.SpeciesID, SetVisits.VisitDate AS SetVisitD"
    "ate, Photos.ImageDate, Photos.ImageNum, True AS Include, IIf(Count([Detections]."
    "[DetectionID])>1 Or Max([Detections].[StatusID])>1,2,1) AS Obs INTO Activity IN "
    "'C:\\Users\\NewkirkE\\Desktop\\CPWPhotoWarehouse_v4AP.accdb'\015\012FROM (Visits"
    " INNER JOIN Visits AS SetVisits ON Visits.SetVisitID = SetVisits.VisitID) INNER "
    "JOIN (Photos INNER JOIN Detections ON Photos.ImageID = Detections.ImageID) ON Vi"
    "sits.VisitID = Photos.VisitID\015\012WHERE (((Detections.SpeciesID)>0) AND ((Det"
    "ections.StatusID)<3))\015\012GROUP BY Visits.LocationID, Detections.SpeciesID, S"
    "etVisits.VisitDate, Photos.ImageDate, Photos.ImageNum;\015\012"
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Detections.SpeciesID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.ImageNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.ImageDate"
        dbInteger "ColumnWidth" ="2355"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.APZone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Include"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Obs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryCheckPullVisitsWithSetDate.SetVisitDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Detections.StatusID"
        dbLong "AggregateType" ="-1"
    End
End
