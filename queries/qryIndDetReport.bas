dbMemo "SQL" ="SELECT CameraLocations.StudyAreaID, CameraLocations.LocationID, (SELECT Min(SetV"
    "isits.VisitDate) FROM (Visits AS SetVisits INNER JOIN Visits ON SetVisits.VisitI"
    "D = Visits.SetVisitID) INNER JOIN Photos ON Visits.VisitID = Photos.VisitID WHER"
    "E (((SetVisits.LocationID) = [CameraLocations].[LocationID]) AND ((Photos.ImageD"
    "ate) = [qryIndDetRecSource].[ModifiedStart]))) AS FilterDate, qryIndDetRecSource"
    ".IndDetectionID, qryIndDetRecSource.SpeciesID, qryIndDetRecSource.ModifiedStart\015"
    "\012FROM CameraLocations INNER JOIN qryIndDetRecSource ON CameraLocations.Locati"
    "onID = qryIndDetRecSource.LocationID\015\012WHERE (((qryIndDetRecSource.Deleted)"
    "=False));\015\012"
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
Begin
    Begin
        dbText "Name" ="CameraLocations.StudyAreaID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryIndDetRecSource.ModifiedStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryIndDetRecSource.IndDetectionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryIndDetRecSource.SpeciesID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FilterDate"
        dbLong "AggregateType" ="-1"
    End
End
