dbMemo "SQL" ="SELECT StudyAreas.StudyAreaName, First(CameraLocations.LocationName) AS Location"
    "Name, CameraLocations.LocationID, First(Year([SetVisits].[VisitDate])) AS CamYea"
    "r, Visits.VisitID, First(Visits.VisitDate) AS VisitDate, Sum(qryPhotoIDSummarySu"
    "bquery2.Photos) AS TotalPhotos, Sum(qryPhotoIDSummarySubquery2.NoID) AS PhotosNo"
    "ID, Sum(qryPhotoIDSummarySubquery2.ID) AS PhotosID, Sum(qryPhotoIDSummarySubquer"
    "y2.VerifiedID) AS PhotosVerifiedID\015\012FROM (((StudyAreas INNER JOIN CameraLo"
    "cations ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID) INNER JOIN Visi"
    "ts AS SetVisits ON CameraLocations.LocationID = SetVisits.LocationID) INNER JOIN"
    " Visits ON SetVisits.VisitID = Visits.SetVisitID) INNER JOIN qryPhotoIDSummarySu"
    "bquery2 ON Visits.VisitID = qryPhotoIDSummarySubquery2.VisitID\015\012GROUP BY S"
    "tudyAreas.StudyAreaName, CameraLocations.LocationID, Visits.VisitID\015\012ORDER"
    " BY StudyAreas.StudyAreaName, First(CameraLocations.LocationName), First(Visits."
    "VisitDate);\015\012"
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
        dbText "Name" ="CamYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalPhotos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhotosNoID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhotosID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhotosVerifiedID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StudyAreas.StudyAreaName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VisitDate"
        dbLong "AggregateType" ="-1"
    End
End
