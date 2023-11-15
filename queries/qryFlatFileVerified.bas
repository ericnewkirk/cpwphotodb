dbMemo "SQL" ="SELECT StudyAreas.StudyAreaName, StudyAreas.StudyAreaID, CameraLocations.Locatio"
    "nName, CameraLocations.LocationID, CameraLocations.UTM_E, CameraLocations.UTM_N,"
    " CameraLocations.UTMZone, CameraLocations.LatitudeDD, CameraLocations.LongitudeD"
    "D, Year([SetVisits].[VisitDate]) AS FieldSeason, Photos.FileName, Visits.VisitID"
    ", Photos.ImageID AS ImgID, Photos.ImageNum, Photos.ImageDate, Photos.Highlight, "
    "[Photos].[FilePath] & [Photos].[FileName] AS ImgPath, DetQuery.SpeciesID, DetQue"
    "ry.CommonName, DetQuery.DetailText, DetQuery.Individuals\015\012FROM (StudyAreas"
    " INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyAre"
    "aID) INNER JOIN (((Visits AS SetVisits INNER JOIN Visits ON SetVisits.VisitID = "
    "Visits.SetVisitID) INNER JOIN Photos ON Visits.VisitID = Photos.VisitID) LEFT JO"
    "IN (SELECT Q.ImageID, \015\012          Species.SpeciesID, \015\012          Spe"
    "cies.CommonName, \015\012          DetectionDetails.DetailText, \015\012        "
    "  Q.MaxInd As Individuals \015\012        FROM ((SELECT Detections.ImageID, \015"
    "\012                      Detections.SpeciesID, \015\012                      De"
    "tections.DetailID, \015\012                      Max(IIf([Detections].[SpeciesID"
    "]>0,[Detections].[Individuals],0)) AS MaxInd\015\012                    FROM Det"
    "ections\015\012                    WHERE (((Detections.StatusID)=2))\015\012    "
    "                GROUP BY Detections.ImageID, Detections.SpeciesID, Detections.De"
    "tailID) As Q INNER JOIN \015\012          Species ON \015\012            Q.Speci"
    "esID = Species.SpeciesID) LEFT JOIN \015\012          DetectionDetails ON \015\012"
    "            Q.DetailID = DetectionDetails.DetailID)  AS DetQuery ON Photos.Image"
    "ID = DetQuery.ImageID) ON CameraLocations.LocationID = SetVisits.LocationID\015\012"
    "ORDER BY Photos.ImageID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbText "Description" ="Generated from the PhotoViewer form"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="FieldSeason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.FileName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="ImgPath"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.UTM_E"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.UTM_N"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.UTMZone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StudyAreas.StudyAreaName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.LongitudeDD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StudyAreas.StudyAreaID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.LatitudeDD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.VisitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DetQuery.CommonName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.ImageNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DetQuery.Individuals"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DetQuery.DetailText"
        dbLong "AggregateType" ="-1"
    End
End
