dbMemo "SQL" ="SELECT StudyAreas.StudyAreaName, StudyAreas.StudyAreaID, CameraLocations.Locatio"
    "nName, CameraLocations.LocationID, CameraLocations.UTM_E, CameraLocations.UTM_N,"
    " CameraLocations.UTMZone, CameraLocations.LatitudeDD, CameraLocations.LongitudeD"
    "D, Year([SetVisits].[VisitDate]) AS FieldSeason, Photos.FileName, Visits.VisitID"
    ", Photos.ImageID AS ImgID, Photos.ImageNum, Photos.ImageDate, Photos.Highlight, "
    "[Photos].[FilePath] & [Photos].[FileName] AS ImgPath, qrySpPivotVerified.*\015\012"
    "FROM ((((StudyAreas INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = Camer"
    "aLocations.StudyAreaID) INNER JOIN Visits AS SetVisits ON CameraLocations.Locati"
    "onID = SetVisits.LocationID) INNER JOIN Visits ON SetVisits.VisitID = Visits.Set"
    "VisitID) INNER JOIN Photos ON Visits.VisitID = Photos.VisitID) LEFT JOIN qrySpPi"
    "votVerified ON Photos.ImageID = qrySpPivotVerified.ImageID\015\012ORDER BY Photo"
    "s.ImageID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="StudyAreas.StudyAreaName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.FileName"
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
        dbText "Name" ="qrySpeciesPivot.argali mountain sheep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.domestic horse"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.human"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.Mongolian gazelle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.None"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StudyAreas.StudyAreaID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.VisitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.Highlight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.ImageID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.corsac fox"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.Eurasian lynx"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.LongitudeDD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImgPath"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.birds"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.domestic cattle/goat/sheep/yak"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.gray wolf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.mountain hare"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.red deer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.Steppe polecat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.UTM_N"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.UTM_E"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.UTMZone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.LatitudeDD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FieldSeason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.ImageNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.badger"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.domestic dog"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.ground squirrel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.long-tailed ground squirrel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.Mongolian marmot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.no ID possible"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.Przewalski’s horse"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.red fox"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrySpeciesPivot.vehicle"
        dbLong "AggregateType" ="-1"
    End
End
