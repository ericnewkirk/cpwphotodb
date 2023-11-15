dbMemo "SQL" ="SELECT Photos.ImageID, [Photos].[FilePath] & [Photos].[FileName] AS ImgPath, Pho"
    "tos.FileName, Photos.ImageDate, qryIndDetRecSource.IndDetectionID\015\012FROM (V"
    "isits INNER JOIN (Photos INNER JOIN (SELECT Detections.ImageID, \011\011\011\011"
    "Detections.SpeciesID \011\011\011FROM Detections \011\011\011WHERE (((Detections"
    ".StatusID)<3) And ((Detections.SpeciesID)>0)) \011\011\011GROUP BY Detections.Im"
    "ageID, \011\011\011\011Detections.SpeciesID)  AS Q ON Photos.ImageID = Q.ImageID"
    ") ON Visits.VisitID = Photos.VisitID) INNER JOIN qryIndDetRecSource ON ((Photos."
    "ImageDate < qryIndDetRecSource.NextDetection) OR (qryIndDetRecSource.NextDetecti"
    "on Is Null)) AND (Photos.ImageDate >= qryIndDetRecSource.ModifiedStart) AND (Q.S"
    "peciesID = qryIndDetRecSource.SpeciesID) AND (Visits.LocationID = qryIndDetRecSo"
    "urce.LocationID);\015\012"
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
        dbText "Name" ="Photos.ImageID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImgPath"
        dbInteger "ColumnWidth" ="6210"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryIndDetRecSource.IndDetectionID"
        dbLong "AggregateType" ="-1"
    End
End
