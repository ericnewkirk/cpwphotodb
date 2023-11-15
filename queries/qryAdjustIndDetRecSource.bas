dbMemo "SQL" ="SELECT Photos.ImageID, Photos.FileName, [Photos].[FilePath] & [Photos].[FileName"
    "] AS ImgPath, qryIndDetRecSource.IndDetectionID, Photos.ImageDate, IIf([Photos]."
    "[ImageDate]=[IndependentDetections].[ModifiedStart], \"Current Start\", \"\") AS"
    " [Current]\015\012FROM (Visits INNER JOIN (Photos INNER JOIN (SELECT Detections."
    "ImageID, \011\011\011\011Detections.SpeciesID \011\011\011FROM Detections \011\011"
    "\011WHERE (((Detections.StatusID)<3)) \011\011\011GROUP BY Detections.ImageID, \011"
    "\011\011\011Detections.SpeciesID)  AS Q ON Photos.ImageID = Q.ImageID) ON Visits"
    ".VisitID = Photos.VisitID) INNER JOIN qryIndDetRecSource ON (Visits.LocationID ="
    " qryIndDetRecSource.LocationID) AND (Q.SpeciesID = qryIndDetRecSource.SpeciesID)"
    " AND (Photos.ImageDate > Nz(qryIndDetRecSource.PrevDetection, 0)) AND ((Photos.I"
    "mageDate < qryIndDetRecSource.NextDetection) OR (qryIndDetRecSource.NextDetectio"
    "n Is Null));\015\012"
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
    Begin
        dbText "Name" ="Photos.ImageDate"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
