dbMemo "SQL" ="SELECT PhotoIDRecSource.*\015\012FROM PhotoIDRecSource LEFT JOIN (SELECT Detecti"
    "ons.ImageID\015\012FROM Detections INNER JOIN CurrentObserver ON Detections.ObsI"
    "D = CurrentObserver.ObserverID \015\012UNION \015\012SELECT Detections.ImageID F"
    "ROM Detections WHERE (((StatusID)=2)))  AS HaveID ON PhotoIDRecSource.ImageID = "
    "HaveID.ImageID\015\012WHERE (((HaveID.ImageID) Is Null))\015\012ORDER BY PhotoID"
    "RecSource.ImageID;\015\012"
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
        dbText "Name" ="PhotoIDRecSource.FieldSeason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhotoIDRecSource.Visits.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhotoIDRecSource.Photos.ImageID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhotoIDRecSource.Photos.FileName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhotoIDRecSource.ImgPath"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhotoIDRecSource.Photos.Highlight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhotoIDRecSource.Photos.ImageDate"
        dbLong "AggregateType" ="-1"
    End
End
