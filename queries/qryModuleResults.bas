dbMemo "SQL" ="SELECT Detections.DetectionID, Detections.StatusID\015\012FROM Detections INNER "
    "JOIN NewDetections ON Detections.DetectionID = NewDetections.DetectionID;\015\012"
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
        dbText "Name" ="Detections.DetectionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Detections.StatusID"
        dbLong "AggregateType" ="-1"
    End
End
