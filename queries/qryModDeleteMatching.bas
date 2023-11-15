dbMemo "SQL" ="DELETE DISTINCTROW NewDetections.*\015\012FROM Detections INNER JOIN NewDetectio"
    "ns ON (Detections.ImageID = NewDetections.ImageID) AND (Detections.ObsID = NewDe"
    "tections.ObsID);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Detections.DetailID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewDetections.DetectionID"
        dbLong "AggregateType" ="-1"
    End
End
