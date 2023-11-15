dbMemo "SQL" ="UPDATE Detections INNER JOIN Detections AS D ON (Detections.ImageID = D.ImageID)"
    " AND (Nz(Detections.DetailID,0) = Nz(D.DetailID,0)) AND (Detections.Individuals "
    "= D.Individuals) AND (Detections.SpeciesID = D.SpeciesID) SET Detections.StatusI"
    "D = 3\015\012WHERE (((D.StatusID)=3) AND ((Detections.StatusID)=1));\015\012"
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
        dbText "Name" ="Detections_1.StatusID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Detections.StatusID"
        dbLong "AggregateType" ="-1"
    End
End
