dbMemo "SQL" ="UPDATE (Photos INNER JOIN Detections ON Photos.ImageID = Detections.ImageID) INN"
    "ER JOIN Detections AS D1 ON (Detections.ImageID = D1.ImageID) AND (Detections.In"
    "dividuals = D1.Individuals) AND (Detections.DetailID = D1.DetailID) AND (Detecti"
    "ons.SpeciesID = D1.SpeciesID) AND (Detections.DetectionID <> D1.DetectionID) SET"
    " Detections.StatusID = 2\015\012WHERE (((Photos.NeedsUpdate)=True) AND ((Detecti"
    "ons.StatusID)=1) AND ((D1.StatusID)<3));\015\012"
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
        dbText "Name" ="Detections.StatusID"
        dbLong "AggregateType" ="-1"
    End
End
