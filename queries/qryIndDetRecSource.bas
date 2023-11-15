dbMemo "SQL" ="SELECT IndependentDetections.IndDetectionID, IndependentDetections.SpeciesID, In"
    "dependentDetections.LocationID, IndependentDetections.DefaultStart, IndependentD"
    "etections.ModifiedStart, IndependentDetections.Deleted, (SELECT Min(ModifiedStar"
    "t) \015\012\011\011FROM IndependentDetections AS ID2 \015\012\011\011WHERE (((ID"
    "2.LocationID)=IndependentDetections.LocationID) AND \015\012\011\011\011((ID2.Sp"
    "eciesID)=IndependentDetections.SpeciesID) AND \015\012\011\011\011((ID2.Modified"
    "Start)>IndependentDetections.ModifiedStart) AND \015\012((ID2.Deleted)=FALSE))) "
    "AS NextDetection, (SELECT Max(ModifiedStart) \015\012\011\011FROM IndependentDet"
    "ections AS ID3 \015\012\011\011WHERE (((ID3.LocationID)=IndependentDetections.Lo"
    "cationID) AND \015\012\011\011\011((ID3.SpeciesID)=IndependentDetections.Species"
    "ID) AND \015\012\011\011\011((ID3.ModifiedStart)<IndependentDetections.ModifiedS"
    "tart) AND \015\012((ID3.Deleted)=FALSE))) AS PrevDetection, (SELECT First(Year(V"
    "isitDate)) FROM Visits WHERE (((VisitID)=(SELECT First(SetVisitID) FROM Visits I"
    "NNER JOIN Photos ON Visits.VisitID = Photos.VisitID WHERE (((Visits.LocationID)="
    "IndependentDetections.LocationID) AND ((Photos.ImageDate)=IndependentDetections."
    "DefaultStart)))))) AS FieldSeason\015\012FROM IndependentDetections\015\012WHERE"
    " (((IndependentDetections.Deleted)=False))\015\012ORDER BY IndependentDetections"
    ".LocationID, IndependentDetections.ModifiedStart;\015\012"
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
        dbText "Name" ="IndependentDetections.IndDetectionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IndependentDetections.SpeciesID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IndependentDetections.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IndependentDetections.DefaultStart"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IndependentDetections.ModifiedStart"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IndependentDetections.Deleted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NextDetection"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrevDetection"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FieldSeason"
        dbLong "AggregateType" ="-1"
    End
End
