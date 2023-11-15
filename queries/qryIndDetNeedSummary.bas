dbMemo "SQL" ="SELECT IndependentDetections.IndDetectionID, IndependentDetections.SpeciesID, In"
    "dependentDetections.LocationID, IndependentDetections.DefaultStart, IndependentD"
    "etections.ModifiedStart, IndependentDetections.Deleted, (SELECT Min(ModifiedStar"
    "t) \015\012\011\011FROM IndependentDetections AS ID2 \015\012\011\011WHERE (((ID"
    "2.LocationID)=IndependentDetections.LocationID) AND \015\012\011\011\011((ID2.Sp"
    "eciesID)=IndependentDetections.SpeciesID) AND \015\012\011\011\011((ID2.Modified"
    "Start)>IndependentDetections.ModifiedStart))) AS NextDetection, (SELECT Max(Modi"
    "fiedStart) \015\012\011\011FROM IndependentDetections AS ID2 \015\012\011\011WHE"
    "RE (((ID2.LocationID)=IndependentDetections.LocationID) AND \015\012\011\011\011"
    "((ID2.SpeciesID)=IndependentDetections.SpeciesID) AND \015\012\011\011\011((ID2."
    "ModifiedStart)<IndependentDetections.ModifiedStart) AND \015\012((ID2.Deleted)=F"
    "ALSE))) AS PrevDetection, (SELECT First(Year(VisitDate)) FROM Visits WHERE (((Vi"
    "sitID)=(SELECT First(SetVisitID) FROM Visits INNER JOIN Photos ON Visits.VisitID"
    " = Photos.VisitID WHERE (((Visits.LocationID)=IndependentDetections.LocationID) "
    "AND ((Photos.ImageDate)=IndependentDetections.DefaultStart)))))) AS FieldSeason\015"
    "\012FROM IndependentDetections LEFT JOIN IndDetGroups ON IndependentDetections.I"
    "ndDetectionID = IndDetGroups.IndDetectionID\015\012WHERE (((IndependentDetection"
    "s.Deleted)=False) AND ((IndDetGroups.IndDetGroupID) Is Null))\015\012ORDER BY In"
    "dependentDetections.LocationID, IndependentDetections.ModifiedStart;\015\012"
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
End
