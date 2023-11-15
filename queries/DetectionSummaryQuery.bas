dbMemo "SQL" ="SELECT StudyAreas.StudyAreaName AS StudyArea, CameraLocations.LocationName AS Lo"
    "cation, Species.CommonName AS Species, qSub.StartDateTime, qSub.TotalIndividuals"
    ", qSub.TotalAdults, qSub.TotalJuveniles, qSub.TotalSubadults, qSub.TotalFemales,"
    " qSub.TotalMales, qSub.TotalKnownInd\015\012FROM StudyAreas INNER JOIN (CameraLo"
    "cations INNER JOIN (Species INNER JOIN (SELECT IndependentDetections.SpeciesID, "
    "IndependentDetections.LocationID, IndependentDetections.ModifiedStart As StartDa"
    "teTime, qGroup.TotalIndividuals, qGroup.TotalAdults, qGroup.TotalJuveniles, qGro"
    "up.TotalSubadults, qGroup.TotalMales, qGroup.TotalMales, qGroup.TotalFemales, qK"
    "nown.KnownIndividuals As TotalKnownInd FROM (IndependentDetections INNER JOIN (S"
    "ELECT IndependentDetections.IndDetectionID, Nz(Sum(IndDetGroups.Individuals),0) "
    "AS TotalIndividuals, Sum(IIf([AgeClassID]=1,[Individuals],0)) AS TotalAdults, Su"
    "m(IIf([AgeClassID]=2,[Individuals],0)) AS TotalJuveniles, Sum(IIf([AgeClassID]=3"
    ",[Individuals],0)) AS TotalSubadults, Sum(IIf([GenderID]=1,[Individuals],0)) AS "
    "TotalFemales, Sum(IIf([GenderID]=2,[Individuals],0)) AS TotalMales FROM Independ"
    "entDetections LEFT JOIN IndDetGroups ON IndependentDetections.IndDetectionID = I"
    "ndDetGroups.IndDetectionID GROUP BY IndependentDetections.IndDetectionID) As qGr"
    "oup ON IndependentDetections.IndDetectionID = qGroup.IndDetectionID) INNER JOIN "
    "(SELECT IndependentDetections.IndDetectionID, Sum(IIf(IndividualDetections.Indiv"
    "idualID>0,1,0)) AS KnownIndividuals FROM IndependentDetections LEFT JOIN Individ"
    "ualDetections ON IndependentDetections.IndDetectionID = IndividualDetections.Ind"
    "DetectionID GROUP BY IndependentDetections.IndDetectionID) As qKnown  ON Indepen"
    "dentDetections.IndDetectionID = qKnown.IndDetectionID WHERE (((IndependentDetect"
    "ions.Deleted)=False)))  AS qSub ON Species.SpeciesID = qSub.SpeciesID) ON Camera"
    "Locations.LocationID = qSub.LocationID) ON StudyAreas.StudyAreaID = CameraLocati"
    "ons.StudyAreaID\015\012ORDER BY Species.CommonName, StudyAreas.StudyAreaName, Ca"
    "meraLocations.LocationName, qSub.StartDateTime;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="StudyArea"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSub.TotalIndividuals"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSub.TotalAdults"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSub.TotalJuveniles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSub.TotalSubadults"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSub.TotalFemales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSub.TotalMales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSub.TotalKnownInd"
        dbLong "AggregateType" ="-1"
    End
End
