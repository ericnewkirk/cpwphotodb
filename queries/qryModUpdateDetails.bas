dbMemo "SQL" ="UPDATE (Species INNER JOIN ((ModuleDetails INNER JOIN NewDetections ON ModuleDet"
    "ails.DetailID = NewDetections.DetailID) INNER JOIN DetectionDetails ON ModuleDet"
    "ails.DetailText = DetectionDetails.DetailText) ON Species.SpeciesID = DetectionD"
    "etails.SpeciesID) INNER JOIN ModuleSpecies ON (Species.CommonName = ModuleSpecie"
    "s.CommonName) AND (ModuleDetails.SpeciesID = ModuleSpecies.SpeciesID) SET NewDet"
    "ections.DetailID = [DetectionDetails].[DetailID];\015\012"
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
        dbText "Name" ="NewDetections.DetailID"
        dbLong "AggregateType" ="-1"
    End
End
