dbMemo "SQL" ="UPDATE (ModuleSpecies INNER JOIN NewDetections ON ModuleSpecies.SpeciesID = NewD"
    "etections.SpeciesID) INNER JOIN Species ON ModuleSpecies.CommonName = Species.Co"
    "mmonName SET NewDetections.SpeciesID = [Species].[SpeciesID];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Detections.ObsID"
        dbLong "AggregateType" ="-1"
    End
End
