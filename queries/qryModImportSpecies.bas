dbMemo "SQL" ="INSERT INTO Species ( CommonName, Genus, Species, ShortName )\015\012SELECT Modu"
    "leSpecies.CommonName, ModuleSpecies.Genus, ModuleSpecies.Species, ModuleSpecies."
    "ShortName\015\012FROM ModuleSpecies LEFT JOIN Species ON ModuleSpecies.CommonNam"
    "e = Species.CommonName\015\012WHERE (((Species.SpeciesID) Is Null) AND ((ModuleS"
    "pecies.SpeciesID) In (SELECT SpeciesID FROM ModuleDetections GROUP BY SpeciesID)"
    "));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Species_1.CommonName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species_1.Genus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species_1.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species.SpeciesID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species1.CommonName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ModuleSpecies.CommonName"
        dbLong "AggregateType" ="-1"
    End
End
