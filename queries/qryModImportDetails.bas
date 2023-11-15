dbMemo "SQL" ="INSERT INTO DetectionDetails ( DetailText, SpeciesID )\015\012SELECT MSD.DetailT"
    "ext, MSD.SpeciesID\015\012FROM (SELECT ModuleDetails.DetailText, Species.Species"
    "ID FROM (ModuleDetails INNER JOIN ModuleSpecies ON ModuleDetails.SpeciesID = Mod"
    "uleSpecies.SpeciesID) INNER JOIN Species ON ModuleSpecies.CommonName = Species.C"
    "ommonName WHERE (((ModuleDetails.DetailText) Is Not Null) AND ((ModuleDetails.De"
    "tailID) In (SELECT DetailID FROM ModuleDetections GROUP BY DetailID))))  AS MSD "
    "LEFT JOIN DetectionDetails ON (MSD.SpeciesID = DetectionDetails.SpeciesID) AND ("
    "MSD.DetailText = DetectionDetails.DetailText)\015\012WHERE (((DetectionDetails.D"
    "etailID) Is Null));\015\012"
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
    Begin
        dbText "Name" ="ModuleDetails.DetailText"
        dbLong "AggregateType" ="-1"
    End
End
