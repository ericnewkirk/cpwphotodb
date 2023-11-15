Operation =1
Option =0
Begin InputTables
    Name ="Species"
    Name ="SpeciesShortcuts"
End
Begin OutputColumns
    Expression ="Species.SpeciesID"
    Expression ="Species.CommonName"
    Expression ="Species.Genus"
    Expression ="Species.Species"
    Expression ="SpeciesShortcuts.Shortcut"
    Expression ="Species.ShortName"
    Expression ="Species.GroupID"
End
Begin Joins
    LeftTable ="Species"
    RightTable ="SpeciesShortcuts"
    Expression ="Species.SpeciesID = SpeciesShortcuts.SpeciesID"
    Flag =2
End
Begin OrderBy
    Expression ="Species.CommonName"
    Flag =0
End
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
        dbText "Name" ="Species.SpeciesID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species.CommonName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species.Genus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpeciesShortcuts.Shortcut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species.ShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species.GroupID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1447
    Bottom =852
    Left =-1
    Top =-1
    Right =1431
    Bottom =204
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Species"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="SpeciesShortcuts"
        Name =""
    End
End
