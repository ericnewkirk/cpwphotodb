Operation =1
Option =0
Begin InputTables
    Name ="Species"
    Name ="DetectionDetails"
    Name ="DetailShortcuts"
End
Begin OutputColumns
    Expression ="DetectionDetails.DetailID"
    Expression ="DetectionDetails.DetailText"
    Expression ="DetectionDetails.SpeciesID"
    Expression ="DetailShortcuts.Shortcut"
End
Begin Joins
    LeftTable ="DetectionDetails"
    RightTable ="DetailShortcuts"
    Expression ="DetectionDetails.DetailID = DetailShortcuts.DetailID"
    Flag =2
    LeftTable ="Species"
    RightTable ="DetectionDetails"
    Expression ="Species.SpeciesID = DetectionDetails.SpeciesID"
    Flag =1
End
Begin OrderBy
    Expression ="Species.CommonName"
    Flag =0
    Expression ="DetailShortcuts.Shortcut"
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
        dbText "Name" ="DetectionDetails.DetailID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DetectionDetails.DetailText"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DetectionDetails.SpeciesID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DetailShortcuts.Shortcut"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1705
    Bottom =972
    Left =-1
    Top =-1
    Right =1689
    Bottom =187
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
        Name ="DetectionDetails"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="DetailShortcuts"
        Name =""
    End
End
