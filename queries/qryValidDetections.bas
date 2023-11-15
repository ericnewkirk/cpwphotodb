Operation =1
Option =0
Where ="(((Detections.SpeciesID)>0) AND ((Detections.StatusID)<3))"
Begin InputTables
    Name ="Visits"
    Alias ="SetVisits"
    Name ="Visits"
    Name ="Photos"
    Name ="Detections"
    Name ="Species"
End
Begin OutputColumns
    Expression ="SetVisits.LocationID"
    Alias ="SetVisitDate"
    Expression ="SetVisits.VisitDate"
    Expression ="Photos.ImageID"
    Expression ="Photos.ImageDate"
    Expression ="Photos.VisitID"
    Expression ="Detections.SpeciesID"
    Expression ="Detections.StatusID"
    Expression ="Species.GroupID"
End
Begin Joins
    LeftTable ="SetVisits"
    RightTable ="Visits"
    Expression ="SetVisits.VisitID = Visits.SetVisitID"
    Flag =1
    LeftTable ="Visits"
    RightTable ="Photos"
    Expression ="Visits.VisitID = Photos.VisitID"
    Flag =1
    LeftTable ="Photos"
    RightTable ="Detections"
    Expression ="Photos.ImageID = Detections.ImageID"
    Flag =1
    LeftTable ="Species"
    RightTable ="Detections"
    Expression ="Species.SpeciesID = Detections.SpeciesID"
    Flag =1
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
        dbText "Name" ="Photos.ImageID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.ImageDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Photos.VisitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Detections.SpeciesID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Detections.StatusID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SetVisits.LocationID"
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
    Right =1418
    Bottom =891
    Left =-1
    Top =-1
    Right =1402
    Bottom =420
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="SetVisits"
        Name =""
    End
    Begin
        Left =239
        Top =39
        Right =383
        Bottom =183
        Top =0
        Name ="Visits"
        Name =""
    End
    Begin
        Left =422
        Top =37
        Right =566
        Bottom =181
        Top =0
        Name ="Photos"
        Name =""
    End
    Begin
        Left =593
        Top =34
        Right =737
        Bottom =178
        Top =0
        Name ="Detections"
        Name =""
    End
    Begin
        Left =785
        Top =12
        Right =929
        Bottom =156
        Top =0
        Name ="Species"
        Name =""
    End
End
