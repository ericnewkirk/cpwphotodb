Operation =1
Option =0
Begin InputTables
    Name ="CameraLocations"
    Name ="Visits"
    Name ="Photos"
    Name ="Visits"
    Alias ="SetVisits"
End
Begin OutputColumns
    Expression ="CameraLocations.StudyAreaID"
    Expression ="CameraLocations.LocationID"
    Alias ="FieldSeason"
    Expression ="Year([SetVisits].[VisitDate])"
    Expression ="Visits.VisitID"
    Alias ="ImgID"
    Expression ="Photos.ImageID"
    Expression ="Photos.ImageDate"
    Expression ="Photos.Highlight"
    Expression ="Photos.ObsCount"
    Expression ="Photos.Verified"
    Expression ="Photos.Pending"
    Expression ="Photos.MultiSp"
    Expression ="Photos.NotNone"
End
Begin Joins
    LeftTable ="CameraLocations"
    RightTable ="Visits"
    Expression ="CameraLocations.LocationID = Visits.LocationID"
    Flag =1
    LeftTable ="Visits"
    RightTable ="Photos"
    Expression ="Visits.VisitID = Photos.VisitID"
    Flag =1
    LeftTable ="Visits"
    RightTable ="SetVisits"
    Expression ="Visits.SetVisitID = SetVisits.VisitID"
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
        dbText "Name" ="Photos.ImageDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.Highlight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImgID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FieldSeason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.NotNone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.StudyAreaID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.VisitID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1459
    Bottom =852
    Left =-1
    Top =-1
    Right =1443
    Bottom =421
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =207
        Top =63
        Right =351
        Bottom =207
        Top =0
        Name ="CameraLocations"
        Name =""
    End
    Begin
        Left =433
        Top =110
        Right =577
        Bottom =254
        Top =0
        Name ="Visits"
        Name =""
    End
    Begin
        Left =695
        Top =38
        Right =839
        Bottom =182
        Top =0
        Name ="Photos"
        Name =""
    End
    Begin
        Left =695
        Top =208
        Right =839
        Bottom =352
        Top =0
        Name ="SetVisits"
        Name =""
    End
End
