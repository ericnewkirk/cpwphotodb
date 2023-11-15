Operation =1
Option =0
Begin InputTables
    Name ="Visits"
    Name ="Photos"
    Name ="Visits"
    Alias ="SetVisits"
End
Begin OutputColumns
    Alias ="FieldSeason"
    Expression ="Year([SetVisits].[VisitDate])"
    Expression ="Visits.LocationID"
    Expression ="Photos.ImageID"
    Expression ="Photos.FileName"
    Alias ="ImgPath"
    Expression ="[FilePath] & [FileName]"
    Expression ="Photos.Highlight"
    Expression ="Photos.ImageDate"
End
Begin Joins
    LeftTable ="Visits"
    RightTable ="SetVisits"
    Expression ="Visits.SetVisitID = SetVisits.VisitID"
    Flag =1
    LeftTable ="Visits"
    RightTable ="Photos"
    Expression ="Visits.VisitID = Photos.VisitID"
    Flag =1
End
Begin OrderBy
    Expression ="Photos.ImageID"
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
        dbText "Name" ="Visits.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.ImageID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.FileName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.Highlight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FieldSeason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImgPath"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.ImageDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1250
    Bottom =852
    Left =-1
    Top =-1
    Right =1234
    Bottom =358
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =295
        Top =19
        Right =439
        Bottom =163
        Top =0
        Name ="Visits"
        Name =""
    End
    Begin
        Left =472
        Top =11
        Right =616
        Bottom =155
        Top =0
        Name ="Photos"
        Name =""
    End
    Begin
        Left =483
        Top =167
        Right =627
        Bottom =311
        Top =0
        Name ="SetVisits"
        Name =""
    End
End
