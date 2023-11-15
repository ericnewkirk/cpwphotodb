Operation =1
Option =0
Begin InputTables
    Name ="Visits"
End
Begin OutputColumns
    Expression ="Visits.LocationID"
    Alias ="Trap Nights"
    Expression ="Sum(Round([Visits].[ActiveEnd]-[Visits].[ActiveStart],3))"
End
Begin Groups
    Expression ="Visits.LocationID"
    GroupLevel =0
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
        dbText "Name" ="Trap Nights"
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
    Bottom =646
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Visits"
        Name =""
    End
End
