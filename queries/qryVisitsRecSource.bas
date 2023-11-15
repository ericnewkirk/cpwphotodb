﻿Operation =1
Option =0
Begin InputTables
    Name ="Visits"
End
Begin OutputColumns
    Expression ="Visits.VisitID"
    Expression ="Visits.LocationID"
    Expression ="Visits.VisitTypeID"
    Expression ="Visits.VisitDate"
    Expression ="Visits.Comments"
    Alias ="PhotoCount"
    Expression ="(SELECT Count(ImageID) FROM Photos WHERE ((([Photos].[VisitID]) = [Visits].[Visi"
        "tID])))"
    Expression ="Visits.SetVisitID"
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
        dbText "Name" ="Visits.VisitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.VisitTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.VisitDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PhotoCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.SetVisitID"
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
    Bottom =590
    Left =0
    Top =0
    ColumnsShown =539
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
