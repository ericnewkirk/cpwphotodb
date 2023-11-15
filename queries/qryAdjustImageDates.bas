Operation =1
Option =0
Begin InputTables
    Name ="Visits"
End
Begin OutputColumns
    Expression ="Visits.VisitID"
    Alias ="FirstPhotoDateTime"
    Expression ="(SELECT Min(ImageDate) FROM Photos WHERE ((([Photos].[VisitID])=[Visits].[VisitI"
        "D])))"
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
        dbText "Name" ="FirstPhotoDateTime"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1705
    Bottom =852
    Left =-1
    Top =-1
    Right =1449
    Bottom =607
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
