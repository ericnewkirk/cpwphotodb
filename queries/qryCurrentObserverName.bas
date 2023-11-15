Operation =1
Option =0
Begin InputTables
    Name ="CurrentObserver"
    Name ="Observers"
End
Begin OutputColumns
    Expression ="CurrentObserver.ObserverID"
    Alias ="COName"
    Expression ="[FirstName] & \" \" & [LastName]"
End
Begin Joins
    LeftTable ="Observers"
    RightTable ="CurrentObserver"
    Expression ="Observers.ObserverID = CurrentObserver.ObserverID"
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
        dbText "Name" ="COName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrentObserver.ObserverID"
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
    Bottom =563
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="CurrentObserver"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Observers"
        Name =""
    End
End
