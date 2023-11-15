Operation =1
Option =0
Begin InputTables
    Name ="qrySECRDetSubquery"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="LocationID"
    Alias ="Expr2"
    Expression ="UTM_E"
    Alias ="Expr3"
    Expression ="UTM_N"
    Alias ="Expr4"
    Expression ="Effort"
    Alias ="Expr5"
    Expression ="Location"
    Alias ="Expr6"
    Expression ="Session"
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
        dbText "Name" ="Effort"
        dbInteger "ColumnWidth" ="4095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
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
    Bottom =495
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qrySECRDetSubquery"
        Name =""
    End
End
