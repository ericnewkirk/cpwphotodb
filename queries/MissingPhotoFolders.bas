Operation =1
Option =0
Begin InputTables
    Name ="MissingPhotos"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="MissingPhotos.FilePath"
End
Begin Groups
    Expression ="MissingPhotos.FilePath"
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
dbText "Description" ="Lists directories for jpeg files that can't be located"
Begin
    Begin
        dbText "Name" ="MissingPhotos.FilePath"
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
    Bottom =213
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="MissingPhotos"
        Name =""
    End
End
