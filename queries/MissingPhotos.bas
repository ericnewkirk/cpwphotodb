Operation =1
Option =0
Where ="(((FileExists([FilePath] & [FileName]))=False))"
Begin InputTables
    Name ="Photos"
End
Begin OutputColumns
    Expression ="Photos.ImageID"
    Expression ="Photos.FilePath"
    Expression ="Photos.FileName"
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
dbText "Description" ="Lists directories and file names for jpeg files that can't be located"
Begin
    Begin
        dbText "Name" ="Photos.FilePath"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.FileName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.ImageID"
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
    Bottom =179
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Photos"
        Name =""
    End
End
