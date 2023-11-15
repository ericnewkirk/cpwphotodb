Operation =1
Option =0
Begin InputTables
    Name ="PhotoIDRecSource"
End
Begin OutputColumns
    Expression ="PhotoIDRecSource.ImageID"
    Expression ="PhotoIDRecSource.FileName"
    Alias ="Current"
    Expression ="IIf(PhotoIDRecSource.ImageID=[Forms]![PhotoID]![ImageID],'Current Photo', '')"
    Expression ="PhotoIDRecSource.ImgPath"
End
Begin OrderBy
    Expression ="PhotoIDRecSource.ImageID"
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
        dbText "Name" ="PhotoIDRecSource.ImageID"
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
    Bottom =362
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="PhotoIDRecSource"
        Name =""
    End
End
