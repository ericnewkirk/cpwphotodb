Operation =1
Option =0
Having ="(((Count(PhotoViewerQuery.ImgID))>1))"
Begin InputTables
    Name ="PhotoViewerQuery"
End
Begin OutputColumns
    Expression ="PhotoViewerQuery.FileName"
    Alias ="CountOfImgID"
    Expression ="Count(PhotoViewerQuery.ImgID)"
End
Begin Groups
    Expression ="PhotoViewerQuery.FileName"
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
        dbText "Name" ="PhotoViewerQuery.FileName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfImgID"
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
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="PhotoViewerQuery"
        Name =""
    End
End
