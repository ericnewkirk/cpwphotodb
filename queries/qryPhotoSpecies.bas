Operation =1
Option =2
Where ="(((Detections.StatusID)<3))"
Begin InputTables
    Name ="Detections"
End
Begin OutputColumns
    Expression ="Detections.ImageID"
    Expression ="Detections.SpeciesID"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Detections.SpeciesID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Detections.ImageID"
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
    Bottom =556
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Detections"
        Name =""
    End
End
