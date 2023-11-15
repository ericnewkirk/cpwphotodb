Operation =1
Option =0
Begin InputTables
    Name ="qryCamStudyAreaSpSummary"
End
Begin OutputColumns
    Expression ="qryCamStudyAreaSpSummary.CameraYear"
    Expression ="qryCamStudyAreaSpSummary.CommonName"
    Alias ="Images"
    Expression ="Sum(qryCamStudyAreaSpSummary.Images)"
End
Begin Groups
    Expression ="qryCamStudyAreaSpSummary.CameraYear"
    GroupLevel =0
    Expression ="qryCamStudyAreaSpSummary.CommonName"
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
        dbText "Name" ="qryCamStudyAreaSpSummary.CameraYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Images"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryCamStudyAreaSpSummary.CommonName"
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
    Bottom =512
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qryCamStudyAreaSpSummary"
        Name =""
    End
End
