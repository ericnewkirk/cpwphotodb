Operation =1
Option =0
Begin InputTables
    Name ="Photos"
    Name ="Detections"
End
Begin OutputColumns
    Expression ="Photos.VisitID"
    Expression ="Photos.ImageID"
    Alias ="NoIDs"
    Expression ="Sum(IIf([Detections].[DetectionID] Is Null,1,0))"
    Alias ="PendingIDs"
    Expression ="Sum(IIf([Detections].[StatusID]=1,1,0))"
    Alias ="VerifiedIDs"
    Expression ="Sum(IIf([Detections].[StatusID]=2,1,0))"
End
Begin Joins
    LeftTable ="Photos"
    RightTable ="Detections"
    Expression ="Photos.ImageID = Detections.ImageID"
    Flag =2
End
Begin Groups
    Expression ="Photos.VisitID"
    GroupLevel =0
    Expression ="Photos.ImageID"
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
        dbText "Name" ="Photos.VisitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.ImageID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NoIDs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PendingIDs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VerifiedIDs"
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
    Bottom =522
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Photos"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Detections"
        Name =""
    End
End
