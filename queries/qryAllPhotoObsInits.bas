Operation =1
Option =0
Begin InputTables
    Name ="Photos"
    Name ="Detections"
    Name ="Observers"
    Name ="Visits"
End
Begin OutputColumns
    Expression ="Visits.VisitID"
    Expression ="Observers.Initials"
End
Begin Joins
    LeftTable ="Photos"
    RightTable ="Detections"
    Expression ="Photos.ImageID = Detections.ImageID"
    Flag =1
    LeftTable ="Observers"
    RightTable ="Detections"
    Expression ="Observers.ObserverID = Detections.ObsID"
    Flag =1
    LeftTable ="Visits"
    RightTable ="Photos"
    Expression ="Visits.VisitID = Photos.VisitID"
    Flag =1
End
Begin Groups
    Expression ="Visits.VisitID"
    GroupLevel =0
    Expression ="Observers.Initials"
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
        dbText "Name" ="Observers.Initials"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.VisitID"
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
    Bottom =527
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =221
        Top =12
        Right =365
        Bottom =156
        Top =0
        Name ="Photos"
        Name =""
    End
    Begin
        Left =413
        Top =12
        Right =557
        Bottom =156
        Top =0
        Name ="Detections"
        Name =""
    End
    Begin
        Left =605
        Top =12
        Right =749
        Bottom =156
        Top =0
        Name ="Observers"
        Name =""
    End
    Begin
        Left =48
        Top =16
        Right =192
        Bottom =160
        Top =0
        Name ="Visits"
        Name =""
    End
End
