Operation =1
Option =0
Begin InputTables
    Name ="StudyAreas"
    Name ="CameraLocations"
    Name ="Visits"
    Name ="lkupVisitTypes"
End
Begin OutputColumns
    Expression ="Visits.VisitID"
    Expression ="StudyAreas.StudyAreaName"
    Expression ="CameraLocations.LocationName"
    Expression ="lkupVisitTypes.VisitType"
    Expression ="Visits.VisitDate"
    Expression ="Visits.ActiveStart"
    Expression ="Visits.ActiveEnd"
    Alias ="FirstPhoto"
    Expression ="DMin(\"ImageDate\",\"Photos\",\"VisitID=\" & [VisitID])"
    Alias ="LastPhoto"
    Expression ="DMax(\"ImageDate\",\"Photos\",\"VisitID=\" & [VisitID])"
    Expression ="Visits.Comments"
End
Begin Joins
    LeftTable ="lkupVisitTypes"
    RightTable ="Visits"
    Expression ="lkupVisitTypes.ID = Visits.VisitTypeID"
    Flag =1
    LeftTable ="StudyAreas"
    RightTable ="CameraLocations"
    Expression ="StudyAreas.StudyAreaID = CameraLocations.StudyAreaID"
    Flag =1
    LeftTable ="CameraLocations"
    RightTable ="Visits"
    Expression ="CameraLocations.LocationID = Visits.LocationID"
    Flag =1
End
Begin OrderBy
    Expression ="StudyAreas.StudyAreaName"
    Flag =0
    Expression ="CameraLocations.LocationName"
    Flag =0
    Expression ="Visits.VisitDate"
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
        dbText "Name" ="StudyAreas.StudyAreaName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lkupVisitTypes.VisitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.VisitDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.ActiveStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.ActiveEnd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstPhoto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastPhoto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Visits.Comments"
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
    Right =1705
    Bottom =972
    Left =-1
    Top =-1
    Right =1689
    Bottom =383
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="StudyAreas"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="CameraLocations"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Visits"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="lkupVisitTypes"
        Name =""
    End
End
