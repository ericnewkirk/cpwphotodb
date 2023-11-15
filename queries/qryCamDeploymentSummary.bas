Operation =1
Option =0
Where ="(((Visits.VisitTypeID)=3))"
Begin InputTables
    Name ="CameraLocations"
    Name ="Visits"
    Name ="SELECT * FROM Visits WHERE (((VisitTypeID) = 2))"
    Alias ="PullVisits"
End
Begin OutputColumns
    Expression ="CameraLocations.LocationID"
    Expression ="CameraLocations.LocationName"
    Expression ="CameraLocations.StudyAreaID"
    Alias ="CamYear"
    Expression ="Year([Visits].[VisitDate])"
    Alias ="SetDate"
    Expression ="Visits.VisitDate"
    Alias ="PullDate"
    Expression ="PullVisits.VisitDate"
    Alias ="DaysDeployed"
    Expression ="DateDiff(\"d\",[SetDate],[PullDate])"
    Expression ="CameraLocations.UTM_E"
    Expression ="CameraLocations.UTM_N"
    Alias ="TotalPhotos"
    Expression ="DSum(\"PhotoCount\",\"qryVisitsRecSource\",\"SetVisitID=\" & [Visits].[VisitID])"
    Expression ="PullVisits.Comments"
    Alias ="TrapNights"
    Expression ="Round(DSum(\"ActiveEnd - ActiveStart\",\"Visits\",\"SetVisitID=\" & [Visits].[Vi"
        "sitID]),3)"
End
Begin Joins
    LeftTable ="Visits"
    RightTable ="PullVisits"
    Expression ="Visits.VisitID = PullVisits.SetVisitID"
    Flag =2
    LeftTable ="CameraLocations"
    RightTable ="Visits"
    Expression ="CameraLocations.LocationID = Visits.LocationID"
    Flag =1
End
Begin OrderBy
    Expression ="CameraLocations.LocationID"
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
        dbText "Name" ="CameraLocations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.StudyAreaID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SetDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CamYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DaysDeployed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.UTM_E"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.UTM_N"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PullDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PullVisits.Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalPhotos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrapNights"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="CameraLocations.LocationName"
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
    Bottom =250
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="CameraLocations"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Visits"
        Name =""
    End
    Begin
        Left =647
        Top =12
        Right =791
        Bottom =156
        Top =0
        Name ="PullVisits"
        Name =""
    End
End
