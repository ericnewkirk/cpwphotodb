Operation =1
Option =0
Where ="(((Visits.VisitTypeID)=3))"
Begin InputTables
    Name ="StudyAreas"
    Name ="CameraLocations"
    Name ="Visits"
End
Begin OutputColumns
    Expression ="CameraLocations.LocationID"
    Expression ="CameraLocations.UTM_E"
    Expression ="CameraLocations.UTM_N"
    Alias ="Effort"
    Expression ="SECREffortString(CameraLocations.LocationID, qryCRSubquery.OccasionStart, 14, 4,"
        " 12)"
    Alias ="Location"
    Expression ="'#' & CameraLocations.LocationName"
    Alias ="Session"
    Expression ="Left(Replace([StudyAreas].[StudyAreaName],' ',''),12) & DCount('*','Visits','Vis"
        "itTypeID=3 And LocationID=' & [qryCRSubquery].[LocationID] & ' And VisitDate<=#'"
        " & [qryCRSubquery].[CameraSetDate] & '#')"
End
Begin Joins
    LeftTable ="StudyAreas"
    RightTable ="CameraLocations"
    Expression ="StudyAreas.StudyAreaID = CameraLocations.StudyAreaID"
    Flag =1
    LeftTable ="CameraLocations"
    RightTable ="Visits"
    Expression ="CameraLocations.LocationID = Visits.LocationID"
    Flag =1
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
        dbText "Name" ="CameraSetDates.LocationID"
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
        dbText "Name" ="Effort"
        dbInteger "ColumnWidth" ="4095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Session"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
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
    Bottom =393
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="CameraSetDates"
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
        Name ="StudyAreas"
        Name =""
    End
End
