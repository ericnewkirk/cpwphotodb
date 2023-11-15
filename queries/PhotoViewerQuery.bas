Operation =1
Option =0
Where ="(((CameraLocations.LocationID) = 365))"
Begin InputTables
    Name ="CameraLocations"
    Name ="Visits"
    Name ="Visits"
    Alias ="SetVisits"
    Name ="Photos"
End
Begin OutputColumns
    Expression ="CameraLocations.LocationID"
    Expression ="CameraLocations.StudyAreaID"
    Expression ="CameraLocations.UTM_E"
    Expression ="CameraLocations.UTM_N"
    Expression ="CameraLocations.UTMZone"
    Alias ="FieldSeason"
    Expression ="Year([SetVisits].[VisitDate])"
    Expression ="Photos.FileName"
    Expression ="Visits.VisitID"
    Alias ="ImgID"
    Expression ="Photos.ImageID"
    Expression ="Photos.ImageNum"
    Expression ="Photos.ImageDate"
    Expression ="Photos.Highlight"
    Alias ="ImgPath"
    Expression ="[Photos].[FilePath] & [Photos].[FileName]"
End
Begin Joins
    LeftTable ="CameraLocations"
    RightTable ="Visits"
    Expression ="CameraLocations.LocationID = Visits.LocationID"
    Flag =1
    LeftTable ="Visits"
    RightTable ="SetVisits"
    Expression ="Visits.SetVisitID = SetVisits.VisitID"
    Flag =1
    LeftTable ="Visits"
    RightTable ="Photos"
    Expression ="Visits.VisitID = Photos.VisitID"
    Flag =1
End
Begin OrderBy
    Expression ="Photos.ImageDate"
    Flag =0
    Expression ="Photos.ImageID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbText "Description" ="Generated from the PhotoViewer form"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="FieldSeason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.FileName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ImgID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.ImageDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.Highlight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImgPath"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.StudyAreaID"
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
        dbText "Name" ="CameraLocations.UTMZone"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1465
    Bottom =852
    Left =-1
    Top =-1
    Right =1449
    Bottom =492
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
        Left =570
        Top =16
        Right =714
        Bottom =160
        Top =0
        Name ="Visits"
        Name =""
    End
    Begin
        Left =774
        Top =23
        Right =918
        Bottom =167
        Top =0
        Name ="Photos"
        Name =""
    End
End
