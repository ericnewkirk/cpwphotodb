Operation =1
Option =0
Begin InputTables
    Name ="StudyAreas"
    Name ="CameraLocations"
End
Begin OutputColumns
    Expression ="CameraLocations.LocationID"
    Alias ="LocName"
    Expression ="[StudyAreas].[StudyAreaAbbr] & \" - \" & [CameraLocations].[LocationName]"
    Expression ="CameraLocations.AccessNotes"
End
Begin Joins
    LeftTable ="StudyAreas"
    RightTable ="CameraLocations"
    Expression ="StudyAreas.StudyAreaID = CameraLocations.StudyAreaID"
    Flag =1
End
Begin OrderBy
    Expression ="[StudyAreas].[StudyAreaAbbr] & \" - \" & [CameraLocations].[LocationName]"
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
        dbText "Name" ="CameraLocations.AccessNotes"
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
    Bottom =590
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
End
