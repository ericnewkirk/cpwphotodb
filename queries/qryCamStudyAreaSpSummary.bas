Operation =1
Option =0
Begin InputTables
    Name ="Photos"
    Name ="CameraLocations"
    Name ="Visits"
    Alias ="SetVisits"
    Name ="Visits"
    Name ="Species"
    Name ="qryPhotoSpecies"
End
Begin OutputColumns
    Alias ="CameraYear"
    Expression ="Year([SetVisits].[VisitDate])"
    Expression ="CameraLocations.StudyAreaID"
    Expression ="Species.CommonName"
    Alias ="Images"
    Expression ="Count(qryPhotoSpecies.ImageID)"
End
Begin Joins
    LeftTable ="Visits"
    RightTable ="Photos"
    Expression ="Visits.VisitID = Photos.VisitID"
    Flag =1
    LeftTable ="SetVisits"
    RightTable ="Visits"
    Expression ="SetVisits.VisitID = Visits.SetVisitID"
    Flag =1
    LeftTable ="CameraLocations"
    RightTable ="SetVisits"
    Expression ="CameraLocations.LocationID = SetVisits.LocationID"
    Flag =1
    LeftTable ="Photos"
    RightTable ="qryPhotoSpecies"
    Expression ="Photos.ImageID = qryPhotoSpecies.ImageID"
    Flag =1
    LeftTable ="Species"
    RightTable ="qryPhotoSpecies"
    Expression ="Species.SpeciesID = qryPhotoSpecies.SpeciesID"
    Flag =1
End
Begin Groups
    Expression ="Year([SetVisits].[VisitDate])"
    GroupLevel =0
    Expression ="CameraLocations.StudyAreaID"
    GroupLevel =0
    Expression ="Species.CommonName"
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
        dbText "Name" ="CameraLocations.StudyAreaID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Images"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species.CommonName"
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
    Bottom =495
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =629
        Top =169
        Right =773
        Bottom =313
        Top =0
        Name ="Photos"
        Name =""
    End
    Begin
        Left =49
        Top =179
        Right =193
        Bottom =323
        Top =0
        Name ="CameraLocations"
        Name =""
    End
    Begin
        Left =242
        Top =177
        Right =386
        Bottom =321
        Top =0
        Name ="SetVisits"
        Name =""
    End
    Begin
        Left =429
        Top =173
        Right =573
        Bottom =317
        Top =0
        Name ="Visits"
        Name =""
    End
    Begin
        Left =1001
        Top =167
        Right =1145
        Bottom =311
        Top =0
        Name ="Species"
        Name =""
    End
    Begin
        Left =812
        Top =168
        Right =956
        Bottom =312
        Top =0
        Name ="qryPhotoSpecies"
        Name =""
    End
End
