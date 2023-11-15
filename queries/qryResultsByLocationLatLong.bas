Operation =1
Option =0
Where ="(((CameraLocations.LatitudeDD) Is Not Null) AND ((CameraLocations.LongitudeDD) I"
    "s Not Null))"
Begin InputTables
    Name ="StudyAreas"
    Name ="CameraLocations"
    Name ="Visits"
    Name ="Photos"
    Name ="qryPhotoSpecies"
    Name ="Species"
End
Begin OutputColumns
    Alias ="StudyAreaID"
    Expression ="First(StudyAreas.StudyAreaID)"
    Alias ="StudyAreaName"
    Expression ="First(StudyAreas.StudyAreaName)"
    Expression ="CameraLocations.LocationID"
    Alias ="LocationName"
    Expression ="First(CameraLocations.LocationName)"
    Expression ="Species.SpeciesID"
    Alias ="CommonName"
    Expression ="First(Species.CommonName)"
    Alias ="Photos"
    Expression ="Count(Photos.ImageID)"
    Alias ="Latitude"
    Expression ="First(CameraLocations.LatitudeDD)"
    Alias ="Longitude"
    Expression ="First(CameraLocations.LongitudeDD)"
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
    LeftTable ="Photos"
    RightTable ="qryPhotoSpecies"
    Expression ="Photos.ImageID = qryPhotoSpecies.ImageID"
    Flag =1
    LeftTable ="qryPhotoSpecies"
    RightTable ="Species"
    Expression ="qryPhotoSpecies.SpeciesID = Species.SpeciesID"
    Flag =1
    LeftTable ="Visits"
    RightTable ="Photos"
    Expression ="Visits.VisitID = Photos.VisitID"
    Flag =1
End
Begin OrderBy
    Expression ="First(StudyAreas.StudyAreaName)"
    Flag =0
    Expression ="First(CameraLocations.LocationName)"
    Flag =0
    Expression ="First(Species.CommonName)"
    Flag =0
End
Begin Groups
    Expression ="CameraLocations.LocationID"
    GroupLevel =0
    Expression ="Species.SpeciesID"
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
        dbText "Name" ="CameraLocations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species.SpeciesID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CommonName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StudyAreaName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StudyAreaID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Latitude"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Longitude"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1418
    Bottom =891
    Left =-1
    Top =-1
    Right =1402
    Bottom =573
    Left =0
    Top =0
    ColumnsShown =543
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
        Name ="Photos"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="qryPhotoSpecies"
        Name =""
    End
    Begin
        Left =625
        Top =183
        Right =769
        Bottom =327
        Top =0
        Name ="Species"
        Name =""
    End
End
