Operation =1
Option =0
Begin InputTables
    Name ="SELECT DISTINCT ImageID, SpeciesID FROM Detections WHERE StatusID<3"
    Alias ="Q"
    Name ="Species"
    Name ="qryPVComboRowSource"
End
Begin OutputColumns
    Expression ="qryPVComboRowSource.StudyAreaID"
    Expression ="qryPVComboRowSource.LocationID"
    Expression ="qryPVComboRowSource.FieldSeason"
    Expression ="qryPVComboRowSource.VisitID"
    Expression ="qryPVComboRowSource.ImgID"
    Expression ="qryPVComboRowSource.ImageDate"
    Expression ="qryPVComboRowSource.Highlight"
    Expression ="qryPVComboRowSource.ObsCount"
    Expression ="qryPVComboRowSource.Verified"
    Expression ="qryPVComboRowSource.Pending"
    Expression ="qryPVComboRowSource.MultiSp"
    Expression ="qryPVComboRowSource.NotNone"
    Expression ="Q.SpeciesID"
    Expression ="Species.ShortName"
End
Begin Joins
    LeftTable ="Q"
    RightTable ="Species"
    Expression ="Q.SpeciesID = Species.SpeciesID"
    Flag =1
    LeftTable ="qryPVComboRowSource"
    RightTable ="Q"
    Expression ="qryPVComboRowSource.ImgID = Q.ImageID"
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
        dbText "Name" ="Q.SpeciesID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPVComboRowSource.StudyAreaID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPVComboRowSource.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPVComboRowSource.FieldSeason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPVComboRowSource.ImgID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPVComboRowSource.ImageDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPVComboRowSource.Highlight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPVComboRowSource.NotNone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPVComboRowSource.VisitID"
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
    Bottom =370
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =1008
        Top =12
        Right =1152
        Bottom =156
        Top =0
        Name ="Q"
        Name =""
    End
    Begin
        Left =1200
        Top =12
        Right =1344
        Bottom =156
        Top =0
        Name ="Species"
        Name =""
    End
    Begin
        Left =739
        Top =15
        Right =883
        Bottom =159
        Top =0
        Name ="qryPVComboRowSource"
        Name =""
    End
End
