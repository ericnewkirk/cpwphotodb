Operation =1
Option =0
Where ="(((IndividualDetections.IndividualID)>0))"
Begin InputTables
    Name ="IndependentDetections"
    Name ="IndividualDetections"
End
Begin OutputColumns
    Expression ="IndependentDetections.IndDetectionID"
    Alias ="KnownIndividuals"
    Expression ="Count(IndividualDetections.ID)"
End
Begin Joins
    LeftTable ="IndependentDetections"
    RightTable ="IndividualDetections"
    Expression ="IndependentDetections.IndDetectionID = IndividualDetections.IndDetectionID"
    Flag =1
End
Begin Groups
    Expression ="IndependentDetections.IndDetectionID"
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
        dbText "Name" ="IndependentDetections.IndDetectionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KnownIndividuals"
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
    Bottom =629
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="IndependentDetections"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="IndividualDetections"
        Name =""
    End
End
