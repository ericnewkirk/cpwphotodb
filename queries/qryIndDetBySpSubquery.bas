Operation =1
Option =0
Begin InputTables
    Name ="IndependentDetections"
    Name ="qryIndDetGroupsSummary"
    Name ="qryIndDetKnownIndividuals"
End
Begin OutputColumns
    Expression ="IndependentDetections.SpeciesID"
    Alias ="Detectons"
    Expression ="Count(IndependentDetections.IndDetectionID)"
    Alias ="TotalIndividuals"
    Expression ="Sum(qryIndDetGroupsSummary.TotalIndividuals)"
    Alias ="TotalAdults"
    Expression ="Sum(qryIndDetGroupsSummary.TotalAdults)"
    Alias ="TotalJuveniles"
    Expression ="Sum(qryIndDetGroupsSummary.TotalJuveniles)"
    Alias ="TotalSubadults"
    Expression ="Sum(qryIndDetGroupsSummary.TotalSubadults)"
    Alias ="TotalFemales"
    Expression ="Sum(qryIndDetGroupsSummary.TotalFemales)"
    Alias ="TotalMales"
    Expression ="Sum(qryIndDetGroupsSummary.TotalMales)"
    Alias ="TotalKnownInd"
    Expression ="Nz(Sum([qryIndDetKnownIndividuals].[KnownIndividuals]),0)"
End
Begin Joins
    LeftTable ="IndependentDetections"
    RightTable ="qryIndDetGroupsSummary"
    Expression ="IndependentDetections.IndDetectionID = qryIndDetGroupsSummary.IndDetectionID"
    Flag =1
    LeftTable ="IndependentDetections"
    RightTable ="qryIndDetKnownIndividuals"
    Expression ="IndependentDetections.IndDetectionID = qryIndDetKnownIndividuals.IndDetectionID"
    Flag =2
End
Begin Groups
    Expression ="IndependentDetections.SpeciesID"
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
        dbText "Name" ="IndependentDetections.SpeciesID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Detectons"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalAdults"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalJuveniles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalSubadults"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalFemales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalMales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalIndividuals"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalKnownInd"
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
    Bottom =561
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
        Name ="qryIndDetGroupsSummary"
        Name =""
    End
    Begin
        Left =241
        Top =170
        Right =385
        Bottom =314
        Top =0
        Name ="qryIndDetKnownIndividuals"
        Name =""
    End
End
