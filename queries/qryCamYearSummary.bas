Operation =1
Option =0
Begin InputTables
    Name ="qryCamDeploymentSummary"
End
Begin OutputColumns
    Expression ="qryCamDeploymentSummary.CamYear"
    Alias ="MinSet"
    Expression ="Min(qryCamDeploymentSummary.SetDate)"
    Alias ="MaxSet"
    Expression ="Max(qryCamDeploymentSummary.SetDate)"
    Alias ="MinPull"
    Expression ="Min(qryCamDeploymentSummary.PullDate)"
    Alias ="MaxPull"
    Expression ="Max(qryCamDeploymentSummary.PullDate)"
    Alias ="AvgDays"
    Expression ="Avg(qryCamDeploymentSummary.DaysDeployed)"
    Alias ="MinDays"
    Expression ="Min(qryCamDeploymentSummary.DaysDeployed)"
    Alias ="MaxDays"
    Expression ="Max(qryCamDeploymentSummary.DaysDeployed)"
    Alias ="AvgPhotos"
    Expression ="Avg(qryCamDeploymentSummary.TotalPhotos)"
    Alias ="TotalPhotos"
    Expression ="Sum(qryCamDeploymentSummary.TotalPhotos)"
    Alias ="CamerasSet"
    Expression ="Count(qryCamDeploymentSummary.LocationID)"
    Alias ="CamerasRetrieved"
    Expression ="Count(qryCamDeploymentSummary.PullDate)"
    Alias ="FilterDate"
    Expression ="Min(qryCamDeploymentSummary.SetDate)"
    Alias ="TtlEffort"
    Expression ="Sum(qryCamDeploymentSummary.TrapNights)"
    Alias ="MinEffort"
    Expression ="Min(qryCamDeploymentSummary.TrapNights)"
    Alias ="MaxEffort"
    Expression ="Max(qryCamDeploymentSummary.TrapNights)"
    Alias ="AvgEffort"
    Expression ="Avg(qryCamDeploymentSummary.TrapNights)"
End
Begin Groups
    Expression ="qryCamDeploymentSummary.CamYear"
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
        dbText "Name" ="AvgPhotos"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TotalPhotos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryCamDeploymentSummary.CamYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgDays"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MinDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinSet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxSet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinPull"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxPull"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CamerasSet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CamerasRetrieved"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FilterDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TtlEffort"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinEffort"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxEffort"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgEffort"
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
    Bottom =450
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =444
        Top =8
        Right =588
        Bottom =152
        Top =0
        Name ="qryCamDeploymentSummary"
        Name =""
    End
End
