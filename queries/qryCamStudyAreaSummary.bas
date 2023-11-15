dbMemo "SQL" ="SELECT qryCamDeploymentSummary.CamYear, qryCamDeploymentSummary.StudyAreaID, Min"
    "(qryCamDeploymentSummary.SetDate) AS MinSet, Max(qryCamDeploymentSummary.SetDate"
    ") AS MaxSet, Min(qryCamDeploymentSummary.PullDate) AS MinPull, Max(qryCamDeploym"
    "entSummary.PullDate) AS MaxPull, Round(Avg([DaysDeployed]),3) AS AvgDays, Min(qr"
    "yCamDeploymentSummary.DaysDeployed) AS MinDays, Max(qryCamDeploymentSummary.Days"
    "Deployed) AS MaxDays, Round(Avg([qryCamDeploymentSummary].[TotalPhotos]),3) AS A"
    "vgPhotos, Sum(qryCamDeploymentSummary.TotalPhotos) AS TotalPhotos, Count(qryCamD"
    "eploymentSummary.LocationID) AS CamerasSet, Count(qryCamDeploymentSummary.PullDa"
    "te) AS CamerasRetrieved, Round(Sum([TrapNights]),3) AS TtlEffort, Round(Min([Tra"
    "pNights]),3) AS MinEffort, Round(Max([TrapNights]),3) AS MaxEffort, Round(Avg([T"
    "rapNights]),3) AS AvgEffort\015\012FROM qryCamDeploymentSummary\015\012GROUP BY "
    "qryCamDeploymentSummary.CamYear, qryCamDeploymentSummary.StudyAreaID;\015\012"
dbMemo "Connect" =""
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
        dbText "Name" ="qryCamDeploymentSummary.StudyAreaID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TtlEffort"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
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
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
    End
End
