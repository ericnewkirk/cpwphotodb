dbMemo "SQL" ="SELECT Q.Initials, Q.TtlDet AS TotalDetections, Q.VDet/(Q.TtlDet-Q.PDet) AS PctV"
    "erified, Q.DDet/(Q.TtlDet-Q.PDet) AS PctDeleted, Q.DelSp/(Q.TtlDet - Q.PDet) AS "
    "WrongSp, Q.PDet AS Pending\015\012FROM (SELECT Observers.Initials, Count(Detecti"
    "ons.DetectionID) AS TtlDet, Sum(IIf([Detections].[StatusID]=2,1,0)) AS VDet, Sum"
    "(IIf([Detections].[StatusID]=3,1,0)) AS DDet, Sum(IIf([Detections].[StatusID]=1,"
    "1,0)) AS PDet, SUM(IIF([Detections].[StatusID]=3 AND [VS].[ImageID] Is Null,1,0)"
    ") AS DelSp\015\012FROM Observers INNER JOIN (Detections LEFT JOIN (SELECT ImageI"
    "D, SpeciesID FROM Detections WHERE StatusID=2 GROUP BY ImageID, SpeciesID) AS VS"
    " ON Detections.ImageID = VS.ImageID AND Detections.SpeciesID = VS.SpeciesID) ON "
    "Observers.ObserverID = Detections.ObsID \015\012GROUP BY Observers.Initials)  AS"
    " Q;\015\012"
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
        dbText "Name" ="TotalDetections"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PctVerified"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PctDeleted"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Pending"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Q.Initials"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WrongSp"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
