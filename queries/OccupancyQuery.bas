dbMemo "SQL" ="SELECT [StudyAreas].[StudyAreaAbbr] & ' - ' & [CameraLocations].[LocationName] A"
    "S Location, GetOccasion([CameraLocations].[LocationID],DateAdd('d',0,[StartDateT"
    "ime]),DateAdd('s',-1,DateAdd('d',10, [StartDateTime])),3, False, False) AS Occas"
    "ion1, GetOccasion([CameraLocations].[LocationID],DateAdd('d',10,[StartDateTime])"
    ",DateAdd('s',-1,DateAdd('d',20, [StartDateTime])),3, False, False) AS Occasion2,"
    " 1 AS Group1, CameraLocations.StudyAreaID AS StudyArea, Year([Visits].[VisitDate"
    "]) AS SurveyYear, Nz(CameraLocations.LatitudeDD,'.') AS CameraLat, Nz(CameraLoca"
    "tions.LongitudeDD,'.') AS CameraLong, DateAdd('yyyy', Year(Nz((SELECT Min(V.Acti"
    "veStart) FROM Visits AS V WHERE (((V.SetVisitID) = [Visits].[VisitID]))),[Visits"
    "].[VisitDate])) - 2014 + IIf(DatePart('y', Nz((SELECT Min(V.ActiveStart) FROM Vi"
    "sits AS V WHERE (((V.SetVisitID) = [Visits].[VisitID]))),[Visits].[VisitDate])) "
    "> 355, 1, 0), #12/1/2014 12:00:00 PM#) AS StartDateTime\015\012FROM (StudyAreas "
    "INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyArea"
    "ID) INNER JOIN Visits ON CameraLocations.LocationID = Visits.LocationID\015\012W"
    "HERE (((Visits.VisitTypeID) = 3))\015\012ORDER BY CameraLocations.LocationID, Vi"
    "sits.VisitDate;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Location"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2010"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SurveyYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Occasion1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StudyArea"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StartDateTime"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
    End
End
