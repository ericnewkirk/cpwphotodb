dbMemo "SQL" ="SELECT CameraLocations.StudyAreaID, StudyAreas.StudyAreaName, CameraLocations.Lo"
    "cationID, CameraLocations.LocationName, CameraLocations.UTM_E, CameraLocations.U"
    "TM_N, SetVisits.VisitDate AS CameraSetDate, Year(SetVisits.VisitDate) AS CameraY"
    "ear, Nz(DMin(\"ActiveStart\",\"Visits\",\"SetVisitID=\"&SetVisits.VisitID),SetVi"
    "sits.VisitDate) AS OccasionStart, IndependentDetections.ModifiedStart, Independe"
    "ntDetections.SpeciesID, IndividualDetections.IndividualID, Individuals.Individua"
    "lName\015\012FROM ((StudyAreas INNER JOIN CameraLocations ON StudyAreas.StudyAre"
    "aID = CameraLocations.StudyAreaID) INNER JOIN ((Visits AS SetVisits INNER JOIN ("
    "SELECT Visits.SetVisitID, IndependentDetections.IndDetectionID\015\012FROM Visit"
    "s INNER JOIN (Photos INNER JOIN IndependentDetections ON Photos.ImageDate = Inde"
    "pendentDetections.DefaultStart) ON (Visits.LocationID = IndependentDetections.Lo"
    "cationID) AND (Visits.VisitID = Photos.VisitID)\015\012GROUP BY Visits.SetVisitI"
    "D, IndependentDetections.IndDetectionID)  AS VID ON SetVisits.VisitID = VID.SetV"
    "isitID) INNER JOIN IndependentDetections ON VID.IndDetectionID = IndependentDete"
    "ctions.IndDetectionID) ON CameraLocations.LocationID = SetVisits.LocationID) INN"
    "ER JOIN (Individuals INNER JOIN IndividualDetections ON Individuals.IndividualID"
    " = IndividualDetections.IndividualID) ON IndependentDetections.IndDetectionID = "
    "IndividualDetections.IndDetectionID\015\012WHERE (((IndependentDetections.Delete"
    "d)=False));\015\012"
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
        dbText "Name" ="CameraLocations.StudyAreaID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IndividualDetections.IndividualID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IndependentDetections.ModifiedStart"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="CameraLocations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IndependentDetections.SpeciesID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Individuals.IndividualName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StudyAreas.StudyAreaName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.UTM_E"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraLocations.UTM_N"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraSetDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CameraYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OccasionStart"
        dbLong "AggregateType" ="-1"
    End
End
