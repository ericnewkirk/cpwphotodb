dbMemo "SQL" ="SELECT CameraLocations.StudyAreaID AS [Zone], Species.CommonName AS Species, For"
    "mat((DatePart('n',[ImageDate])+DatePart('h',[ImageDate])*60)/1440,'0.000') AS [T"
    "ime]\015\012FROM (Activity INNER JOIN Species ON Activity.SpeciesID = Species.Sp"
    "eciesID) INNER JOIN CameraLocations ON Activity.LocationID = CameraLocations.Loc"
    "ationID\015\012WHERE (((Activity.Include)=True) AND ((Species.SpeciesID)=44) AND"
    " ((CameraLocations.StudyAreaID)=6))\015\012ORDER BY CameraLocations.StudyAreaID,"
    " Species.CommonName, Activity.ImageDate - Int(Activity.ImageDate);\015\012"
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Zone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Time"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3645"
        dbBoolean "ColumnHidden" ="0"
    End
End
