﻿dbMemo "SQL" ="SELECT [StudyAreas].[StudyAreaAbbr] & \" - \" & [CameraLocations].[LocationName]"
    " AS Location, CameraLocations.AccessNotes AS Access\015\012FROM StudyAreas INNER"
    " JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID\015"
    "\012ORDER BY CameraLocations.LocationID;\015\012"
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
        dbText "Name" ="Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Access"
        dbLong "AggregateType" ="-1"
    End
End
