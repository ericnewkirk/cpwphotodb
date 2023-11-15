dbMemo "SQL" ="INSERT INTO Detections ( DetectionID, ImageID, SpeciesID, DetailID, Individuals,"
    " ObsID, Comments )\015\012SELECT NewDetections.DetectionID, NewDetections.ImageI"
    "D, NewDetections.SpeciesID, NewDetections.DetailID, NewDetections.Individuals, N"
    "ewDetections.ObsID, NewDetections.Comments\015\012FROM NewDetections;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
