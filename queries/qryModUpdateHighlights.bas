dbMemo "SQL" ="UPDATE ModulePhotos INNER JOIN Photos ON ModulePhotos.ImageID = Photos.ImageID S"
    "ET Photos.Highlight = True\015\012WHERE (((ModulePhotos.Highlight)=True));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ModulePhotos.Highlight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos.Highlight"
        dbLong "AggregateType" ="-1"
    End
End
