dbMemo "SQL" ="INSERT INTO ObsPhone ( ObserverID, TypeID, PhoneNumber, Extension )\015\012SELEC"
    "T Observers.ObserverID, ModuleObsPhone.TypeID, ModuleObsPhone.PhoneNumber, Modul"
    "eObsPhone.Extension\015\012FROM ((ModuleObservers INNER JOIN Observers ON (Modul"
    "eObservers.FirstName = Observers.FirstName) AND (ModuleObservers.LastName = Obse"
    "rvers.LastName)) INNER JOIN ModuleObsPhone ON ModuleObservers.ObserverID = Modul"
    "eObsPhone.ObserverID) LEFT JOIN ObsPhone ON Observers.ObserverID = ObsPhone.Obse"
    "rverID\015\012WHERE (((ObsPhone.ID) Is Null));\015\012"
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
        dbText "Name" ="Observers.ObserverID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ObsPhone_1.TypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ObsPhone_1.PhoneNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ObsPhone_1.Extension"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ObsPhone.ID"
        dbLong "AggregateType" ="-1"
    End
End
