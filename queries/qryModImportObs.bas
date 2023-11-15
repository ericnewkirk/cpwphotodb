dbMemo "SQL" ="INSERT INTO Observers ( LastName, FirstName, Initials, Role, Email )\015\012SELE"
    "CT ModuleObservers.LastName, ModuleObservers.FirstName, ModuleObservers.Initials"
    ", ModuleObservers.Role, ModuleObservers.Email\015\012FROM ModuleObservers LEFT J"
    "OIN Observers ON (ModuleObservers.LastName = Observers.LastName) AND (ModuleObse"
    "rvers.FirstName = Observers.FirstName)\015\012WHERE (((Observers.ObserverID) Is "
    "Null) AND ((ModuleObservers.ObserverID) In (SELECT ObsID FROM ModuleDetections G"
    "ROUP BY ObsID)));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Observers.ObserverID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Observers1.Email"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ModuleObservers.LastName"
        dbLong "AggregateType" ="-1"
    End
End
