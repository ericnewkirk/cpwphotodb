dbMemo "SQL" ="SELECT [Individual] & ': ' & [Session] AS Label, Max(IIf([Occasion]=1,1,0)) AS O"
    "ccasion1, Max(IIf([Occasion]=2,1,0)) AS Occasion2, Max(IIf([Occasion]=3,1,0)) AS"
    " Occasion3, Max(IIf([Occasion]=4,1,0)) AS Occasion4, Max(IIf([Session]='MolasPas"
    "s1',1,0)) AS Group1, Max(IIf([Session]='Telluride1',1,0)) AS Group2\015\012FROM "
    "SECRQuery\015\012GROUP BY [Individual] & ': ' & [Session];\015\012"
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
        dbText "Name" ="Label"
        dbInteger "ColumnWidth" ="4050"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Occasion1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Occasion2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Occasion3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Occasion4"
        dbLong "AggregateType" ="-1"
    End
End
