dbMemo "SQL" ="SELECT qrySECRDetSubquery.Session, Sum(SECRSplitEffortString(qrySECRDetSubquery."
    "Effort, 1)) AS GroupEffort1, Sum(SECRSplitEffortString(qrySECRDetSubquery.Effort"
    ", 2)) AS GroupEffort2, Sum(SECRSplitEffortString(qrySECRDetSubquery.Effort, 3)) "
    "AS GroupEffort3, Sum(SECRSplitEffortString(qrySECRDetSubquery.Effort, 4)) AS Gro"
    "upEffort4\015\012FROM qrySECRDetSubquery\015\012GROUP BY qrySECRDetSubquery.Sess"
    "ion;\015\012"
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
        dbText "Name" ="GroupEffort1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="GroupEffort2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2865"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="GroupEffort3"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3285"
        dbBoolean "ColumnHidden" ="0"
    End
End
