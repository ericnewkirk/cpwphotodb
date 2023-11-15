Operation =1
Option =0
Where ="(((qryCRSubquery.ModifiedStart) Between qryCRSubquery.OccasionStart And DateAdd("
    "'s', -1, DateAdd('d', 56, qryCRSubquery.OccasionStart))) AND ((qryCRSubquery.Spe"
    "ciesID)=65) AND ((qryCRSubquery.IndividualID)>0))"
Begin InputTables
    Name ="qryCRSubquery"
End
Begin OutputColumns
    Alias ="Session"
    Expression ="Left(Replace([qryCRSubquery].[StudyAreaName],' ',''),12) & DCount('*','Visits','"
        "VisitTypeID=3 And LocationID=' & [qryCRSubquery].[LocationID] & ' And VisitDate<"
        "=#' & [qryCRSubquery].[CameraSetDate] & '#')"
    Expression ="qryCRSubquery.IndividualID"
    Alias ="Occasion"
    Expression ="SECROccNumber(qryCRSubquery.OccasionStart, 14, qryCRSubquery.ModifiedStart)"
    Expression ="qryCRSubquery.LocationID"
    Alias ="StudyArea"
    Expression ="'# ' & qryCRSubquery.StudyAreaName"
    Alias ="Location"
    Expression ="qryCRSubquery.LocationName"
    Alias ="Individual"
    Expression ="qryCRSubquery.IndividualName"
End
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
        dbText "Name" ="Session"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Occasion"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryCRSubquery.IndividualID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryCRSubquery.LocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StudyArea"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1705
    Bottom =972
    Left =-1
    Top =-1
    Right =1689
    Bottom =444
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =30
        Top =19
        Right =174
        Bottom =163
        Top =0
        Name ="qryCRSubquery"
        Name =""
    End
End
