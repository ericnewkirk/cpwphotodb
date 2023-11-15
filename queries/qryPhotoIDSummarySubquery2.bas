Operation =1
Option =0
Begin InputTables
    Name ="qryPhotoIDSummarySubquery1"
End
Begin OutputColumns
    Expression ="qryPhotoIDSummarySubquery1.VisitID"
    Alias ="Photos"
    Expression ="Count(qryPhotoIDSummarySubquery1.ImageID)"
    Alias ="NoID"
    Expression ="Sum(qryPhotoIDSummarySubquery1.NoIDs)"
    Alias ="ID"
    Expression ="Sum(IIf([qryPhotoIDSummarySubquery1].[PendingIDs]+[qryPhotoIDSummarySubquery1].["
        "VerifiedIDs]>0,1,0))"
    Alias ="VerifiedID"
    Expression ="Sum(IIf([qryPhotoIDSummarySubquery1].[VerifiedIDs]>0,1,0))"
End
Begin Groups
    Expression ="qryPhotoIDSummarySubquery1.VisitID"
    GroupLevel =0
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
        dbText "Name" ="qryPhotoIDSummarySubquery1.VisitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NoID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Photos"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VerifiedID"
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
    Bottom =165
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qryPhotoIDSummarySubquery1"
        Name =""
    End
End
