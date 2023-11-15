dbMemo "SQL" ="TRANSFORM Nz(Sum(Q.MaxInd),0) AS Individuals\015\012SELECT Q.ImageID\015\012FROM"
    " ((SELECT Detections.ImageID, \015\012                      Detections.SpeciesID"
    ", \015\012                      Detections.DetailID, \015\012                   "
    "   Max(IIf([Detections].[SpeciesID]>0,[Detections].[Individuals],0)) AS MaxInd\015"
    "\012                    FROM Detections\015\012                    WHERE (((Dete"
    "ctions.StatusID)<3))\015\012                    GROUP BY Detections.ImageID, Det"
    "ections.SpeciesID, Detections.DetailID)  AS Q INNER JOIN Species ON Q.SpeciesID "
    "= Species.SpeciesID) LEFT JOIN DetectionDetails ON Q.DetailID = DetectionDetails"
    ".DetailID\015\012GROUP BY Q.ImageID\015\012PIVOT [Species].[CommonName] & IIf([D"
    "etectionDetails].[DetailID] Is Null,\"\",\" - \" & [DetectionDetails].[DetailTex"
    "t]);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
Begin
    Begin
        dbText "Name" ="Q.ImageID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="corsac fox"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Eurasian lynx"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="None"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="argali mountain sheep"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="domestic horse"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="human"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mongolian gazelle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mongolian marmot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="no ID possible"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Przewalski’s horse"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="red fox"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vehicle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="badger"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="domestic dog"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ground squirrel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="long-tailed ground squirrel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="mountain hare"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="red deer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Steppe polecat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="birds"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="domestic cattle/goat/sheep/yak"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="gray wolf"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpDetail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Species.CommonName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Abert's Squirrel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="American Elk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="American Robin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Black Bear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Black Bear - Female"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Black Bear - Male"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Black-billed Magpie"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bobcat"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Chipmunk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clark's Nutcracker"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Common Raven"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cooper's Hawk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Coyote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dark-eyed Junco"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gray Fox"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Merriam's Wild Turkey"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mountain Cottontail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mountain Lion"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mule Deer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Northern Flicker"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rock Squirrel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stellar's Jay"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Striped Skunk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unknown"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Western Spotted Skunk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
End
