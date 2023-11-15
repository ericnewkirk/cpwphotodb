Attribute VB_Name = "CoordinateConverter"
Option Compare Database
Option Explicit

'Code written by Eric Newkirk 5/27/2016
'Formulas based on:
'http://www.uwgb.edu/dutchs/usefuldata/utmformulas.htm

'Spheroid parameters
Private a As Double
Private b As Double
Private f As Double
Private e As Double
Private n As Double
Private AA As Double
Private pi As Double
Private Const k0 As Double = 0.9996
Private Const FalseEasting As Long = 500000

'Series variables for LL to UTM
Private alpha1 As Double
Private alpha2 As Double
Private alpha3 As Double
Private alpha4 As Double
Private alpha5 As Double
Private alpha6 As Double
Private alpha7 As Double
Private alpha8 As Double
Private alpha9 As Double
Private alpha10 As Double

'Series variables for UTM to LL
Private beta1 As Double
Private beta2 As Double
Private beta3 As Double
Private beta4 As Double
Private beta5 As Double
Private beta6 As Double
Private beta7 As Double
Private beta8 As Double
Private beta9 As Double
Private beta10 As Double

'List of available datums
Public Enum Datum
    WGS84 = 1
    NAD83 = 2
    GRS80 = 3
    WGS72 = 4
    Australian1965 = 5
    Krasovsky1940 = 6
    NAD27 = 7
    Intl1924 = 8
    Hayford1909 = 9
    Clarke1880 = 10
    Clarke1866 = 11
    Airy1830 = 12
    Bessel1841 = 13
    Everest1830 = 14
End Enum

Public Function GetDefaultZone(Lon As Double) As Integer
'Calculate UTM zone for a given longitude

If Lon < 0 Then
    GetDefaultZone = Int((180 + Lon) / 6) + 1
Else
    GetDefaultZone = Int(Abs((Lon) / 6) + 31)
End If

End Function

Private Function GetMeridian(Zone As Integer) As Integer
'Find the central meridian of a UTM zone in degrees

GetMeridian = Zone * 6 - 183

End Function

Private Function HyperArcSine(dblIn As Double) As Double
'Hyperbolic arcsine (ASINH in excel)

HyperArcSine = Log(dblIn + Sqr(dblIn * dblIn + 1))

End Function

Private Function HyperArcTangent(dblIn As Double) As Double
'Hyperbolic arctangent (ATNH in excel)

HyperArcTangent = Log((1 + dblIn) / (1 - dblIn)) / 2

End Function

Private Function HyperCosine(dblIn As Double) As Double
'Hyperbolic cosine (COSH in excel)

HyperCosine = (Exp(dblIn) + Exp(-1 * dblIn)) / 2

End Function

Private Function HyperSine(dblIn As Double) As Double
'Hyperbolic sine (SINH in excel)

HyperSine = (Exp(dblIn) - Exp(-1 * dblIn)) / 2

End Function

Public Function LLDecToLLDMS(LatOrLon As Double) As String
'Converts lat lon decimal to deg min sec

Dim sign As String
Dim d As Integer
Dim m As Integer
Dim s As Double
Dim strTemp As String

'Check for negative value
If LatOrLon < 0 Then
    sign = "-"
End If

'Get degrees
d = Int(Abs(LatOrLon))
strTemp = CStr(d) & "d "

'Get minutes
m = Int((Abs(LatOrLon) - CDbl(d)) * 60)
strTemp = strTemp & CStr(m) & "m "

s = (((Abs(LatOrLon) - CDbl(d)) * 60) - CDbl(m)) * 60
strTemp = strTemp & CStr(s) & "s"

LLDecToLLDMS = sign & strTemp

End Function

Public Function LLDecToUTM(Lat As Double, Lon As Double, _
    TargetDatum As Datum, Optional TargetZone As Integer = 0) As String
'Converts decimal lat lon coordinates to UTM string

Dim Zone As Integer
Dim LatTemp As Double
Dim LonTemp As Double
Dim Meridian As Integer
Dim EastOfCM As Integer

Dim ConfLat As Double
Dim tau As Double
Dim tauprime As Double
Dim xi As Double
Dim xiprime As Double
Dim eta As Double
Dim etaprime As Double

'Set values for a and b (spheroid diameters)
SetSpheroid TargetDatum

'Assign zone if necessary
If TargetZone = 0 Then
    Zone = GetDefaultZone(Lon)
Else
    Zone = TargetZone
End If

'Get other parameters
Meridian = GetMeridian(Zone)
If Lon > Meridian Then
    EastOfCM = 1
Else
    EastOfCM = -1
End If

'Convert coordinates to radians
LatTemp = Abs(Lat) * pi / 180
LonTemp = Abs(Lon - Meridian) * pi / 180

'Run the crazy stuff
tau = Tan(LatTemp)
ConfLat = Atn(HyperSine(HyperArcSine(Tan(LatTemp)) - e * HyperArcTangent(e * Sin(LatTemp))))
tauprime = Tan(ConfLat)
xiprime = Atn(tauprime / Cos(LonTemp))
etaprime = HyperArcSine(Sin(LonTemp) / Sqr(tauprime ^ 2 + (Cos(LonTemp) ^ 2)))
xi = xiprime + alpha1 * Sin(2 * xiprime) * HyperCosine(2 * etaprime) + _
    alpha2 * Sin(4 * xiprime) * HyperCosine(4 * etaprime) + _
    alpha3 * Sin(6 * xiprime) * HyperCosine(6 * etaprime) + _
    alpha4 * Sin(8 * xiprime) * HyperCosine(8 * etaprime) + _
    alpha5 * Sin(10 * xiprime) * HyperCosine(10 * etaprime) + _
    alpha6 * Sin(12 * xiprime) * HyperCosine(12 * etaprime) + _
    alpha7 * Sin(14 * xiprime) * HyperCosine(14 * etaprime)
eta = etaprime + alpha1 * Cos(2 * xiprime) * HyperSine(2 * etaprime) + _
    alpha2 * Cos(4 * xiprime) * HyperSine(4 * etaprime) + _
    alpha3 * Cos(6 * xiprime) * HyperSine(6 * etaprime) + _
    alpha4 * Cos(8 * xiprime) * HyperSine(8 * etaprime) + _
    alpha5 * Cos(10 * xiprime) * HyperSine(10 * etaprime) + _
    alpha6 * Cos(12 * xiprime) * HyperSine(12 * etaprime) + _
    alpha6 * Cos(14 * xiprime) * HyperSine(14 * etaprime)

'Get the output
LonTemp = eta * AA * k0
LatTemp = xi * AA * k0
If Lat < 0 Then
    LatTemp = 10000000 - LatTemp
End If
LonTemp = FalseEasting + EastOfCM * LonTemp

LLDecToUTM = "Z" & [Zone] & " " & LonTemp & ", " & LatTemp

End Function

Public Function LLDecToUTME(Lat As Double, Lon As Double, _
    TargetDatum As Datum, Optional TargetZone As Integer = 0) As Double
'Retrieves UTM easting from decimal lat lon coordinates

Dim Zone As Integer
Dim LatTemp As Double
Dim LonTemp As Double
Dim Meridian As Integer
Dim EastOfCM As Integer

Dim ConfLat As Double
Dim tau As Double
Dim tauprime As Double
Dim xi As Double
Dim xiprime As Double
Dim eta As Double
Dim etaprime As Double

'Set values for a and b (spheroid diameters)
SetSpheroid TargetDatum

'Assign zone if necessary
If TargetZone = 0 Then
    Zone = GetDefaultZone(Lon)
Else
    Zone = TargetZone
End If

'Get other parameters
Meridian = GetMeridian(Zone)
If Lon > Meridian Then
    EastOfCM = 1
Else
    EastOfCM = -1
End If

'Convert coordinates to radians
LatTemp = Abs(Lat) * pi / 180
LonTemp = Abs(Lon - Meridian) * pi / 180

'Run the crazy stuff
tau = Tan(LatTemp)
ConfLat = Atn(HyperSine(HyperArcSine(Tan(LatTemp)) - _
    e * HyperArcTangent(e * Sin(LatTemp))))
tauprime = Tan(ConfLat)
xiprime = Atn(tauprime / Cos(LonTemp))
etaprime = HyperArcSine(Sin(LonTemp) / Sqr(tauprime ^ 2 + (Cos(LonTemp) ^ 2)))
xi = xiprime + alpha1 * Sin(2 * xiprime) * HyperCosine(2 * etaprime) + _
    alpha2 * Sin(4 * xiprime) * HyperCosine(4 * etaprime) + _
    alpha3 * Sin(6 * xiprime) * HyperCosine(6 * etaprime) + _
    alpha4 * Sin(8 * xiprime) * HyperCosine(8 * etaprime) + _
    alpha5 * Sin(10 * xiprime) * HyperCosine(10 * etaprime) + _
    alpha6 * Sin(12 * xiprime) * HyperCosine(12 * etaprime) + _
    alpha7 * Sin(14 * xiprime) * HyperCosine(14 * etaprime)
eta = etaprime + alpha1 * Cos(2 * xiprime) * HyperSine(2 * etaprime) + _
    alpha2 * Cos(4 * xiprime) * HyperSine(4 * etaprime) + _
    alpha3 * Cos(6 * xiprime) * HyperSine(6 * etaprime) + _
    alpha4 * Cos(8 * xiprime) * HyperSine(8 * etaprime) + _
    alpha5 * Cos(10 * xiprime) * HyperSine(10 * etaprime) + _
    alpha6 * Cos(12 * xiprime) * HyperSine(12 * etaprime) + _
    alpha6 * Cos(14 * xiprime) * HyperSine(14 * etaprime)

'Get the output
LonTemp = eta * AA * k0
LonTemp = FalseEasting + EastOfCM * LonTemp

LLDecToUTME = LonTemp

End Function

Public Function LLDecToUTMN(Lat As Double, Lon As Double, _
    TargetDatum As Datum, Optional TargetZone As Integer = 0) As Double
'Retrieves UTM northing from decimal lat lon coordinates

Dim Zone As Integer
Dim LatTemp As Double
Dim LonTemp As Double
Dim Meridian As Integer
Dim EastOfCM As Integer

Dim ConfLat As Double
Dim tau As Double
Dim tauprime As Double
Dim xi As Double
Dim xiprime As Double
Dim eta As Double
Dim etaprime As Double

'Set values for a and b (spheroid diameters)
SetSpheroid TargetDatum

'Assign zone if necessary
If TargetZone = 0 Then
    Zone = GetDefaultZone(Lon)
Else
    Zone = TargetZone
End If

'Get other parameters
Meridian = GetMeridian(Zone)
If Lon > Meridian Then
    EastOfCM = 1
Else
    EastOfCM = -1
End If

'Convert coordinates to radians
LatTemp = Abs(Lat) * pi / 180
LonTemp = Abs(Lon - Meridian) * pi / 180

'Run the crazy stuff
tau = Tan(LatTemp)
ConfLat = Atn(HyperSine(HyperArcSine(Tan(LatTemp)) - e * HyperArcTangent(e * Sin(LatTemp))))
tauprime = Tan(ConfLat)
xiprime = Atn(tauprime / Cos(LonTemp))
etaprime = HyperArcSine(Sin(LonTemp) / Sqr(tauprime ^ 2 + (Cos(LonTemp) ^ 2)))
xi = xiprime + alpha1 * Sin(2 * xiprime) * HyperCosine(2 * etaprime) + _
    alpha2 * Sin(4 * xiprime) * HyperCosine(4 * etaprime) + _
    alpha3 * Sin(6 * xiprime) * HyperCosine(6 * etaprime) + _
    alpha4 * Sin(8 * xiprime) * HyperCosine(8 * etaprime) + _
    alpha5 * Sin(10 * xiprime) * HyperCosine(10 * etaprime) + _
    alpha6 * Sin(12 * xiprime) * HyperCosine(12 * etaprime) + _
    alpha7 * Sin(14 * xiprime) * HyperCosine(14 * etaprime)
eta = etaprime + alpha1 * Cos(2 * xiprime) * HyperSine(2 * etaprime) + _
    alpha2 * Cos(4 * xiprime) * HyperSine(4 * etaprime) + _
    alpha3 * Cos(6 * xiprime) * HyperSine(6 * etaprime) + _
    alpha4 * Cos(8 * xiprime) * HyperSine(8 * etaprime) + _
    alpha5 * Cos(10 * xiprime) * HyperSine(10 * etaprime) + _
    alpha6 * Cos(12 * xiprime) * HyperSine(12 * etaprime) + _
    alpha6 * Cos(14 * xiprime) * HyperSine(14 * etaprime)

'Get the output
LatTemp = xi * AA * k0
If Lat < 0 Then
    LatTemp = 10000000 - LatTemp
End If

LLDecToUTMN = LatTemp

End Function

Public Function LLDMSToLLDec(LatOrLon As String) As Double
'Converts lat lon in deg min sec to decimal degrees

Dim chr As String
Dim d As String
Dim m As String
Dim s As String
Dim iStart As Integer
Dim iLen As Integer
Dim strTemp As String
Dim iSign As Integer

iStart = 1
iLen = 1
strTemp = LatOrLon
chr = Mid(strTemp, iStart, 1)
Do Until IsNumeric(chr) Or iStart > Len(strTemp)
    iStart = iStart + 1
    chr = Mid(strTemp, iStart, 1)
Loop

iSign = 1
If iStart > 1 Then
    If Mid(strTemp, iStart - 1, 1) = "-" Then
        iSign = -1
    End If
End If

chr = Mid(strTemp, iStart, iLen + 1)
Do Until Not IsNumeric(chr) Or iStart > Len(strTemp)
    iLen = iLen + 1
    chr = Mid(strTemp, iStart, iLen + 1)
Loop
d = Mid(strTemp, iStart, iLen)

If Len(chr) = Len(strTemp) Then
    GoTo FunctionExit
End If

strTemp = Mid(strTemp, iStart + iLen)
iStart = 1
iLen = 1
chr = Mid(strTemp, iStart, 1)
Do Until IsNumeric(chr) Or iStart > Len(strTemp)
    iStart = iStart + 1
    chr = Mid(strTemp, iStart, 1)
Loop
chr = Mid(strTemp, iStart, iLen + 1)
Do Until Not IsNumeric(chr)
    iLen = iLen + 1
    chr = Mid(strTemp, iStart, iLen + 1)
Loop
m = Mid(strTemp, iStart, iLen)

If Len(chr) = Len(strTemp) Then
    GoTo FunctionExit
End If

strTemp = Mid(strTemp, iStart + iLen)
iStart = 1
iLen = 1
chr = Mid(strTemp, iStart, 1)
Do Until IsNumeric(chr) Or iStart > Len(strTemp)
    iStart = iStart + 1
    chr = Mid(strTemp, iStart, 1)
Loop
chr = Mid(strTemp, iStart, iLen + 1)
Do Until Not IsNumeric(chr)
    iLen = iLen + 1
    If Len(strTemp) = iLen + iStart - 1 Then
        Exit Do
    End If
    chr = Mid(strTemp, iStart, iLen + 1)
Loop
s = Mid(strTemp, iStart, iLen)

FunctionExit:
    If Len(d) = 0 Then
        d = "0"
    End If
    If Len(m) = 0 Then
        m = "0"
    End If
    If Len(s) = 0 Then
        s = "0"
    End If
    LLDMSToLLDec = iSign * (CDbl(d) + CDbl(m) / 60 + CDbl(s) / 3600)

End Function

Public Function LLDMSToUTM(Lat As String, Lon As String, _
    TargetDatum As Datum, Optional TargetZone As Integer = 0) As String
'Converts lat lon deg min sec coordinates to UTM string

Dim LatDec As Double
Dim LonDec As Double

LatDec = LLDMSToLLDec(Lat)
LonDec = LLDMSToLLDec(Lon)

LLDMSToUTM = LLDecToUTM(LatDec, LonDec, TargetDatum, TargetZone)

End Function

Private Sub SetSpheroid(TargetDatum As Datum)
'Sets all the necessary spheroid parameters for a given datum

'Get spheroid shape
'a = equatorial radius
'b = polar radius
Select Case TargetDatum
    Case WGS84
        a = 6378137
        b = 6356752.314
    Case NAD83
        a = 6378137
        b = 6356752.314
    Case GRS80
        a = 6378137
        b = 6356752.3
    Case WGS72
        a = 6378135
        b = 6356750
    Case Australian1965
        a = 6378160
        b = 6356774.7
    Case Krasovsky1940
        a = 6378245
        b = 6356863
    Case NAD27
        a = 6378206.4
        b = 6356583.8
    Case Intl1924
        a = 6378388
        b = 6356911.9
    Case Hayford1909
        a = 6378388
        b = 6356911.9
    Case Clarke1880
        a = 6378249.1
        b = 6356514.9
    Case Clarke1866
        a = 6378206.4
        b = 6356583.8
    Case Airy1830
        a = 6377563.4
        b = 6356256.9
    Case Bessel1841
        a = 6377397.2
        b = 6356079#
    Case Everest1830
        a = 6377276.3
        b = 6356075.4
End Select

'Calculate basic variables
f = (a - b) / a
e = Sqr(1 - (b / a) ^ 2)
n = (a - b) / (a + b)
AA = (a / (1 + n)) * (1 + (1 / 4) * n ^ 2 + (1 / 64) * n ^ 4 + _
    (1 / 256) * n ^ 6 + (25 / 16384) * n ^ 8 + (49 / 65536) * n ^ 10)
pi = Atn(1) * 4

'Calculate series variables
alpha1 = (1 / 2) * n - (2 / 3) * n ^ 2 + (5 / 16) * n ^ 3 + (41 / 180) * n ^ 4 - _
    (127 / 288) * n ^ 5 + (7891 / 37800) * n ^ 6 + (72161 / 387072) * n ^ 7 - _
    (18975107 / 50803200) * n ^ 8 + (60193001 / 290304000) * n ^ 9 + _
    (134592031 / 1026432000) * n ^ 10
alpha2 = (13 / 48) * n ^ 2 - (3 / 5) * n ^ 3 + (557 / 1440) * n ^ 4 + _
    (281 / 630) * n ^ 5 - (1983433 / 1935360) * n ^ 6 + (13769 / 28800) * n ^ 7 + _
    (148003883 / 174182400) * n ^ 8 - (705286231 / 465696000) * n ^ 9 + _
    (1703267974087# / 3218890752000#) * n ^ 10
alpha3 = (61 / 240) * n ^ 3 - (103 / 140) * n ^ 4 + (15061 / 26880) * n ^ 5 + _
    (167603 / 181440) * n ^ 6 - (67102379 / 29030400) * n ^ 7 + _
    (79682431 / 79833600) * n ^ 8 + (6304945039# / 2128896000) * n ^ 9 - _
    (6601904925257# / 1307674368000#) * n ^ 10
alpha4 = (49561 / 161280) * n ^ 4 - (179 / 168) * n ^ 5 + _
    (6601661 / 7257600) * n ^ 6 + (97445 / 49896) * n ^ 7 - _
    (40176129013# / 7664025600#) * n ^ 8 + (138471097 / 66528000) * n ^ 9 + _
    (48087451385201# / 5230697472000#) * n ^ 10
alpha5 = (34729 / 80640) * n ^ 5 - (3418889 / 1995840) * n ^ 6 + _
    (14644087 / 9123840) * n ^ 7 + (2605413599# / 622702080) * n ^ 8 - _
    (31015475399# / 2583060480#) * n ^ 9 + (5820486440369# / 1307674368000#) * n ^ 10
alpha6 = (212378941 / 319334400) * n ^ 6 - (30705481 / 10378368) * n ^ 7 + _
    (175214326799# / 58118860800#) * n ^ 8 + (870492877 / 96096000) * n ^ 9 - _
    (1.328004581729E+15 / 47823519744000#) * n ^ 10
alpha7 = (1522256789 / 1383782400) * n ^ 7 - _
    (16759934899# / 3113510400#) * n ^ 8 + (1315149374443# / 221405184000#) * n ^ 9 + _
    (71809987837451# / 3629463552000#) * n ^ 10
alpha8 = (1424729850961# / 743921418240#) * n ^ 8 - _
    (256783708069# / 25204608000#) * n ^ 9 + _
    (2.46874929298989E+15 / 203249958912000#) * n ^ 10
alpha9 = (21091646195357# / 6080126976000#) * n ^ 9 - _
    (6.71961821383558E+16 / 3.379030566912E+15) * n ^ 10
alpha10 = (7.79115156232328E+16 / 1.2014330904576E+16) * n ^ 10

beta1 = (1 / 2) * n - (2 / 3) * n ^ 2 + (37 / 96) * n ^ 3 - (1 / 360) * n ^ 4 - _
    (81 / 512) * n ^ 5 + (96199 / 604800) * n ^ 6 - (5406467 / 38707200) * n ^ 7 + _
    (7944359 / 67737600) * n ^ 8 - (7378753979# / 97542144000#) * n ^ 9 + _
    (25123531261# / 804722688000#) * n ^ 10
beta2 = (1 / 48) * n ^ 2 + (1 / 15) * n ^ 3 - (437 / 1440) * n ^ 4 + _
    (46 / 105) * n ^ 5 - (1118711 / 3870720) * n ^ 6 + (51841 / 1209600) * n ^ 7 + _
    (24749483 / 348364800) * n ^ 8 - (115295683 / 1397088000) * n ^ 9 + _
    (5487737251099# / 51502252032000#) * n ^ 10
beta3 = (17 / 480) * n ^ 3 - (37 / 840) * n ^ 4 - (209 / 4480) * n ^ 5 + _
    (5569 / 90720) * n ^ 6 + (9261899 / 58060800) * n ^ 7 - _
    (6457463 / 17740800) * n ^ 8 + (2473691167# / 9289728000#) * n ^ 9 - _
    (852549456029# / 20922789888000#) * n ^ 10
beta4 = (4397 / 161280) * n ^ 4 - (11 / 504) * n ^ 5 - (830251 / 7257600) * n ^ 6 + _
    (466511 / 2494800) * n ^ 7 + (324154477 / 7664025600#) * n ^ 8 - _
    (937932223 / 3891888000#) * n ^ 9 - (89112264211# / 5230697472000#) * n ^ 10
beta5 = (4583 / 161280) * n ^ 5 - (108847 / 3991680) * n ^ 6 - _
    (8005831 / 63866880) * n ^ 7 + (22894433 / 124540416) * n ^ 8 + _
    (112731569449# / 557941063680#) * n ^ 9 - _
    (5391039814733# / 10461394944000#) * n ^ 10
beta6 = (20648693 / 638668800) * n ^ 6 - (16363163 / 518918400) * n ^ 7 - _
    (2204645983# / 12915302400#) * n ^ 8 + (4543317553# / 18162144000#) * n ^ 9 + _
    (54894890298749# / 167382319104000#) * n ^ 10
beta7 = (219941297 / 5535129600#) * n ^ 7 - (497323811 / 12454041600#) * n ^ 8 - _
    (79431132943# / 332107776000#) * n ^ 9 + (4346429528407# / 12703122432000#) * n ^ 10
beta8 = (191773887257# / 3719607091200#) * n ^ 8 - _
    (17822319343# / 336825216000#) * n ^ 9 - _
    (497155444501631# / 1.422749712384E+15) * n ^ 10
beta9 = (11025641854267# / 158083301376000#) * n ^ 9 - _
    (492293158444691# / 6.758061133824E+15) * n ^ 10
beta10 = (7.02850453042962E+15 / 7.2085985427456E+16) * n ^ 10

End Sub

Public Function UTMToLatDec(UTME As Double, UTMN As Double, UTMDatum As Datum, _
    UTMZone As Integer, Optional Hem As String = "N") As Double
'Retrieves decimal lat from UTM coordinates

Dim LatTemp As Double
Dim LonTemp As Double
Dim Meridian As Integer
Dim EastOfCM As Integer
Dim i As Integer

Dim tau As Double
Dim tauprime As Double
Dim ftau As Double
Dim dtau As Double
Dim sigma As Double
Dim xi As Double
Dim xiprime As Double
Dim eta As Double
Dim etaprime As Double

'Set values for a and b (spheroid diameters)
SetSpheroid UTMDatum

'Get other parameters
Meridian = GetMeridian(UTMZone)
LonTemp = UTME - FalseEasting
If LonTemp > 0 Then
    EastOfCM = 1
Else
    EastOfCM = -1
End If
eta = LonTemp / (k0 * AA)

LatTemp = UTMN
If Hem = "S" Then
    LatTemp = 10000000 - LatTemp
End If
xi = LatTemp / (k0 * AA)

'Run the crazy stuff
xiprime = xi - (beta1 * Sin(2 * xi) * HyperCosine(2 * eta) + _
    beta2 * Sin(4 * xi) * HyperCosine(4 * eta) + _
    beta3 * Sin(6 * xi) * HyperCosine(6 * eta) + _
    beta4 * Sin(8 * xi) * HyperCosine(8 * eta) + _
    beta5 * Sin(10 * xi) * HyperCosine(10 * eta) + _
    beta6 * Sin(12 * xi) * HyperCosine(12 * eta) + _
    beta7 * Sin(14 * xi) * HyperCosine(14 * eta))
etaprime = eta - (beta1 * Cos(2 * xi) * HyperSine(2 * eta) + _
    beta2 * Cos(4 * xi) * HyperSine(4 * eta) + _
    beta3 * Cos(6 * xi) * HyperSine(6 * eta) + _
    beta4 * Cos(8 * xi) * HyperSine(8 * eta) + _
    beta5 * Cos(10 * xi) * HyperSine(10 * eta) + _
    beta6 * Cos(12 * xi) * HyperSine(12 * eta) + _
    beta7 * Cos(14 * xi) * HyperSine(14 * eta))
tauprime = Sin(xiprime) / Sqr(HyperSine(etaprime) ^ 2 + Cos(xiprime) ^ 2)

tau = tauprime
Do Until i = 10
    sigma = HyperSine(e * HyperArcTangent(e * tau / Sqr(1 + tau ^ 2)))
    ftau = tau * Sqr(1 + sigma ^ 2) - sigma * Sqr(1 + tau ^ 2) - tauprime
    If ftau = 0 Then
        Exit Do
    End If
    dtau = (Sqr((1 + sigma ^ 2) * (1 + tau ^ 2)) - sigma * tau) * _
        (1 - e ^ 2) * Sqr(1 + tau ^ 2) / (1 + (1 - e ^ 2) * tau ^ 2)
    tau = tau - ftau / dtau
    i = i + 1
Loop

'Get output
LatTemp = Abs(Atn(tau) * 180 / pi)
If Hem = "S" Then
    LatTemp = -1 * LatTemp
End If

UTMToLatDec = LatTemp

End Function

Public Function UTMToLLDec(UTME As Double, UTMN As Double, UTMDatum As Datum, _
    UTMZone As Integer, Optional Hem As String = "N") As String
'Converts UTM coordinates to decimal lat lon

Dim LatTemp As Double
Dim LonTemp As Double
Dim Meridian As Integer
Dim EastOfCM As Integer
Dim i As Integer

Dim tau As Double
Dim tauprime As Double
Dim ftau As Double
Dim dtau As Double
Dim sigma As Double
Dim xi As Double
Dim xiprime As Double
Dim eta As Double
Dim etaprime As Double

'Set values for a and b (spheroid diameters)
SetSpheroid UTMDatum

'Get other parameters
Meridian = GetMeridian(UTMZone)
LonTemp = UTME - FalseEasting
If LonTemp > 0 Then
    EastOfCM = 1
Else
    EastOfCM = -1
End If
eta = LonTemp / (k0 * AA)

LatTemp = UTMN
If Hem = "S" Then
    LatTemp = 10000000 - LatTemp
End If
xi = LatTemp / (k0 * AA)

'Run the crazy stuff
xiprime = xi - (beta1 * Sin(2 * xi) * HyperCosine(2 * eta) + _
    beta2 * Sin(4 * xi) * HyperCosine(4 * eta) + _
    beta3 * Sin(6 * xi) * HyperCosine(6 * eta) + _
    beta4 * Sin(8 * xi) * HyperCosine(8 * eta) + _
    beta5 * Sin(10 * xi) * HyperCosine(10 * eta) + _
    beta6 * Sin(12 * xi) * HyperCosine(12 * eta) + _
    beta7 * Sin(14 * xi) * HyperCosine(14 * eta))
etaprime = eta - (beta1 * Cos(2 * xi) * HyperSine(2 * eta) + _
    beta2 * Cos(4 * xi) * HyperSine(4 * eta) + _
    beta3 * Cos(6 * xi) * HyperSine(6 * eta) + _
    beta4 * Cos(8 * xi) * HyperSine(8 * eta) + _
    beta5 * Cos(10 * xi) * HyperSine(10 * eta) + _
    beta6 * Cos(12 * xi) * HyperSine(12 * eta) + _
    beta7 * Cos(14 * xi) * HyperSine(14 * eta))
tauprime = Sin(xiprime) / Sqr(HyperSine(etaprime) ^ 2 + Cos(xiprime) ^ 2)

tau = tauprime
Do Until i = 10
    sigma = HyperSine(e * HyperArcTangent(e * tau / Sqr(1 + tau ^ 2)))
    ftau = tau * Sqr(1 + sigma ^ 2) - sigma * Sqr(1 + tau ^ 2) - tauprime
    If ftau = 0 Then
        Exit Do
    End If
    dtau = (Sqr((1 + sigma ^ 2) * (1 + tau ^ 2)) - sigma * tau) * _
        (1 - e ^ 2) * Sqr(1 + tau ^ 2) / (1 + (1 - e ^ 2) * tau ^ 2)
    tau = tau - ftau / dtau
    i = i + 1
Loop

'Get output
LatTemp = Abs(Atn(tau) * 180 / pi)
If Hem = "S" Then
    LatTemp = -1 * LatTemp
End If

LonTemp = Atn(HyperSine(etaprime) / Cos(xiprime))
LonTemp = LonTemp * 180 / pi
LonTemp = LonTemp + Meridian

UTMToLLDec = LatTemp & ", " & LonTemp

End Function

Public Function UTMToLLDMS(UTME As Double, UTMN As Double, UTMDatum As Datum, _
    UTMZone As Integer, Optional Hem As String = "N") As String
'Converts UTM coordinates to lat lon in deg min sec

Dim LatDec As Double
Dim LonDec As Double

LatDec = UTMToLatDec(UTME, UTMN, UTMDatum, UTMZone, Hem)
LonDec = UTMToLonDec(UTME, UTMN, UTMDatum, UTMZone, Hem)

UTMToLLDMS = LLDecToLLDMS(LatDec) & "; " & LLDecToLLDMS(LonDec)

End Function

Public Function UTMToLonDec(UTME As Double, UTMN As Double, UTMDatum As Datum, _
    UTMZone As Integer, Optional Hem As String = "N") As Double
'Retrieves decimal lon from UTM coordinates

Dim LatTemp As Double
Dim LonTemp As Double
Dim Meridian As Integer
Dim EastOfCM As Integer
Dim i As Integer

Dim tau As Double
Dim tauprime As Double
Dim ftau As Double
Dim dtau As Double
Dim sigma As Double
Dim xi As Double
Dim xiprime As Double
Dim eta As Double
Dim etaprime As Double

'Set values for a and b (spheroid diameters)
SetSpheroid UTMDatum

'Get other parameters
Meridian = GetMeridian(UTMZone)
LonTemp = UTME - FalseEasting
If LonTemp > 0 Then
    EastOfCM = 1
Else
    EastOfCM = -1
End If
eta = LonTemp / (k0 * AA)

LatTemp = UTMN
If Hem = "S" Then
    LatTemp = 10000000 - LatTemp
End If
xi = LatTemp / (k0 * AA)

'Run the crazy stuff
xiprime = xi - (beta1 * Sin(2 * xi) * HyperCosine(2 * eta) + _
    beta2 * Sin(4 * xi) * HyperCosine(4 * eta) + _
    beta3 * Sin(6 * xi) * HyperCosine(6 * eta) + _
    beta4 * Sin(8 * xi) * HyperCosine(8 * eta) + _
    beta5 * Sin(10 * xi) * HyperCosine(10 * eta) + _
    beta6 * Sin(12 * xi) * HyperCosine(12 * eta) + _
    beta7 * Sin(14 * xi) * HyperCosine(14 * eta))
etaprime = eta - (beta1 * Cos(2 * xi) * HyperSine(2 * eta) + _
    beta2 * Cos(4 * xi) * HyperSine(4 * eta) + _
    beta3 * Cos(6 * xi) * HyperSine(6 * eta) + _
    beta4 * Cos(8 * xi) * HyperSine(8 * eta) + _
    beta5 * Cos(10 * xi) * HyperSine(10 * eta) + _
    beta6 * Cos(12 * xi) * HyperSine(12 * eta) + _
    beta7 * Cos(14 * xi) * HyperSine(14 * eta))
tauprime = Sin(xiprime) / Sqr(HyperSine(etaprime) ^ 2 + Cos(xiprime) ^ 2)

tau = tauprime
Do Until i = 10
    sigma = HyperSine(e * HyperArcTangent(e * tau / Sqr(1 + tau ^ 2)))
    ftau = tau * Sqr(1 + sigma ^ 2) - sigma * Sqr(1 + tau ^ 2) - tauprime
    If ftau = 0 Then
        Exit Do
    End If
    dtau = (Sqr((1 + sigma ^ 2) * (1 + tau ^ 2)) - sigma * tau) * _
        (1 - e ^ 2) * Sqr(1 + tau ^ 2) / (1 + (1 - e ^ 2) * tau ^ 2)
    tau = tau - ftau / dtau
    i = i + 1
Loop

'Get output
LonTemp = Atn(HyperSine(etaprime) / Cos(xiprime))
LonTemp = LonTemp * 180 / pi
LonTemp = LonTemp + Meridian

UTMToLonDec = LonTemp

End Function

Public Function UTMToUTM(UTME As Double, UTMN As Double, DatumIn As Datum, _
    DatumOut As Datum, ZoneIn As Integer, ZoneOut As Integer, _
    Optional Hem As String = "N") As String
'Converts input UTM coordinates to specified datum/zone

Dim LatDec As Double
Dim LonDec As Double

LatDec = UTMToLatDec(UTME, UTMN, DatumIn, ZoneIn, Hem)
LonDec = UTMToLonDec(UTME, UTMN, DatumIn, ZoneIn, Hem)

UTMToUTM = LLDecToUTM(LatDec, LonDec, DatumOut, ZoneOut)

End Function

Public Function UTMToUTME(UTME As Double, UTMN As Double, DatumIn As Datum, _
    DatumOut As Datum, ZoneIn As Integer, ZoneOut As Integer, _
    Optional Hem As String = "N") As Double
'Retrieves UTM easting in specified datum/zone from input UTM coordinates

Dim LatDec As Double
Dim LonDec As Double

LatDec = UTMToLatDec(UTME, UTMN, DatumIn, ZoneIn, Hem)
LonDec = UTMToLonDec(UTME, UTMN, DatumIn, ZoneIn, Hem)

UTMToUTME = LLDecToUTME(LatDec, LonDec, DatumOut, ZoneOut)

End Function

Public Function UTMToUTMN(UTME As Double, UTMN As Double, DatumIn As Datum, _
    DatumOut As Datum, ZoneIn As Integer, ZoneOut As Integer, _
    Optional Hem As String = "N") As Double
'Retrieves UTM northing in specified datum/zone from input UTM coordinates

Dim LatDec As Double
Dim LonDec As Double

LatDec = UTMToLatDec(UTME, UTMN, DatumIn, ZoneIn, Hem)
LonDec = UTMToLonDec(UTME, UTMN, DatumIn, ZoneIn, Hem)

UTMToUTMN = LLDecToUTMN(LatDec, LonDec, DatumOut, ZoneOut)

End Function
