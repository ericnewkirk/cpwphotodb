SELECT
  [StudyAreas].[StudyAreaAbbr] & ' - ' & [CameraLocations].[LocationName] AS Location,
  GetOccasion(
    [CameraLocations].[LocationID],
    DateAdd('d', 0, [StartDateTime]),
    DateAdd(
      's',
      -1,
      DateAdd('d', 10, [StartDateTime])
    ),
    3,
    False,
    False
  ) AS Occasion1,
  GetOccasion(
    [CameraLocations].[LocationID],
    DateAdd('d', 10, [StartDateTime]),
    DateAdd(
      's',
      -1,
      DateAdd('d', 20, [StartDateTime])
    ),
    3,
    False,
    False
  ) AS Occasion2,
  1 AS Grou p1,
  CameraLocations.StudyAreaID AS StudyArea,
  Year([Visits].[VisitDate]) AS SurveyYear,
  Nz(CameraLocations.LatitudeDD, '.') AS CameraLat,
  Nz(
    CameraLocations.LongitudeDD, '.'
  ) AS CameraLong,
  DateAdd(
    'yyyy',
    Year(
      Nz(
        (
          SELECT
            Min(V.ActiveStart)
          FROM
            Visits AS V
          WHERE
            (
              (
                (V.SetVisitID)= [Visits].[VisitID]
              )
            )
        ),
        [Visits].[VisitDate]
      )
    )-2014 + IIf(
      DatePart(
        'y',
        Nz(
          (
            SELECT
              Min(V.ActiveStart)
            FROM
              Visits AS V
            WHERE
              (
                (
                  (V.SetVisitID)= [Visits].[VisitID]
                )
              )
          ),
          [Visits].[VisitDate]
        )
      )> 355,
      1,
      0
    ),
    #12/1/2014 12:00:00 PM#
  ) AS StartDateTime
FROM
  (
    StudyAreas
    INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID
  )
  INNER JOIN Visits ON CameraLocations.LocationID = Visits.LocationID
WHERE
  (
    (
      (Visits.VisitTypeID)= 3
    )
  )
ORDER BY
  CameraLocations.LocationID,
  Visits.VisitDate;
