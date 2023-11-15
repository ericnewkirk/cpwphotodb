SELECT
  Left(
    Replace(
      [qryCRSubquery].[StudyAreaName],
      ' ', ''
    ),
    12
  )& DCount(
    '*', 'Visits', 'VisitTypeID=3 And LocationID=' & [qryCRSubquery].[LocationID] & ' And VisitDate<=#' & [qryCRSubquery].[CameraSetDate] & '#'
  ) AS [Session],
  qryCRSubquery.IndividualID,
  SECROccNumber(
    qryCRSubquery.OccasionStart, 14,
    qryCRSubquery.ModifiedStart
  ) AS Occasion,
  qryCRSubquery.LocationID,
  '# ' & qryCRSubquery.StudyAreaName AS StudyArea,
  qryCRSubquery.LocationName AS Location,
  qryCRSubquery.IndividualName AS Individual
FROM
  qryCRSubquery
WHERE
  (
    (
      (qryCRSubquery.ModifiedStart) Between qryCRSubquery.OccasionStart
      And DateAdd(
        's',
        -1,
        DateAdd(
          'd', 56, qryCRSubquery.OccasionStart
        )
      )
    )
    AND (
      (qryCRSubquery.SpeciesID)= 65
    )
    AND (
      (qryCRSubquery.IndividualID)> 0
    )
  );
