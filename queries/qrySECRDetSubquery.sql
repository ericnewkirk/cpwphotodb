SELECT
  CameraLocations.LocationID,
  CameraLocations.UTM_E,
  CameraLocations.UTM_N,
  SECREffortString(
    CameraLocations.LocationID, qryCRSubquery.OccasionStart,
    14, 4, 12
  ) AS Effort,
  '#' & CameraLocations.LocationName AS Location,
  Left(
    Replace(
      [StudyAreas].[StudyAreaName], ' ',
      ''
    ),
    12
  )& DCount(
    '*', 'Visits', 'VisitTypeID=3 And LocationID=' & [qryCRSubquery].[LocationID] & ' And VisitDate<=#' & [qryCRSubquery].[CameraSetDate] & '#'
  ) AS [Session]
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
  );
