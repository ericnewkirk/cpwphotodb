SELECT
  CameraLocations.StudyAreaID,
  CameraLocations.LocationID,
  (
    SELECT
      Min(SetVisits.VisitDate)
    FROM
      (
        Visits AS SetVisits
        INNER JOIN Visits ON SetVisits.VisitID = Visits.SetVisitID
      )
      INNER JOIN Photos ON Visits.VisitID = Photos.VisitID
    WHERE
      (
        (
          (SetVisits.LocationID)= [CameraLocations].[LocationID]
        )
        AND (
          (Photos.ImageDate)= [qryIndDetRecSource].[ModifiedStart]
        )
      )
  ) AS FilterDate,
  qryIndDetRecSource.IndDetectionID,
  qryIndDetRecSource.SpeciesID,
  qryIndDetRecSource.ModifiedStart
FROM
  CameraLocations
  INNER JOIN qryIndDetRecSource ON CameraLocations.LocationID = qryIndDetRecSource.LocationID
WHERE
  (
    (
      (qryIndDetRecSource.Deleted)= False
    )
  );
