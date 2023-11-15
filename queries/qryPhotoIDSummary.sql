SELECT
  StudyAreas.StudyAreaName,
  First(CameraLocations.LocationName) AS LocationName,
  CameraLocations.LocationID,
  First(
    Year([SetVisits].[VisitDate])
  ) AS CamYear,
  Visits.VisitID,
  First(Visits.VisitDate) AS VisitDate,
  Sum(
    qryPhotoIDSummarySubquery2.Photos
  ) AS TotalPhotos,
  Sum(
    qryPhotoIDSummarySubquery2.NoID
  ) AS PhotosNoID,
  Sum(qryPhotoIDSummarySubquery2.ID) AS PhotosID,
  Sum(
    qryPhotoIDSummarySubquery2.VerifiedID
  ) AS PhotosVerifiedID
FROM
  (
    (
      (
        StudyAreas
        INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID
      )
      INNER JOIN Visits AS SetVisits ON CameraLocations.LocationID = SetVisits.LocationID
    )
    INNER JOIN Visits ON SetVisits.VisitID = Visits.SetVisitID
  )
  INNER JOIN qryPhotoIDSummarySubquery2 ON Visits.VisitID = qryPhotoIDSummarySubquery2.VisitID
GROUP BY
  StudyAreas.StudyAreaName,
  CameraLocations.LocationID,
  Visits.VisitID
ORDER BY
  StudyAreas.StudyAreaName,
  First(CameraLocations.LocationName),
  First(Visits.VisitDate);
