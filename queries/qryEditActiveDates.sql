SELECT
  Visits.VisitID,
  StudyAreas.StudyAreaName,
  CameraLocations.LocationName,
  lkupVisitTypes.VisitType,
  Visits.VisitDate,
  Visits.ActiveStart,
  Visits.ActiveEnd,
  DMin(
    "ImageDate", "Photos", "VisitID=" & [VisitID]
  ) AS FirstPhoto,
  DMax(
    "ImageDate", "Photos", "VisitID=" & [VisitID]
  ) AS LastPhoto,
  Visits.Comments
FROM
  (
    StudyAreas
    INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID
  )
  INNER JOIN (
    lkupVisitTypes
    INNER JOIN Visits ON lkupVisitTypes.ID = Visits.VisitTypeID
  ) ON CameraLocations.LocationID = Visits.LocationID
ORDER BY
  StudyAreas.StudyAreaName,
  CameraLocations.LocationName,
  Visits.VisitDate;
