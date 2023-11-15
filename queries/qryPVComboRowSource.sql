SELECT
  CameraLocations.StudyAreaID,
  CameraLocations.LocationID,
  Year([SetVisits].[VisitDate]) AS FieldSeason,
  Visits.VisitID,
  Photos.ImageID AS ImgID,
  Photos.ImageDate,
  Photos.Highlight,
  Photos.ObsCount,
  Photos.Verified,
  Photos.Pending,
  Photos.MultiSp,
  Photos.NotNone
FROM
  (
    (
      CameraLocations
      INNER JOIN Visits ON CameraLocations.LocationID = Visits.LocationID
    )
    INNER JOIN Photos ON Visits.VisitID = Photos.VisitID
  )
  INNER JOIN Visits AS SetVisits ON Visits.SetVisitID = SetVisits.VisitID;
