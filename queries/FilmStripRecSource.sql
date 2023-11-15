SELECT
  CameraLocations.LocationID,
  CameraLocations.StudyAreaID,
  CameraLocations.UTM_E,
  CameraLocations.UTM_N,
  CameraLocations.UTMZone,
  Year([SetVisits].[VisitDate]) AS FieldSeason,
  Photos.FileName,
  Visits.VisitID,
  Photos.ImageID AS ImgID,
  Photos.ImageNum,
  Photos.ImageDate,
  Photos.Highlight,
  [Photos].[FilePath] & [Photos].[FileName] AS ImgPath
FROM
  (
    (
      CameraLocations
      INNER JOIN Visits ON CameraLocations.LocationID = Visits.LocationID
    )
    INNER JOIN Visits AS SetVisits ON Visits.SetVisitID = SetVisits.VisitID
  )
  INNER JOIN Photos ON Visits.VisitID = Photos.VisitID
ORDER BY
  Photos.ImageDate,
  Photos.ImageID;
