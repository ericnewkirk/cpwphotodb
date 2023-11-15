SELECT
  StudyAreas.StudyAreaName,
  StudyAreas.StudyAreaID,
  CameraLocations.LocationName,
  CameraLocations.LocationID,
  CameraLocations.UTM_E,
  CameraLocations.UTM_N,
  CameraLocations.UTMZone,
  CameraLocations.LatitudeDD,
  CameraLocations.LongitudeDD,
  Year([SetVisits].[VisitDate]) AS FieldSeason,
  Photos.FileName,
  Visits.VisitID,
  Photos.ImageID AS ImgID,
  Photos.ImageNum,
  Photos.ImageDate,
  Photos.Highlight,
  [Photos].[FilePath] & [Photos].[FileName] AS ImgPath,
  qrySpPivotPending.*
FROM
  (
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
    INNER JOIN Photos ON Visits.VisitID = Photos.VisitID
  )
  LEFT JOIN qrySpPivotPending ON Photos.ImageID = qrySpPivotPending.ImageID
ORDER BY
  Photos.ImageID;
