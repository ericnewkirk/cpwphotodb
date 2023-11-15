SELECT
  CameraLocations.LocationID,
  CameraLocations.LocationName,
  StudyAreas.StudyAreaID,
  StudyAreas.StudyAreaName,
  CameraLocations.LatitudeDD AS Latitude,
  CameraLocations.LongitudeDD AS Longitude,
  CameraLocations.AccessNotes
FROM
  StudyAreas
  INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID
WHERE
  (
    (
      (CameraLocations.LatitudeDD) Is Not Null
    )
    AND (
      (CameraLocations.LongitudeDD) Is Not Null
    )
  )
ORDER BY
  CameraLocations.LocationName;
