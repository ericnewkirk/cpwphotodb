SELECT
  CameraLocations.LocationID,
  CameraLocations.LocationName,
  StudyAreas.StudyAreaID,
  StudyAreas.StudyAreaName,
  CameraLocations.UTM_E,
  CameraLocations.UTM_N,
  CameraLocations.UTMZone,
  CameraLocations.UTMDatum,
  CameraLocations.UTMHemisphere,
  CameraLocations.AccessNotes
FROM
  StudyAreas
  INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID
WHERE
  (
    (
      (CameraLocations.UTM_E) Is Not Null
    )
    AND (
      (CameraLocations.UTM_N) Is Not Null
    )
  )
ORDER BY
  CameraLocations.LocationName;
