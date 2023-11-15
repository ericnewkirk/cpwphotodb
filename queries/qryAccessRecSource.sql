SELECT
  [StudyAreas].[StudyAreaAbbr] & " - " & [CameraLocations].[LocationName] AS Location,
  CameraLocations.AccessNotes AS Access
FROM
  StudyAreas
  INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID
ORDER BY
  CameraLocations.LocationID;
