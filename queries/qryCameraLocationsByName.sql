SELECT
  CameraLocations.LocationID,
  [StudyAreas].[StudyAreaAbbr] & " - " & [CameraLocations].[LocationName] AS LocName,
  CameraLocations.AccessNotes
FROM
  StudyAreas
  INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID
ORDER BY
  [StudyAreas].[StudyAreaAbbr] & " - " & [CameraLocations].[LocationName];
