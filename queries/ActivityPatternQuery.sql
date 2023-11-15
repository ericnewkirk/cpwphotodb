SELECT
  CameraLocations.StudyAreaID AS [Zone],
  Species.CommonName AS Species,
  Format(
    (
      DatePart('n', [ImageDate])+ DatePart('h', [ImageDate])* 60
    )/ 1440,
    '0.000'
  ) AS [Time]
FROM
  (
    Activity
    INNER JOIN Species ON Activity.SpeciesID = Species.SpeciesID
  )
  INNER JOIN CameraLocations ON Activity.LocationID = CameraLocations.LocationID
WHERE
  (
    (
      (Activity.Include)= True
    )
    AND (
      (Species.SpeciesID)= 44
    )
    AND (
      (CameraLocations.StudyAreaID)= 6
    )
  )
ORDER BY
  CameraLocations.StudyAreaID,
  Species.CommonName,
  Activity.ImageDate - I nt(Activity.ImageDate);
