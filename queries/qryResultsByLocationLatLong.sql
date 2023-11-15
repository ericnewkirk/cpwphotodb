SELECT
  First(StudyAreas.StudyAreaID) AS StudyAreaID,
  First(StudyAreas.StudyAreaName) AS StudyAreaName,
  CameraLocations.LocationID,
  First(CameraLocations.LocationName) AS LocationName,
  Species.SpeciesID,
  First(Species.CommonName) AS CommonName,
  Count(Photos.ImageID) AS Photos,
  First(CameraLocations.LatitudeDD) AS Latitude,
  First(CameraLocations.LongitudeDD) AS Longitude
FROM
  (
    (
      StudyAreas
      INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID
    )
    INNER JOIN Visits ON CameraLocations.LocationID = Visits.LocationID
  )
  INNER JOIN (
    (
      Photos
      INNER JOIN qryPhotoSpecies ON Photos.ImageID = qryPhotoSpecies.ImageID
    )
    INNER JOIN Species ON qryPhotoSpecies.SpeciesID = Species.SpeciesID
  ) ON Visits.VisitID = Photos.VisitID
WHERE
  (
    (
      (CameraLocations.LatitudeDD) Is Not Null
    )
    AND (
      (CameraLocations.LongitudeDD) Is Not Null
    )
  )
GROUP BY
  CameraLocations.LocationID,
  Species.SpeciesID
ORDER BY
  First(StudyAreas.StudyAreaName),
  First(CameraLocations.LocationName),
  First(Species.CommonName);
