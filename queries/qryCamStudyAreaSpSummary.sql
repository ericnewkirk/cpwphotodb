SELECT
  Year([SetVisits].[VisitDate]) AS CameraYear,
  CameraLocations.StudyAreaID,
  Species.CommonName,
  Count(qryPhotoSpecies.ImageID) AS Images
FROM
  Species
  INNER JOIN (
    (
      CameraLocations
      INNER JOIN (
        Visits AS SetVisits
        INNER JOIN (
          Visits
          INNER JOIN Photos ON Visits.VisitID = Photos.VisitID
        ) ON SetVisits.VisitID = Visits.SetVisitID
      ) ON CameraLocations.LocationID = SetVisits.LocationID
    )
    INNER JOIN qryPhotoSpecies ON Photos.ImageID = qryPhotoSpecies.ImageID
  ) ON Species.SpeciesID = qryPhotoSpecies.SpeciesID
GROUP BY
  Year([SetVisits].[VisitDate]),
  CameraLocations.StudyAreaID,
  Species.CommonName;
