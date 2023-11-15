SELECT
  SetVisits.LocationID,
  SetVisits.VisitDate AS SetVisitDate,
  Photos.ImageID,
  Photos.ImageDate,
  Photos.VisitID,
  Detections.SpeciesID,
  Detections.StatusID,
  Species.GroupID
FROM
  Species
  INNER JOIN (
    (
      (
        Visits AS SetVisits
        INNER JOIN Visits ON SetVisits.VisitID = Visits.SetVisitID
      )
      INNER JOIN Photos ON Visits.VisitID = Photos.VisitID
    )
    INNER JOIN Detections ON Photos.ImageID = Detections.ImageID
  ) ON Species.SpeciesID = Detections.SpeciesID
WHERE
  (
    (
      (Detections.SpeciesID)> 0
    )
    AND (
      (Detections.StatusID)< 3
    )
  );
