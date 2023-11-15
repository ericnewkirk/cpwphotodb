UPDATE
  Detections
  INNER JOIN Detections AS D ON (Detections.ImageID = D.ImageID)
  AND (
    Nz(Detections.DetailID, 0)= Nz(D.DetailID, 0)
  )
  AND (
    Detections.Individuals = D.Individuals
  )
  AND (
    Detections.SpeciesID = D.SpeciesID
  )
SET
  Detections.StatusID = 3
WHERE
  (
    (
      (D.StatusID)= 3
    )
    AND (
      (Detections.StatusID)= 1
    )
  );
