UPDATE
  (
    Photos
    INNER JOIN Detections ON Photos.ImageID = Detections.ImageID
  )
  INNER JOIN Detections AS D1 ON (Detections.ImageID = D1.ImageID)
  AND (
    Detections.Individuals = D1.Individuals
  )
  AND (
    Detections.DetailID = D1.DetailID
  )
  AND (
    Detections.SpeciesID = D1.SpeciesID
  )
  AND (
    Detections.DetectionID <> D1.DetectionID
  )
SET
  Detections.StatusID = 2
WHERE
  (
    (
      (Photos.NeedsUpdate)= True
    )
    AND (
      (Detections.StatusID)= 1
    )
    AND (
      (D1.StatusID)< 3
    )
  );
