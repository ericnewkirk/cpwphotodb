SELECT
  DISTINCT Detections.ImageID,
  Detections.SpeciesID
FROM
  Detections
WHERE
  (
    (
      (Detections.StatusID)< 3
    )
  );
