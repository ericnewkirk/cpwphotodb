DELETE DISTINCTROW NewDetections.*
FROM
  Detections
  INNER JOIN NewDetections ON (
    Detections.ImageID = NewDetections.ImageID
  )
  AND (
    Detections.ObsID = NewDetections.ObsID
  );
