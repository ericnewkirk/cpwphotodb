DELETE DISTINCTROW NewDetections.*
FROM
  Detections
  INNER JOIN NewDetections ON Detections.DetectionID = NewDetections.DetectionID;
