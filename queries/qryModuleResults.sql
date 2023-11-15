SELECT
  Detections.DetectionID,
  Detections.StatusID
FROM
  Detections
  INNER JOIN NewDetections ON Detections.DetectionID = NewDetections.DetectionID;
