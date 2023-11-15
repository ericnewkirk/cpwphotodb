SELECT
  IndependentDetections.ModifiedStart,
  IndependentDetections.LocationID,
  IndividualDetections.IndividualID
FROM
  IndependentDetections
  INNER JOIN IndividualDetections ON IndependentDetections.IndDetectionID = IndividualDetections.IndDetectionID;
