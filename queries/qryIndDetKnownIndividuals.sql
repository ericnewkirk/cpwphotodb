SELECT
  IndependentDetections.IndDetectionID,
  Count(IndividualDetections.ID) AS KnownIndividuals
FROM
  IndependentDetections
  INNER JOIN IndividualDetections ON IndependentDetections.IndDetectionID = IndividualDetections.IndDetectionID
WHERE
  (
    (
      (
        IndividualDetections.IndividualID
      )> 0
    )
  )
GROUP BY
  IndependentDetections.IndDetectionID;
