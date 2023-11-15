SELECT
  IndependentDetections.IndDetectionID,
  Sum(IndDetGroups.Individuals) AS TotalIndividuals,
  Sum(
    IIf([AgeClassID] = 1, 1, 0)
  ) AS TotalAdults,
  Sum(
    IIf([AgeClassID] = 2, 1, 0)
  ) AS TotalJuveniles,
  Sum(
    IIf([AgeClassID] = 3, 1, 0)
  ) AS TotalSubadults,
  Sum(
    IIf([GenderID] = 1, 1, 0)
  ) AS TotalFemales,
  Sum(
    IIf([GenderID] = 2, 1, 0)
  ) AS TotalMales
FROM
  IndependentDetections
  INNER JOIN IndDetGroups ON IndependentDetections.IndDetectionID = IndDetGroups.IndDetectionID
WHERE
  (
    (
      (IndependentDetections.Deleted)= False
    )
  )
GROUP BY
  IndependentDetections.IndDetectionID;
