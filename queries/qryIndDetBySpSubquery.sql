SELECT
  IndependentDetections.SpeciesID,
  Count(
    IndependentDetections.IndDetectionID
  ) AS Detectons,
  Sum(
    qryIndDetGroupsSummary.TotalIndividuals
  ) AS TotalIndividuals,
  Sum(
    qryIndDetGroupsSummary.TotalAdults
  ) AS TotalAdults,
  Sum(
    qryIndDetGroupsSummary.TotalJuveniles
  ) AS TotalJuveniles,
  Sum(
    qryIndDetGroupsSummary.TotalSubadults
  ) AS TotalSubadults,
  Sum(
    qryIndDetGroupsSummary.TotalFemales
  ) AS TotalFemales,
  Sum(
    qryIndDetGroupsSummary.TotalMales
  ) AS TotalMales,
  Nz(
    Sum(
      [qryIndDetKnownIndividuals].[KnownIndividuals]
    ),
    0
  ) AS TotalKnownInd
FROM
  (
    IndependentDetections
    INNER JOIN qryIndDetGroupsSummary ON IndependentDetections.IndDetectionID = qryIndDetGroupsSummary.IndDetectionID
  )
  LEFT JOIN qryIndDetKnownIndividuals ON IndependentDetections.IndDetectionID = qryIndDetKnownIndividuals.IndDetectionID
GROUP BY
  IndependentDetections.SpeciesID;
