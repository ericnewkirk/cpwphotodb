SELECT
  [Individual] & ': ' & [Session] AS Label,
  Max(
    IIf([Occasion] = 1, 1, 0)
  ) AS Occasion1,
  Max(
    IIf([Occasion] = 2, 1, 0)
  ) AS Occasion2,
  Max(
    IIf([Occasion] = 3, 1, 0)
  ) AS Occasion3,
  Max(
    IIf([Occasion] = 4, 1, 0)
  ) AS Occasion4,
  Max(
    IIf([Session] = 'MolasPass1', 1, 0)
  ) AS Grou p1,
  Max(
    IIf([Session] = 'Telluride1', 1, 0)
  ) AS Group2
FROM
  SECRQuery
GROUP BY
  [Individual] & ': ' & [Session];
