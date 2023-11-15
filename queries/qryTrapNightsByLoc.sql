SELECT
  Visits.LocationID,
  Sum(
    Round(
      [Visits].[ActiveEnd] - [Visits].[ActiveStart],
      3
    )
  ) AS [Trap Nights]
FROM
  Visits
GROUP BY
  Visits.LocationID;
