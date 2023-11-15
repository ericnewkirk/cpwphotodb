SELECT
  Visits.VisitID,
  Visits.LocationID,
  Visits.VisitTypeID,
  Visits.VisitDate,
  Visits.Comments,
  (
    SELECT
      Count(ImageID)
    FROM
      Photos
    WHERE
      (
        (
          ([Photos].[VisitID])= [Visits].[VisitID]
        )
      )
  ) AS PhotoCount,
  Visits.SetVisitID
FROM
  Visits;
