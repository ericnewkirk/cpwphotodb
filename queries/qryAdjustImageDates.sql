SELECT
  Visits.VisitID,
  (
    SELECT
      Min(ImageDate)
    FROM
      Photos
    WHERE
      (
        (
          ([Photos].[VisitID])= [Visits].[VisitID]
        )
      )
  ) AS FirstPhotoDateTime
FROM
  Visits;
