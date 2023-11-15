SELECT
  Year([SetVisits].[VisitDate]) AS FieldSeason,
  Visits.LocationID,
  Photos.ImageID,
  Photos.FileName,
  [FilePath] & [FileName] AS ImgPath,
  Photos.Highlight,
  Photos.ImageDate
FROM
  (
    Visits
    INNER JOIN Visits AS SetVisits ON Visits.SetVisitID = SetVisits.VisitID
  )
  INNER JOIN Photos ON Visits.VisitID = Photos.VisitID
WHERE
  (
    (
      (Photos.Compare)= True
    )
  )
ORDER BY
  Photos.ImageID;
