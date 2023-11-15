SELECT
  Photos.ImageID,
  [Photos].[FilePath] & [Photos].[FileName] AS ImgPath,
  Photos.FileName,
  Photos.ImageDate,
  qryIndDetRecSource.IndDetectionID
FROM
  (
    Visits
    INNER JOIN (
      Photos
      INNER JOIN (
        SELECT
          Detections.ImageID,
          Detections.SpeciesID
        FROM
          Detections
        WHERE
          (
            (
              (Detections.StatusID)< 3
            )
            And (
              (Detections.SpeciesID)> 0
            )
          )
        GROUP BY
          Detections.ImageID,
          Detections.SpeciesID
      ) AS Q ON Photos.ImageID = Q.ImageID
    ) ON Visits.VisitID = Photos.VisitID
  )
  INNER JOIN qryIndDetRecSource ON (
    (
      Photos.ImageDate < qryIndDetRecSource.NextDetection
    )
    OR (
      qryIndDetRecSource.NextDetection Is Null
    )
  )
  AND (
    Photos.ImageDate >= qryIndDetRecSource.ModifiedStart
  )
  AND (
    Q.SpeciesID = qryIndDetRecSource.SpeciesID
  )
  AND (
    Visits.LocationID = qryIndDetRecSource.LocationID
  );
