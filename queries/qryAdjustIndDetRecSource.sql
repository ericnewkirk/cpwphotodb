SELECT
  Photos.ImageID,
  Photos.FileName,
  [Photos].[FilePath] & [Photos].[FileName] AS ImgPath,
  qryIndDetRecSource.IndDetectionID,
  Photos.ImageDate,
  IIf(
    [Photos].[ImageDate] = [IndependentDetections].[ModifiedStart],
    "Current Start", ""
  ) AS [Current]
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
          )
        GROUP BY
          Detections.ImageID,
          Detections.SpeciesID
      ) AS Q ON Photos.ImageID = Q.ImageID
    ) ON Visits.VisitID = Photos.VisitID
  )
  INNER JOIN qryIndDetRecSource ON (
    Visits.LocationID = qryIndDetRecSource.LocationID
  )
  AND (
    Q.SpeciesID = qryIndDetRecSource.SpeciesID
  )
  AND (
    Photos.ImageDate > Nz(
      qryIndDetRecSource.PrevDetection,
      0
    )
  )
  AND (
    (
      Photos.ImageDate < qryIndDetRecSource.NextDetection
    )
    OR (
      qryIndDetRecSource.NextDetection Is Null
    )
  );
