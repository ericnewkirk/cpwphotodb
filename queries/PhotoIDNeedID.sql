SELECT
  PhotoIDRecSource.*
FROM
  PhotoIDRecSource
  LEFT JOIN (
    SELECT
      Detections.ImageID
    FROM
      Detections
      INNER JOIN CurrentObserver ON Detections.ObsID = CurrentObserver.ObserverID
    UNION
    SELECT
      Detections.ImageID
    FROM
      Detections
    WHERE
      (
        (
          (StatusID)= 2
        )
      )
  ) AS HaveID ON PhotoIDRecSource.ImageID = HaveID.ImageID
WHERE
  (
    (
      (HaveID.ImageID) Is Null
    )
  )
ORDER BY
  PhotoIDRecSource.ImageID;
