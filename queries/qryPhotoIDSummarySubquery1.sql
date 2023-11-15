SELECT
  Photos.VisitID,
  Photos.ImageID,
  Sum(
    IIf(
      [Detections].[DetectionID] Is Null,
      1, 0
    )
  ) AS NoIDs,
  Sum(
    IIf([Detections].[StatusID] = 1, 1, 0)
  ) AS PendingIDs,
  Sum(
    IIf([Detections].[StatusID] = 2, 1, 0)
  ) AS VerifiedIDs
FROM
  Photos
  LEFT JOIN Detections ON Photos.ImageID = Detections.ImageID
GROUP BY
  Photos.VisitID,
  Photos.ImageID;
