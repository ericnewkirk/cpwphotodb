SELECT
  Visits.LocationID,
  Detections.SpeciesID,
  SetVisits.VisitDate AS SetVisitDate,
  Photos.ImageDate,
  Photos.ImageNum,
  True AS Include,
  IIf(
    Count([Detections].[DetectionID])> 1
    Or Max([Detections].[StatusID])> 1,
    2,
    1
  ) AS Obs INTO Activity IN 'C:\Users\NewkirkE\Desktop\CPWPhotoWarehouse_v4AP.accdb'
FROM
  (
    Visits
    INNER JOIN Visits AS SetVisits ON Visits.SetVisitID = SetVisits.VisitID
  )
  INNER JOIN (
    Photos
    INNER JOIN Detections ON Photos.ImageID = Detections.ImageID
  ) ON Visits.VisitID = Photos.VisitID
WHERE
  (
    (
      (Detections.SpeciesID)> 0
    )
    AND (
      (Detections.StatusID)< 3
    )
  )
GROUP BY
  Visits.LocationID,
  Detections.SpeciesID,
  SetVisits.VisitDate,
  Photos.ImageDate,
  Photos.ImageNum;
