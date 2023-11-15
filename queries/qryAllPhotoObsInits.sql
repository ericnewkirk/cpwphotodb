SELECT
  Visits.VisitID,
  Observers.Initials
FROM
  Visits
  INNER JOIN (
    Observers
    INNER JOIN (
      Photos
      INNER JOIN Detections ON Photos.ImageID = Detections.ImageID
    ) ON Observers.ObserverID = Detections.ObsID
  ) ON Visits.VisitID = Photos.VisitID
GROUP BY
  Visits.VisitID,
  Observers.Initials;
