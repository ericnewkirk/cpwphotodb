SELECT
  [StudyAreas].[StudyAreaAbbr] & " - " & [CameraLocations].[LocationName] AS Location,
  CameraLocations.StudyAreaID,
  SetVisits.VisitDate,
  CameraLocations.UTM_E,
  CameraLocations.UTM_N,
  CameraLocations.AccessNotes,
  SetVisits.LocationID
FROM
  (
    (
      StudyAreas
      INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID
    )
    INNER JOIN Visits AS SetVisits ON CameraLocations.LocationID = SetVisits.LocationID
  )
  LEFT JOIN (
    SELECT
      SetVisitID
    FROM
      Visits
    WHERE
      (
        (
          (VisitTypeID)= 2
        )
      )
  ) AS PullVisits ON SetVisits.VisitID = PullVisits.SetVisitID
WHERE
  (
    (
      (SetVisits.VisitTypeID)= 3
    )
    AND (
      (PullVisits.SetVisitID) Is Null
    )
  );
