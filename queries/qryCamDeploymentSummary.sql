SELECT
  CameraLocations.LocationID,
  CameraLocations.LocationName,
  CameraLocations.StudyAreaID,
  Year([Visits].[VisitDate]) AS CamYear,
  Visits.VisitDate AS SetDate,
  PullVisits.VisitDate AS PullDate,
  DateDiff("d", [SetDate], [PullDate]) AS DaysDeployed,
  CameraLocations.UTM_E,
  CameraLocations.UTM_N,
  DSum(
    "PhotoCount", "qryVisitsRecSource",
    "SetVisitID=" & [Visits].[VisitID]
  ) AS TotalPhotos,
  PullVisits.Comments,
  Round(
    DSum(
      "ActiveEnd - ActiveStart", "Visits",
      "SetVisitID=" & [Visits].[VisitID]
    ),
    3
  ) AS TrapNights
FROM
  CameraLocations
  INNER JOIN (
    Visits
    LEFT JOIN (
      SELECT
        *
      FROM
        Visits
      WHERE
        (
          (
            (VisitTypeID)= 2
          )
        )
    ) AS PullVisits ON Visits.VisitID = PullVisits.SetVisitID
  ) ON CameraLocations.LocationID = Visits.LocationID
WHERE
  (
    (
      (Visits.VisitTypeID)= 3
    )
  )
ORDER BY
  CameraLocations.LocationID;
