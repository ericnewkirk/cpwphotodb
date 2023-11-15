SELECT
  qryCamDeploymentSummary.CamYear,
  qryCamDeploymentSummary.StudyAreaID,
  Min(
    qryCamDeploymentSummary.SetDate
  ) AS MinSet,
  Max(
    qryCamDeploymentSummary.SetDate
  ) AS MaxSet,
  Min(
    qryCamDeploymentSummary.PullDate
  ) AS MinPull,
  Max(
    qryCamDeploymentSummary.PullDate
  ) AS MaxPull,
  Round(
    Avg([DaysDeployed]),
    3
  ) AS AvgDays,
  Min(
    qryCamDeploymentSummary.DaysDeployed
  ) AS MinDays,
  Max(
    qryCamDeploymentSummary.DaysDeployed
  ) AS MaxDays,
  Round(
    Avg(
      [qryCamDeploymentSummary].[TotalPhotos]
    ),
    3
  ) AS AvgPhotos,
  Sum(
    qryCamDeploymentSummary.TotalPhotos
  ) AS TotalPhotos,
  Count(
    qryCamDeploymentSummary.LocationID
  ) AS CamerasSet,
  Count(
    qryCamDeploymentSummary.PullDate
  ) AS CamerasRetrieved,
  Round(
    Sum([TrapNights]),
    3
  ) AS TtlEffort,
  Round(
    Min([TrapNights]),
    3
  ) AS MinEffort,
  Round(
    Max([TrapNights]),
    3
  ) AS MaxEffort,
  Round(
    Avg([TrapNights]),
    3
  ) AS AvgEffort
FROM
  qryCamDeploymentSummary
GROUP BY
  qryCamDeploymentSummary.CamYear,
  qryCamDeploymentSummary.StudyAreaID;
