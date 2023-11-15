SELECT
  qryCamDeploymentSummary.CamYear,
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
  Avg(
    qryCamDeploymentSummary.DaysDeployed
  ) AS AvgDays,
  Min(
    qryCamDeploymentSummary.DaysDeployed
  ) AS MinDays,
  Max(
    qryCamDeploymentSummary.DaysDeployed
  ) AS MaxDays,
  Avg(
    qryCamDeploymentSummary.TotalPhotos
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
  Min(
    qryCamDeploymentSummary.SetDate
  ) AS FilterDate,
  Sum(
    qryCamDeploymentSummary.TrapNights
  ) AS TtlEffort,
  Min(
    qryCamDeploymentSummary.TrapNights
  ) AS MinEffort,
  Max(
    qryCamDeploymentSummary.TrapNights
  ) AS MaxEffort,
  Avg(
    qryCamDeploymentSummary.TrapNights
  ) AS AvgEffort
FROM
  qryCamDeploymentSummary
GROUP BY
  qryCamDeploymentSummary.CamYear;
