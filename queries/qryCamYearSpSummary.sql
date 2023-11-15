SELECT
  qryCamStudyAreaSpSummary.CameraYear,
  qryCamStudyAreaSpSummary.CommonName,
  Sum(
    qryCamStudyAreaSpSummary.Images
  ) AS Images
FROM
  qryCamStudyAreaSpSummary
GROUP BY
  qryCamStudyAreaSpSummary.CameraYear,
  qryCamStudyAreaSpSummary.CommonName;
