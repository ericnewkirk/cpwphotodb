SELECT
  qryPhotoIDSummarySubquery1.VisitID,
  Count(
    qryPhotoIDSummarySubquery1.ImageID
  ) AS Photos,
  Sum(
    qryPhotoIDSummarySubquery1.NoIDs
  ) AS NoID,
  Sum(
    IIf(
      [qryPhotoIDSummarySubquery1].[PendingIDs] + [qryPhotoIDSummarySubquery1].[VerifiedIDs] > 0,
      1, 0
    )
  ) AS ID,
  Sum(
    IIf(
      [qryPhotoIDSummarySubquery1].[VerifiedIDs] > 0,
      1, 0
    )
  ) AS VerifiedID
FROM
  qryPhotoIDSummarySubquery1
GROUP BY
  qryPhotoIDSummarySubquery1.VisitID;
