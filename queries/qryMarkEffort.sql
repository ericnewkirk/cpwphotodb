SELECT
  qrySECRDetSubquery.Session,
  Sum(
    SECRSplitEffortString(qrySECRDetSubquery.Effort, 1)
  ) AS GroupEffort1,
  Sum(
    SECRSplitEffortString(qrySECRDetSubquery.Effort, 2)
  ) AS GroupEffort2,
  Sum(
    SECRSplitEffortString(qrySECRDetSubquery.Effort, 3)
  ) AS GroupEffort3,
  Sum(
    SECRSplitEffortString(qrySECRDetSubquery.Effort, 4)
  ) AS GroupEffort4
FROM
  qrySECRDetSubquery
GROUP BY
  qrySECRDetSubquery.Session;
