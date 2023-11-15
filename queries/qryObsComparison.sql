SELECT
  Q.Initials,
  Q.TtlDet AS TotalDetections,
  Q.VDet /(Q.TtlDet - Q.PDet) AS PctVerified,
  Q.DDet /(Q.TtlDet - Q.PDet) AS PctDeleted,
  Q.DelSp /(Q.TtlDet - Q.PDet) AS WrongSp,
  Q.PDet AS Pending
FROM
  (
    SELECT
      Observers.Initials,
      Count(Detections.DetectionID) AS TtlDet,
      Sum(
        IIf([Detections].[StatusID] = 2, 1, 0)
      ) AS VDet,
      Sum(
        IIf([Detections].[StatusID] = 3, 1, 0)
      ) AS DDet,
      Sum(
        IIf([Detections].[StatusID] = 1, 1, 0)
      ) AS PDet,
      SUM(
        IIF(
          [Detections].[StatusID] = 3
          AND [VS].[ImageID] Is Null,
          1, 0
        )
      ) AS DelSp
    FROM
      Observers
      INNER JOIN (
        Detections
        LEFT JOIN (
          SELECT
            ImageID, SpeciesID
          FROM
            Detections
          WHERE
            StatusID = 2
          GROUP BY
            ImageID,
            SpeciesID
          ) AS VS ON Detections.ImageID = VS.ImageID
          AND Detections.SpeciesID = VS.SpeciesID
        ) ON Observers.ObserverID = Detections.ObsID
        GROUP BY
          Observers.Initials
      ) AS Q;
WARNING: unclosed parentheses or section
