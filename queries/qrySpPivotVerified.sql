TRANSFORM Nz(
  Sum(Q.MaxInd),
  0
) AS Individuals
SELECT
  Q.ImageID
FROM
  (
    (
      SELECT
        Detections.ImageID,
        Detections.SpeciesID,
        Detections.DetailID,
        Max(
          IIf(
            [Detections].[SpeciesID] > 0, [Detections].[Individuals],
            0
          )
        ) AS MaxInd
      FROM
        Detections
      WHERE
        (
          (
            (Detections.StatusID)= 2
          )
        )
      GROUP BY
        Detections.ImageID,
        Detections.SpeciesID,
        Detections.DetailID
    ) AS Q
    INNER JOIN Species ON Q.SpeciesID = Species.SpeciesID
  )
  LEFT JOIN DetectionDetails ON Q.DetailID = DetectionDetails.DetailID
GROUP BY
  Q.ImageID PIVOT [Species].[CommonName] & IIf(
    [DetectionDetails].[DetailID] Is Null,
    "", " - " & [DetectionDetails].[DetailText]
  );
