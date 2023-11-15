SELECT
  IndependentDetections.IndDetectionID,
  IndependentDetections.SpeciesID,
  IndependentDetections.LocationID,
  IndependentDetections.DefaultStart,
  IndependentDetections.ModifiedStart,
  IndependentDetections.Deleted,
  (
    SELECT
      Min(ModifiedStart)
    FROM
      IndependentDetections AS ID2
    WHERE
      (
        (
          (ID2.LocationID)= IndependentDetections.LocationID
        )
        AND (
          (ID2.SpeciesID)= IndependentDetections.SpeciesID
        )
        AND (
          (ID2.ModifiedStart)> IndependentDetections.ModifiedStart
        )
      )
  ) AS NextDetection,
  (
    SELECT
      Max(ModifiedStart)
    FROM
      IndependentDetections AS ID2
    WHERE
      (
        (
          (ID2.LocationID)= IndependentDetections.LocationID
        )
        AND (
          (ID2.SpeciesID)= IndependentDetections.SpeciesID
        )
        AND (
          (ID2.ModifiedStart)< IndependentDetections.ModifiedStart
        )
        AND (
          (ID2.Deleted)= FALSE
        )
      )
  ) AS PrevDetection,
  (
    SELECT
      First(
        Year(VisitDate)
      )
    FROM
      Visits
    WHERE
      (
        (
          (VisitID)=(
            SELECT
              First(SetVisitID)
            FROM
              Visits
              INNER JOIN Photos ON Visits.VisitID = Photos.VisitID
            WHERE
              (
                (
                  (Visits.LocationID)= IndependentDetections.LocationID
                )
                AND (
                  (Photos.ImageDate)= IndependentDetections.DefaultStart
                )
              )
          )
        )
      )
  ) AS FieldSeason
FROM
  IndependentDetections
  LEFT JOIN IndividualDetections ON IndependentDetections.IndDetectionID = IndividualDetections.IndDetectionID
WHERE
  (
    (
      (IndependentDetections.Deleted)= False
    )
    AND (
      (IndividualDetections.ID) Is Null
    )
  )
ORDER BY
  IndependentDetections.LocationID,
  IndependentDetections.ModifiedStart;
