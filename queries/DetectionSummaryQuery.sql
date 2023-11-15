SELECT
  StudyAreas.StudyAreaName AS StudyArea,
  CameraLocations.LocationName AS Location,
  Species.CommonName AS Species,
  qSub.StartDateTime,
  qSub.TotalIndividuals,
  qSub.TotalAdults,
  qSub.TotalJuveniles,
  qSub.TotalSubadults,
  qSub.TotalFemales,
  qSub.TotalMales,
  qSub.TotalKnownInd
FROM
  StudyAreas
  INNER JOIN (
    CameraLocations
    INNER JOIN (
      Species
      INNER JOIN (
        SELECT
          IndependentDetections.SpeciesID,
          IndependentDetections.LocationID,
          IndependentDetections.ModifiedStart As StartDateTime,
          qGroup.TotalIndividuals,
          qGroup.TotalAdults,
          qGroup.TotalJuveniles,
          qGroup.TotalSubadults,
          qGroup.TotalMales,
          qGroup.TotalMales,
          qGroup.TotalFemales,
          qKnown.KnownIndividuals As TotalKnownInd
        FROM
          (
            IndependentDetections
            INNER JOIN (
              SELECT
                IndependentDetections.IndDetectionID,
                Nz(
                  Sum(IndDetGroups.Individuals),
                  0
                ) AS TotalIndividuals,
                Sum(
                  IIf([AgeClassID] = 1, [Individuals], 0)
                ) AS TotalAdults,
                Sum(
                  IIf([AgeClassID] = 2, [Individuals], 0)
                ) AS TotalJuveniles,
                Sum(
                  IIf([AgeClassID] = 3, [Individuals], 0)
                ) AS TotalSubadults,
                Sum(
                  IIf([GenderID] = 1, [Individuals], 0)
                ) AS TotalFemales,
                Sum(
                  IIf([GenderID] = 2, [Individuals], 0)
                ) AS TotalMales
              FROM
                IndependentDetections
                LEFT JOIN IndDetGroups ON IndependentDetections.IndDetectionID = IndDetGroups.IndDetectionID
              GROUP BY
                IndependentDetections.IndDetectionID
            ) As qGroup ON IndependentDetections.IndDetectionID = qGroup.IndDetectionID
          )
          INNER JOIN (
            SELECT
              IndependentDetections.IndDetectionID,
              Sum(
                IIf(
                  IndividualDetections.IndividualID > 0,
                  1, 0
                )
              ) AS KnownIndividuals
            FROM
              IndependentDetections
              LEFT JOIN IndividualDetections ON IndependentDetections.IndDetectionID = IndividualDetections.IndDetectionID
            GROUP BY
              IndependentDetections.IndDetectionID
          ) As qKnown ON IndependentDetections.IndDetectionID = qKnown.IndDetectionID
        WHERE
          (
            (
              (IndependentDetections.Deleted)= False
            )
          )
      ) AS qSub ON Species.SpeciesID = qSub.SpeciesID
    ) ON CameraLocations.LocationID = qSub.LocationID
  ) ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID
ORDER BY
  Species.CommonName,
  StudyAreas.StudyAreaName,
  CameraLocations.LocationName,
  qSub.StartDateTime;
