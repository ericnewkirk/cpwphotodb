INSERT INTO DetectionDetails (DetailText, SpeciesID)
SELECT
  MSD.DetailText,
  MSD.SpeciesID
FROM
  (
    SELECT
      ModuleDetails.DetailText,
      Species.SpeciesID
    FROM
      (
        ModuleDetails
        INNER JOIN ModuleSpecies ON ModuleDetails.SpeciesID = ModuleSpecies.SpeciesID
      )
      INNER JOIN Species ON ModuleSpecies.CommonName = Species.CommonName
    WHERE
      (
        (
          (ModuleDetails.DetailText) Is Not Null
        )
        AND (
          (ModuleDetails.DetailID) In (
            SELECT
              DetailID
            FROM
              ModuleDetections
            GROUP BY
              DetailID
            )
          )
        )
      ) AS MSD
      LEFT JOIN DetectionDetails ON (
        MSD.SpeciesID = DetectionDetails.SpeciesID
      )
      AND (
        MSD.DetailText = DetectionDetails.DetailText
      )
    WHERE
      (
        (
          (DetectionDetails.DetailID) Is Null
        )
      );
WARNING: unclosed parentheses or section
