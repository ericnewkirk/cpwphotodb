INSERT INTO Species (
  CommonName, Genus, Species, ShortName
)
SELECT
  ModuleSpecies.CommonName,
  ModuleSpecies.Genus,
  ModuleSpecies.Species,
  ModuleSpecies.ShortName
FROM
  ModuleSpecies
  LEFT JOIN Species ON ModuleSpecies.CommonName = Species.CommonName
WHERE
  (
    (
      (Species.SpeciesID) Is Null
    )
    AND (
      (ModuleSpecies.SpeciesID) In (
        SELECT
          SpeciesID
        FROM
          ModuleDetections
        GROUP BY
          SpeciesID
        )
      )
    );
WARNING: unclosed parentheses or section
