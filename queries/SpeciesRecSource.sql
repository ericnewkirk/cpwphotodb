SELECT
  Species.SpeciesID,
  Species.CommonName,
  Species.Genus,
  Species.Species,
  SpeciesShortcuts.Shortcut,
  Species.ShortName,
  Species.GroupID
FROM
  Species
  LEFT JOIN SpeciesShortcuts ON Species.SpeciesID = SpeciesShortcuts.SpeciesID
ORDER BY
  Species.CommonName;
