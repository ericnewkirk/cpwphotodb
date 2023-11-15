SELECT
  DetectionDetails.DetailID,
  DetectionDetails.DetailText,
  DetectionDetails.SpeciesID,
  DetailShortcuts.Shortcut
FROM
  Species
  INNER JOIN (
    DetectionDetails
    LEFT JOIN DetailShortcuts ON DetectionDetails.DetailID = DetailShortcuts.DetailID
  ) ON Species.SpeciesID = DetectionDetails.SpeciesID
ORDER BY
  Species.CommonName,
  DetailShortcuts.Shortcut;
