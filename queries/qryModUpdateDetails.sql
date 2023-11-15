UPDATE
  (
    Species
    INNER JOIN (
      (
        ModuleDetails
        INNER JOIN NewDetections ON ModuleDetails.DetailID = NewDetections.DetailID
      )
      INNER JOIN DetectionDetails ON ModuleDetails.DetailText = DetectionDetails.DetailText
    ) ON Species.SpeciesID = DetectionDetails.SpeciesID
  )
  INNER JOIN ModuleSpecies ON (
    Species.CommonName = ModuleSpecies.CommonName
  )
  AND (
    ModuleDetails.SpeciesID = ModuleSpecies.SpeciesID
  )
SET
  NewDetections.DetailID = [DetectionDetails].[DetailID];
