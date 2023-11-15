UPDATE
  (
    ModuleSpecies
    INNER JOIN NewDetections ON ModuleSpecies.SpeciesID = NewDetections.SpeciesID
  )
  INNER JOIN Species ON ModuleSpecies.CommonName = Species.CommonName
SET
  NewDetections.SpeciesID = [Species].[SpeciesID];
