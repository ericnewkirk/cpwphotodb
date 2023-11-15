UPDATE
  ModuleDetections
  INNER JOIN Photos ON ModuleDetections.ImageID = Photos.ImageID
SET
  Photos.NeedsUpdate = True;
