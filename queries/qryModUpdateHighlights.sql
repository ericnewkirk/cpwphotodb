UPDATE
  ModulePhotos
  INNER JOIN Photos ON ModulePhotos.ImageID = Photos.ImageID
SET
  Photos.Highlight = True
WHERE
  (
    (
      (ModulePhotos.Highlight)= True
    )
  );
