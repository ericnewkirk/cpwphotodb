SELECT
  Photos.ImageID,
  Photos.FilePath,
  Photos.FileName
FROM
  Photos
WHERE
  (
    (
      (
        FileExists([FilePath] & [FileName])
      )= False
    )
  );
