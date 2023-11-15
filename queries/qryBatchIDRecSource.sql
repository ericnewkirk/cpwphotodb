SELECT
  PhotoIDRecSource.ImageID,
  PhotoIDRecSource.FileName,
  IIf(
    PhotoIDRecSource.ImageID = [Forms] ! [PhotoID] ! [ImageID],
    'Current Photo', ''
  ) AS [Current],
  PhotoIDRecSource.ImgPath
FROM
  PhotoIDRecSource
ORDER BY
  PhotoIDRecSource.ImageID;
