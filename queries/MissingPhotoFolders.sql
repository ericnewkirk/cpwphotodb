SELECT
  MissingPhotos.FilePath AS Expr1
FROM
  MissingPhotos
GROUP BY
  MissingPhotos.FilePath;
