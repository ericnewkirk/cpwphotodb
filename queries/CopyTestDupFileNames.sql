SELECT
  PhotoViewerQuery.FileName,
  Count(PhotoViewerQuery.ImgID) AS CountOfImgID
FROM
  PhotoViewerQuery
GROUP BY
  PhotoViewerQuery.FileName
HAVING
  (
    (
      (
        Count(PhotoViewerQuery.ImgID)
      )> 1
    )
  );
