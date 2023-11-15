INSERT INTO PhotoTags (
  TagX, XLen, TagY, YLen, ImageID, ObsID
)
SELECT
  ModuleTags.TagX,
  ModuleTags.XLen,
  ModuleTags.Tagy,
  ModuleTags.YLen,
  ModuleTags.ImageID,
  ModuleTags.ObsID
FROM
  ModuleTags
  LEFT JOIN PhotoTags ON (
    ModuleTags.ObsID = PhotoTags.ObsID
  )
  AND (
    ModuleTags.ImageID = PhotoTags.ImageID
  )
WHERE
  (
    (
      (PhotoTags.TagID) Is Null
    )
  );
