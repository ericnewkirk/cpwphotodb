SELECT
  qryPVComboRowSource.StudyAreaID,
  qryPVComboRowSource.LocationID,
  qryPVComboRowSource.FieldSeason,
  qryPVComboRowSource.VisitID,
  qryPVComboRowSource.ImgID,
  qryPVComboRowSource.ImageDate,
  qryPVComboRowSource.Highlight,
  qryPVComboRowSource.ObsCount,
  qryPVComboRowSource.Verified,
  qryPVComboRowSource.Pending,
  qryPVComboRowSource.MultiSp,
  qryPVComboRowSource.NotNone,
  Q.SpeciesID,
  Species.ShortName
FROM
  qryPVComboRowSource
  INNER JOIN (
    (
      SELECT
        DISTINCT ImageID, SpeciesID
      FROM
        Detections
      WHERE
        StatusID < 3
      ) AS Q
      INNER JOIN Species ON Q.SpeciesID = Species.SpeciesID
    ) ON qryPVComboRowSource.ImgID = Q.ImageID;
WARNING: unclosed parentheses or section
