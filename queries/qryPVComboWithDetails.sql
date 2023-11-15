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
  Q.DetailID,
  DetectionDetails.DetailText
FROM
  qryPVComboRowSource
  INNER JOIN (
    (
      SELECT
        DISTINCT ImageID, SpeciesID,
        DetailID
      FROM
        Detections
      WHERE
        StatusID < 3
      ) AS Q
      INNER JOIN DetectionDetails ON Q.DetailID = DetectionDetails.DetailID
    ) ON qryPVComboRowSource.ImgID = Q.ImageID;
WARNING: unclosed parentheses or section
