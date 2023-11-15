INSERT INTO Detections (
  DetectionID, ImageID, SpeciesID, DetailID,
  Individuals, ObsID, Comments
)
SELECT
  NewDetections.DetectionID,
  NewDetections.ImageID,
  NewDetections.SpeciesID,
  NewDetections.DetailID,
  NewDetections.Individuals,
  NewDetections.ObsID,
  NewDetections.Comments
FROM
  NewDetections;
