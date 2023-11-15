SELECT
  StudyAreas.StudyAreaName,
  StudyAreas.StudyAreaID,
  CameraLocations.LocationName,
  CameraLocations.LocationID,
  CameraLocations.UTM_E,
  CameraLocations.UTM_N,
  CameraLocations.UTMZone,
  CameraLocations.LatitudeDD,
  CameraLocations.LongitudeDD,
  Year([SetVisits].[VisitDate]) AS FieldSeason,
  Photos.FileName,
  Visits.VisitID,
  Photos.ImageID AS ImgID,
  Photos.ImageNum,
  Photos.ImageDate,
  Photos.Highlight,
  [Photos].[FilePath] & [Photos].[FileName] AS ImgPath,
  DetQuery.SpeciesID,
  DetQuery.CommonName,
  DetQuery.DetailText,
  DetQuery.Individuals
FROM
  (
    StudyAreas
    INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID
  )
  INNER JOIN (
    (
      (
        Visits AS SetVisits
        INNER JOIN Visits ON SetVisits.VisitID = Visits.SetVisitID
      )
      INNER JOIN Photos ON Visits.VisitID = Photos.VisitID
    )
    LEFT JOIN (
      SELECT
        Q.ImageID,
        Species.SpeciesID,
        Species.CommonName,
        DetectionDetails.DetailText,
        Q.MaxInd As Individuals
      FROM
        (
          (
            SELECT
              Detections.ImageID,
              Detections.SpeciesID,
              Detections.DetailID,
              Max(
                IIf(
                  [Detections].[SpeciesID] > 0, [Detections].[Individuals],
                  0
                )
              ) AS MaxInd
            FROM
              Detections
            WHERE
              (
                (
                  (Detections.StatusID)< 3
                )
              )
            GROUP BY
              Detections.ImageID,
              Detections.SpeciesID,
              Detections.DetailID
          ) As Q
          INNER JOIN Species ON 
            Q.SpeciesID = Species.SpeciesID
        )
        LEFT JOIN DetectionDetails ON 
            Q.DetailID = DetectionDetails.DetailID
    ) AS DetQuery ON Photos.ImageID = DetQuery.ImageID
  ) ON CameraLocations.LocationID = SetVisits.LocationID
ORDER BY
  Photos.ImageID;
