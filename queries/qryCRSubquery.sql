SELECT
  CameraLocations.StudyAreaID,
  StudyAreas.StudyAreaName,
  CameraLocations.LocationID,
  CameraLocations.LocationName,
  CameraLocations.UTM_E,
  CameraLocations.UTM_N,
  SetVisits.VisitDate AS CameraSetDate,
  Year(SetVisits.VisitDate) AS CameraYear,
  Nz(
    DMin(
      "ActiveStart", "Visits", "SetVisitID=" & SetVisits.VisitID
    ),
    SetVisits.VisitDate
  ) AS OccasionStart,
  IndependentDetections.ModifiedStart,
  IndependentDetections.SpeciesID,
  IndividualDetections.IndividualID,
  Individuals.IndividualName
FROM
  (
    (
      StudyAreas
      INNER JOIN CameraLocations ON StudyAreas.StudyAreaID = CameraLocations.StudyAreaID
    )
    INNER JOIN (
      (
        Visits AS SetVisits
        INNER JOIN (
          SELECT
            Visits.SetVisitID,
            IndependentDetections.IndDetectionID
          FROM
            Visits
            INNER JOIN (
              Photos
              INNER JOIN IndependentDetections ON Photos.ImageDate = IndependentDetections.DefaultStart
            ) ON (
              Visits.LocationID = IndependentDetections.LocationID
            )
            AND (Visits.VisitID = Photos.VisitID)
          GROUP BY
            Visits.SetVisitID,
            IndependentDetections.IndDetectionID
        ) AS VID ON SetVisits.VisitID = VID.SetVisitID
      )
      INNER JOIN IndependentDetections ON VID.IndDetectionID = IndependentDetections.IndDetectionID
    ) ON CameraLocations.LocationID = SetVisits.LocationID
  )
  INNER JOIN (
    Individuals
    INNER JOIN IndividualDetections ON Individuals.IndividualID = IndividualDetections.IndividualID
  ) ON IndependentDetections.IndDetectionID = IndividualDetections.IndDetectionID
WHERE
  (
    (
      (IndependentDetections.Deleted)= False
    )
  );
