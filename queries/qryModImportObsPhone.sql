INSERT INTO ObsPhone (
  ObserverID, TypeID, PhoneNumber, Extension
)
SELECT
  Observers.ObserverID,
  ModuleObsPhone.TypeID,
  ModuleObsPhone.PhoneNumber,
  ModuleObsPhone.Extension
FROM
  (
    (
      ModuleObservers
      INNER JOIN Observers ON (
        ModuleObservers.FirstName = Observers.FirstName
      )
      AND (
        ModuleObservers.LastName = Observers.LastName
      )
    )
    INNER JOIN ModuleObsPhone ON ModuleObservers.ObserverID = ModuleObsPhone.ObserverID
  )
  LEFT JOIN ObsPhone ON Observers.ObserverID = ObsPhone.ObserverID
WHERE
  (
    (
      (ObsPhone.ID) Is Null
    )
  );
