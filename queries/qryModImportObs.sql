INSERT INTO Observers (
  LastName, FirstName, Initials, Role,
  Email
)
SELECT
  ModuleObservers.LastName,
  ModuleObservers.FirstName,
  ModuleObservers.Initials,
  ModuleObservers.Role,
  ModuleObservers.Email
FROM
  ModuleObservers
  LEFT JOIN Observers ON (
    ModuleObservers.LastName = Observers.LastName
  )
  AND (
    ModuleObservers.FirstName = Observers.FirstName
  )
WHERE
  (
    (
      (Observers.ObserverID) Is Null
    )
    AND (
      (ModuleObservers.ObserverID) In (
        SELECT
          ObsID
        FROM
          ModuleDetections
        GROUP BY
          ObsID
        )
      )
    );
WARNING: unclosed parentheses or section
