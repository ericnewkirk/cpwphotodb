UPDATE
  (
    ModuleObservers
    INNER JOIN NewDetections ON ModuleObservers.ObserverID = NewDetections.ObsID
  )
  INNER JOIN Observers ON (
    ModuleObservers.FirstName = Observers.FirstName
  )
  AND (
    ModuleObservers.LastName = Observers.LastName
  )
SET
  NewDetections.ObsID = [Observers].[ObserverID];
