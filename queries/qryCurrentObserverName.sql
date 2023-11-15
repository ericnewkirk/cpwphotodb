SELECT
  CurrentObserver.ObserverID,
  [FirstName] & " " & [LastName] AS COName
FROM
  Observers
  INNER JOIN CurrentObserver ON Observers.ObserverID = CurrentObserver.ObserverID;
