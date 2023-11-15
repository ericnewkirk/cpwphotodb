CREATE TABLE [lkupMaintenanceTasks] (
  [TaskID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [TaskName] VARCHAR (255),
  [TaskDescription] VARCHAR (255)
)
