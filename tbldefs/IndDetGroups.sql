CREATE TABLE [IndDetGroups] (
  [IndDetGroupID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [GenderID] LONG ,
  [AgeClassID] LONG ,
  [Individuals] LONG ,
  [Comments] VARCHAR (255),
  [IndDetectionID] LONG  CONSTRAINT [IndependentDetectionsIndDetGroups] REFERENCES [IndependentDetections] ([IndDetectionID]) ON UPDATE CASCADE  ON DELETE CASCADE 
)
