CREATE TABLE [StudyAreas] (
  [StudyAreaID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [StudyAreaName] VARCHAR (17),
  [StudyAreaAbbr] VARCHAR (3),
  [StudyAreaDescription] VARCHAR (255)
)
