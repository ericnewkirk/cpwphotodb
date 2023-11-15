CREATE TABLE [Species] (
  [SpeciesID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [CommonName] VARCHAR (64),
  [Genus] VARCHAR (32),
  [Species] VARCHAR (32),
  [ShortName] VARCHAR (255),
  [GroupID] LONG  CONSTRAINT [SpeciesGroupsSpecies] REFERENCES [SpeciesGroups] ([GroupID])
)
