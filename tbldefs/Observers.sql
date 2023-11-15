CREATE TABLE [Observers] (
  [ObserverID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [LastName] VARCHAR (50),
  [FirstName] VARCHAR (50),
  [Initials] VARCHAR (5),
  [Role] VARCHAR (50),
  [Email] VARCHAR (100)
)
