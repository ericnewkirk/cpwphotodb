CREATE TABLE [lkupExports] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [QueryName] VARCHAR (255),
  [QueryName2] VARCHAR (255),
  [Program] VARCHAR (255),
  [Extension] VARCHAR (255)
)
