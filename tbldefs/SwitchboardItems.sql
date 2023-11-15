CREATE TABLE [SwitchboardItems] (
  [ItemID] LONG  CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Caption] VARCHAR (255),
  [Action] LONG ,
  [Argument] VARCHAR (255)
)
