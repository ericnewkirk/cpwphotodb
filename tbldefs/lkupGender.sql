﻿CREATE TABLE [lkupGender] (
  [GenderID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [GenderText] VARCHAR (7),
  [GenderAbbr] VARCHAR (1)
)
