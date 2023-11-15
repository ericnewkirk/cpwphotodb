﻿CREATE TABLE [IndividualDetections] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [IndDetectionID] LONG  CONSTRAINT [{F331C44D-C575-4B74-B144-B73D46648362}] REFERENCES [IndependentDetections] ([IndDetectionID]) ON UPDATE CASCADE  ON DELETE CASCADE ,
  [IndividualID] LONG  CONSTRAINT [{633FEC9B-C292-4C6A-AD1D-9C5898727E0A}] REFERENCES [Individuals] ([IndividualID]) ON UPDATE CASCADE  ON DELETE CASCADE ,
  [Comments] VARCHAR (255)
)
