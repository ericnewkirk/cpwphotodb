CREATE TABLE [Photos] (
  [ImageID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ImageNum] LONG ,
  [ImageDate] DATETIME ,
  [FileName] VARCHAR (255),
  [FilePath] VARCHAR (255),
  [Highlight] BIT ,
  [VisitID] LONG  CONSTRAINT [{D781B057-3357-46C9-B4A8-3F8DD48D78D4}] REFERENCES [Visits] ([VisitID]) ON UPDATE CASCADE  ON DELETE CASCADE ,
  [Compare] BIT ,
  [ObsCount] LONG ,
  [Verified] BIT ,
  [Pending] BIT ,
  [MultiSp] BIT ,
  [NotNone] BIT ,
  [NeedsUpdate] BIT 
)
