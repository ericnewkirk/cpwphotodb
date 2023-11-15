CREATE TABLE [lkupQueryBuilders] (
  [QueryBuilderID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [FormName] VARCHAR (255),
  [AnalysisType] VARCHAR (255),
  [ListName] VARCHAR (255)
)
