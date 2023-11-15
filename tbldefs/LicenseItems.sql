CREATE TABLE [LicenseItems] (
  [ItemID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [ItemLetter] VARCHAR (1),
  [ItemText] LONGTEXT ,
  [HeadingID] LONG  CONSTRAINT [LicenseHeadingsLicenseItems] REFERENCES [LicenseHeadings] ([HeadingID])
)
