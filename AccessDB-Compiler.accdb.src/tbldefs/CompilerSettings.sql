CREATE TABLE [CompilerSettings] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [SourcePath] VARCHAR (255),
  [OutputPath] VARCHAR (255),
  [OverwriteDB] BIT 
)
