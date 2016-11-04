USE [DeNovo_HCloud_Reports]
GO

/****** Object:  Table [dbo].[HCloudReportStage]    Script Date: 04.11.2016 13:33:05 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[HCloudReportStage](
	[ID] [int] NOT NULL,
	[Period] [nvarchar](50) NULL,
	[SheetName] [nvarchar](150) NULL,
	[Client] [nvarchar](512) NULL,
	[ResponsiblePerson] [nvarchar](150) NULL,
	[Currency] [nvarchar](50) NULL,
	[Category] [nvarchar](50) NULL,
	[Item] [nvarchar](512) NULL,
	[Price] [nvarchar](50) NULL,
UNIQUE NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

