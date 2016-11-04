USE [DeNovo_HCloud_Reports]
GO

/****** Object:  Table [dbo].[HCloudReportStage]    Script Date: 04.11.2016 19:42:24 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[HCloudReportStage](
	[ID] [int] NOT NULL,
	[Period] [varchar](50) NULL,
	[SheetName] [varchar](150) NULL,
	[Client] [varchar](512) NULL,
	[ResponsiblePerson] [varchar](150) NULL,
	[Currency] [varchar](50) NULL,
	[Category] [varchar](50) NULL,
	[Item] [varchar](512) NULL,
	[Price] [varchar](50) NULL,
UNIQUE NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

