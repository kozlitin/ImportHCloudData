USE [DeNovo_HCloud_Reports]
GO
/****** Object:  StoredProcedure [dbo].[DeleteECloudReportData]    Script Date: 21.11.2016 17:19:35 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[DeleteECloudReportData] 
	@PeriodBegin datetime,
	@PeriodEnd datetime
AS
BEGIN
	SET NOCOUNT ON;

	DELETE FROM [dbo].[ECloudReportData] WHERE Period BETWEEN @PeriodBegin AND @PeriodEnd

END


GO
/****** Object:  StoredProcedure [dbo].[InsertECloudReportData]    Script Date: 21.11.2016 17:19:35 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[InsertECloudReportData] 
	@Period datetime,
	@PeriodName [varchar](50),
	@Year [smallint],
	@Quarter [smallint],
	@Month [smallint],
	@Client [varchar](512),
	@EDRPOU [varchar](12),
	@Industry [varchar](50),
	@ClientSize [varchar](50),
	@ResponsiblePerson [varchar](150),
	@Currency [varchar](10),
	@Category [varchar](50),
	@Service [varchar](512),
	@ServiceAttribute1 [varchar](50),
	@ServiceAttribute2 [varchar](50),
	@ServiceAttribute3 [varchar](50),
	@Unit [varchar](50),
	@Qty [decimal](15, 2),
	@Price [decimal](15, 2),
	@PriceUAH [decimal](15, 2),
	@PriceUSD [decimal](15, 2),
	@Price0 [decimal](15, 2),
	@Price0UAH [decimal](15, 2),
	@Price0USD [decimal](15, 2),
	@Discount [decimal](15, 2)
AS
BEGIN
	SET NOCOUNT ON;

	DECLARE @MaxID INT;

	SELECT @MaxID = ISNULL(MAX(ID),0)+1 FROM [dbo].[ECloudReportData]

	INSERT INTO [dbo].[ECloudReportData]
           ([ID]
           ,[Period]
           ,[PeriodName]
           ,[Year]
           ,[Quarter]
           ,[Month]
           ,[Client]
		   ,[EDRPOU]
           ,[Industry]
           ,[ClientSize]
           ,[ResponsiblePerson]
           ,[Currency]
           ,[Category]
           ,[Service]
           ,[ServiceAttribute1]
           ,[ServiceAttribute2]
           ,[ServiceAttribute3]
		   ,[Unit]
           ,[Qty]
           ,[Price]
           ,[PriceUAH]
           ,[PriceUSD]
           ,[Price0]
           ,[Price0UAH]
           ,[Price0USD]
           ,[Discount])
     VALUES
           (@MaxID
           ,@Period
           ,@PeriodName
           ,@Year
           ,@Quarter
           ,@Month
           ,@Client
		   ,@EDRPOU
           ,@Industry
           ,@ClientSize
           ,@ResponsiblePerson
           ,@Currency
           ,@Category
           ,@Service
           ,@ServiceAttribute1
           ,@ServiceAttribute2
           ,@ServiceAttribute3
		   ,@Unit
           ,@Qty
           ,@Price
           ,@PriceUAH
           ,@PriceUSD
           ,@Price0
           ,@Price0UAH
           ,@Price0USD
           ,@Discount)

END



GO
/****** Object:  Table [dbo].[ECloudReportData]    Script Date: 21.11.2016 17:19:35 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[ECloudReportData](
	[ID] [int] NOT NULL,
	[Period] [datetime] NOT NULL,
	[PeriodName] [varchar](50) NULL,
	[Year] [smallint] NULL,
	[Quarter] [smallint] NULL,
	[Month] [smallint] NULL,
	[Client] [varchar](512) NULL,
	[EDRPOU] [varchar](12) NULL,
	[Industry] [varchar](50) NULL,
	[ClientSize] [varchar](50) NULL,
	[ResponsiblePerson] [varchar](150) NULL,
	[Currency] [varchar](10) NULL,
	[Category] [varchar](50) NULL,
	[Service] [varchar](512) NULL,
	[ServiceAttribute1] [varchar](50) NULL,
	[ServiceAttribute2] [varchar](50) NULL,
	[ServiceAttribute3] [varchar](50) NULL,
	[Unit] [varchar](50) NULL,
	[Qty] [decimal](15, 2) NULL,
	[Price] [decimal](15, 2) NULL,
	[PriceUAH] [decimal](15, 2) NULL,
	[PriceUSD] [decimal](15, 2) NULL,
	[Price0] [decimal](15, 2) NULL,
	[Price0UAH] [decimal](15, 2) NULL,
	[Price0USD] [decimal](15, 2) NULL,
	[Discount] [decimal](15, 2) NULL,
PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  View [dbo].[vECloudReportData]    Script Date: 21.11.2016 17:19:35 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE VIEW [dbo].[vECloudReportData]
AS
SELECT        PeriodName AS Период, Year AS Год, Quarter AS Квартал, Month AS Месяц, Client AS Контрагент, EDRPOU as [ЕДРПОУ], Industry AS Отрасль, ClientSize AS [Размер контрагента], 
                         ResponsiblePerson AS Ответственный, Currency AS Валюта, Category AS Категория, Service AS Услуга, ServiceAttribute1 AS [Атрибут услуги 1], 
                         ServiceAttribute2 AS [Атрибут услуги 2], ServiceAttribute3 AS [Атрибут услуги 3], Unit AS Единица, Qty AS Количество, Price AS Стоимость, PriceUAH AS [Стоимость UAH], 
                         PriceUSD AS [Стоимость USD], Price0 AS [Стоимость без скидки], Price0UAH AS [Стоимость без скидки UAH], Price0USD AS [Стоимость без скидки USD], 
                         Discount AS Скидка
FROM            dbo.ECloudReportData



GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "ECloudReportData"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 239
               Right = 306
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 2895
         Table = 1170
         Output = 1335
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vECloudReportData'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'vECloudReportData'
GO
