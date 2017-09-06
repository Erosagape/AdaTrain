use AdaTrain
go
/****** Object:  Table [dbo].[TTRMCst]    Script Date: 01/09/2017 10:54:19 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[TTRMCst](
	[FTCstCode] [nvarchar](15) NOT NULL,
	[FTCstName] [nvarchar](100) NOT NULL,
	[FTCstAddress] [nvarchar](100) NULL,
	[FTCstPriceLv] [varchar](1) NOT NULL,
	[FTCstStatus] [varchar](1) NOT NULL,
	[FTCstTel] [nvarchar](10) NULL,
	[FTCstFax] [nvarchar](10) NULL,
	[FCCstARBal] [float] NULL,
	[FCCstChqBal] [float] NULL,
	[FDDateUpd] [datetime] NULL,
	[FTTimeUpd] [varchar](8) NULL,
	[FTWhoUpd] [varchar](50) NULL,
	[FDDateIns] [datetime] NULL,
	[FTTimeIns] [varchar](8) NULL,
	[FTWhoIns] [varchar](50) NULL,
	[FTRemark] [varchar](100) NULL,
	[FDBirthDate] [date] NULL,
	[FCCreditLimit] [decimal](18, 4) NULL,
 CONSTRAINT [PK_TTRMCustomer] PRIMARY KEY CLUSTERED 
(
	[FTCstCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[TTRMPdt]    Script Date: 01/09/2017 10:54:19 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TTRMPdt](
	[FTPdtCode] [nvarchar](15) NOT NULL,
	[FTPdtBarCode] [nvarchar](50) NOT NULL,
	[FTPdtName] [nvarchar](100) NOT NULL,
	[FTPdtUnit] [nvarchar](10) NOT NULL,
	[FTPdtGroup] [nvarchar](10) NOT NULL,
	[FCPriceSale1] [float] NOT NULL,
	[FCPriceSale2] [float] NOT NULL,
	[FCPriceSale3] [float] NOT NULL,
	[FCPriceSale4] [float] NOT NULL,
	[FDDateUpd] [datetime] NULL,
	[FTTimeUpd] [varchar](8) NULL,
	[FTWhoUpd] [varchar](50) NULL,
	[FDDateIns] [datetime] NULL,
	[FTTimeIns] [varchar](8) NULL,
	[FTWhoIns] [varchar](50) NULL,
	[FTRemark] [varchar](100) NULL,
 CONSTRAINT [PK_TTRMProduct] PRIMARY KEY CLUSTERED 
(
	[FTPdtCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[TTRMPdtGrp]    Script Date: 01/09/2017 10:54:19 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TTRMPdtGrp](
	[FTPdtGrpCode] [nvarchar](10) NOT NULL,
	[FTPdtGrpName] [nvarchar](100) NOT NULL,
	[FDDateUpd] [datetime] NULL,
	[FTTimeUpd] [varchar](8) NULL,
	[FTWhoUpd] [varchar](50) NULL,
	[FDDateIns] [datetime] NULL,
	[FTTimeIns] [varchar](8) NULL,
	[FTWhoIns] [varchar](50) NULL,
	[FTRemark] [varchar](100) NULL,
 CONSTRAINT [PK_TTRMPdtGrp] PRIMARY KEY CLUSTERED 
(
	[FTPdtGrpCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[TTRMUnit]    Script Date: 01/09/2017 10:54:19 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TTRMUnit](
	[FTUntCode] [nvarchar](10) NOT NULL,
	[FTUntName] [nvarchar](100) NOT NULL,
	[FDDateUpd] [datetime] NULL,
	[FTTimeUpd] [varchar](8) NULL,
	[FTWhoUpd] [varchar](50) NULL,
	[FDDateIns] [datetime] NULL,
	[FTTimeIns] [varchar](8) NULL,
	[FTWhoIns] [varchar](50) NULL,
	[FTRemark] [varchar](100) NULL,
 CONSTRAINT [PK_TTRMUnit] PRIMARY KEY CLUSTERED 
(
	[FTUntCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

/****** Object:  Table [dbo].[TTRMSlePsn]    Script Date: 01/09/2017 10:54:19 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TTRMSpn](
	[FTSpnCode] [nvarchar](15) NOT NULL,
	[FTSpnName] [nvarchar](50) NOT NULL,
	[FDDateUpd] [datetime] NULL,
	[FTTimeUpd] [varchar](8) NULL,
	[FTWhoUpd] [varchar](50) NULL,
	[FDDateIns] [datetime] NULL,
	[FTTimeIns] [varchar](8) NULL,
	[FTWhoIns] [varchar](50) NULL,
	[FTRemark] [varchar](100) NULL,
 CONSTRAINT [PK_TTRSlePsn] PRIMARY KEY CLUSTERED 
(
	[FTSpnCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[TTRTSleSlipHD]    Script Date: 01/09/2017 10:54:19 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[TTRTSleHD](
	[FTSleHDDocNo] [nvarchar](25) NOT NULL,
	[FDSleHDDocDate] [date] NOT NULL,
	[FTSleHDCstCode] [nvarchar](15) NOT NULL,
	[FTSleHDCstName] [nvarchar](100) NULL,
	[FTSleHDDocType] [varchar](1) NOT NULL,
	[FCSleHDVatRate] [float] NOT NULL,
	[FTSleHDVatType] [varchar](1) NOT NULL,
	[FCSleHDDocAmt] [float] NOT NULL,
	[FCSleHDDiscAmt] [float] NOT NULL,
	[FCSleHDBeforeVat] [float] NOT NULL,
	[FCSleHDVatAmt] [float] NOT NULL,
	[FCSleHDDocTotal] [float] NOT NULL,
	[FTSleHDSpnCode] [nvarchar](15) NOT NULL,
	[FTSleHDSpnName] [nvarchar](100) NULL,
	[FDDateUpd] [datetime] NULL,
	[FTTimeUpd] [varchar](8) NULL,
	[FTWhoUpd] [varchar](50) NULL,
	[FDDateIns] [datetime] NULL,
	[FTTimeIns] [varchar](8) NULL,
	[FTWhoIns] [varchar](50) NULL,
	[FTRemark] [varchar](100) NULL,
 CONSTRAINT [PK_TTRTSaleSlipHD] PRIMARY KEY CLUSTERED 
(
	[FTSleHDDocNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/****** Object:  Table [dbo].[TTRTSleSlipDT]    Script Date: 01/09/2017 10:54:19 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[TTRTSleDT](
	[FTSleHDDocNo] [nvarchar](15) NOT NULL,
	[FNSleDTSeq] [smallint] NOT NULL,
	[FDSleDTDocDate] [date] NOT NULL,
	[FTSleDTPdtCode] [nvarchar](15) NOT NULL,
	[FTSleDTPdtName] [nvarchar](100) NOT NULL,
	[FTSleDTPdtUnit] [nvarchar](10) NOT NULL,
	[FCSleDTQty] [float] NOT NULL,
	[FCSleDTPrice] [float] NOT NULL,
	[FCSleDTDisc] [float] NOT NULL,
	[FCSleDTAmt] [float] NOT NULL,
	[FDDateUpd] [datetime] NULL,
	[FTTimeUpd] [varchar](8) NULL,
	[FTWhoUpd] [varchar](50) NULL,
	[FDDateIns] [datetime] NULL,
	[FTTimeIns] [varchar](8) NULL,
	[FTWhoIns] [varchar](50) NULL,
	[FTRemark] [varchar](100) NULL,
 CONSTRAINT [PK_TTRTSleSlipDT] PRIMARY KEY CLUSTERED 
(
	[FTSleHDDocNo] ASC,
	[FNSleDTSeq] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
create table TTRMPdtSet
(
	[FTPdtCode] [nvarchar](15) not null,
	[FTPdtUntSet] [nvarchar](10) not null,
	[FCPdtQtySet] [float] not null,
	[FDDateUpd] [datetime] NULL,
	[FTTimeUpd] [varchar](8) NULL,
	[FTWhoUpd] [varchar](50) NULL,
	[FDDateIns] [datetime] NULL,
	[FTTimeIns] [varchar](8) NULL,
	[FTWhoIns] [varchar](50) NULL,
	[FTRemark] [varchar](100) NULL,
CONSTRAINT PK_TTRMPdtSet PRIMARY KEY CLUSTERED
(
	FTPdtCode, FTPdtUntSet
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)
GO
SET ANSI_PADDING OFF
GO
