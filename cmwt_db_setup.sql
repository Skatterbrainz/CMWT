USE [CMWT]
GO

/****** Object:  Table [dbo].[EventLog]    Script Date: 12/8/2016 3:35:59 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[EventLog](
	[EventID] [int] IDENTITY(1,1) NOT NULL,
	[EventType] [varchar](50) NOT NULL,
	[EventCategory] [varchar](50) NOT NULL,
	[EventOwner] [varchar](50) NOT NULL,
	[EventDateTime] [smalldatetime] NOT NULL CONSTRAINT [DF_EventLog_EventDateTime]  DEFAULT (getdate()),
	[EventDetails] [varchar](255) NOT NULL,
 CONSTRAINT [PK_EventLog] PRIMARY KEY CLUSTERED 
(
	[EventID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

/****** Object:  Table [dbo].[Notes]    Script Date: 12/8/2016 3:35:59 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[Notes](
	[NoteID] [int] IDENTITY(1,1) NOT NULL,
	[AttachedTo] [varchar](255) NOT NULL,
	[AttachClass] [varchar](50) NOT NULL,
	[Comment] [varchar](2000) NOT NULL,
	[DateCreated] [smalldatetime] NOT NULL,
	[DateModified] [smalldatetime] NULL,
	[CreatedBy] [varchar](50) NOT NULL,
	[ModifiedBy] [varchar](50) NULL,
 CONSTRAINT [PK_Notes] PRIMARY KEY CLUSTERED 
(
	[NoteID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

/****** Object:  Table [dbo].[Reports]    Script Date: 12/8/2016 3:35:59 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[Reports](
	[ReportID] [int] IDENTITY(1,1) NOT NULL,
	[ReportName] [varchar](50) NOT NULL,
	[SearchField] [varchar](50) NOT NULL,
	[SearchValue] [varchar](255) NOT NULL,
	[SearchMode] [varchar](50) NOT NULL,
	[DisplayColumns] [varchar](255) NOT NULL,
	[Comment] [varchar](255) NULL,
	[DateCreated] [smalldatetime] NOT NULL,
	[CreatedBy] [varchar](50) NOT NULL,
 CONSTRAINT [PK_Reports] PRIMARY KEY CLUSTERED 
(
	[ReportID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

/****** Object:  Table [dbo].[Reports2]    Script Date: 12/8/2016 3:35:59 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[Reports2](
	[ReportID] [int] IDENTITY(1,1) NOT NULL,
	[ReportType] [int] NOT NULL,
	[ReportName] [varchar](50) NOT NULL,
	[Query] [varchar](2000) NOT NULL,
	[CreatedBy] [varchar](50) NOT NULL,
	[DateCreated] [smalldatetime] NOT NULL,
	[Comment] [varchar](255) NULL,
 CONSTRAINT [PK_Reports2] PRIMARY KEY CLUSTERED 
(
	[ReportID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

/****** Object:  Table [dbo].[Tasks]    Script Date: 12/8/2016 3:35:59 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[Tasks](
	[ActID] [int] IDENTITY(1,1) NOT NULL,
	[ActivityName] [varchar](50) NOT NULL,
	[ActivityType] [varchar](50) NOT NULL,
	[CreatedBy] [varchar](50) NOT NULL,
	[DateTimeCreated] [smalldatetime] NOT NULL,
	[DateTimeExecuted] [smalldatetime] NULL,
	[Result] [int] NULL,
	[Comment] [varchar](255) NULL,
	[CommandString] [varchar](255) NOT NULL,
 CONSTRAINT [PK_Tasks] PRIMARY KEY CLUSTERED 
(
	[ActID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


