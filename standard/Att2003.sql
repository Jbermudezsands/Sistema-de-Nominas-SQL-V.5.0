if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_Class]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_Class]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_Record]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_Record]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[V_UserClient]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[V_UserClient]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AddTimeSet]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AddTimeSet]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BasePara]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BasePara]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CheckLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CheckLog]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Checkinout]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Checkinout]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClientSet]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClientSet]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DefineField]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DefineField]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dept]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dept]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FingerClient]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FingerClient]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Holiday]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Holiday]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LeaveClass]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LeaveClass]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MemStat]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MemStat]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[OPLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[OPLog]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[OPinfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[OPinfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[OutProg]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[OutProg]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SchTime]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SchTime]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Schedule]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Schedule]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StatItems]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StatItems]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Status]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Status]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[T_Checkinout]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[T_Checkinout]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[T_UpdateClient]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[T_UpdateClient]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TimeTable]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TimeTable]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UserCtrLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UserCtrLog]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UserLeave]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UserLeave]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UserPower]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UserPower]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UserShift]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UserShift]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UserTempShift]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UserTempShift]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Userinfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Userinfo]
GO

CREATE TABLE [dbo].[AddTimeSet] (
	[Userid] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Date] [smalldatetime] NOT NULL ,
	[TimeID] [int] NOT NULL ,
	[AddTime] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BasePara] (
	[Company] [varchar] (100) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[TwoDay] [smallint] NOT NULL ,
	[NoClockIn] [bit] NOT NULL ,
	[NoClockOut] [bit] NOT NULL ,
	[LateTime] [smallint] NOT NULL ,
	[LeaveTime] [smallint] NOT NULL ,
	[ISOverTime] [bit] NOT NULL ,
	[OverTime] [smallint] NOT NULL ,
	[WorkDayLong] [smallint] NOT NULL ,
	[WOverTime] [numeric](18, 1) NULL ,
	[HOverTime] [numeric](18, 1) NULL ,
	[FOverTime] [numeric](18, 1) NULL ,
	[IsAutoDownRec] [bit] NULL ,
	[DownRecTime] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[OvertimeIn] [smallint] NULL ,
	[IsovertimeIn] [bit] NULL ,
	[DeductIn] [bit] NULL ,
	[DeductOut] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CheckLog] (
	[Logid] [int] IDENTITY (1, 1) NOT NULL ,
	[Userid] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Checktime] [datetime] NOT NULL ,
	[Checktype] [varchar] (2) COLLATE Chinese_PRC_CI_AS NULL ,
	[Sensorid] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[OPFlag] [smallint] NULL ,
	[Whys] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[OPname] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[DTime] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Checkinout] (
	[Logid] [int] IDENTITY (1, 1) NOT NULL ,
	[Userid] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[CheckTime] [datetime] NOT NULL ,
	[CheckType] [varchar] (2) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Sensorid] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[Checked] [bit] NULL ,
	[WorkType] [int] NULL ,
	[AttFlag] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClientSet] (
	[Clientid] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Pwd] [varchar] (150) COLLATE Chinese_PRC_CI_AS NULL ,
	[DTime] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DefineField] (
	[Fieldid] [int] IDENTITY (1, 1) NOT NULL ,
	[FieldName] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[FieldValue] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dept] (
	[Deptid] [int] IDENTITY (1, 1) NOT NULL ,
	[DeptName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[SupDeptid] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FingerClient] (
	[Clientid] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Linkmode] [smallint] NOT NULL ,
	[ClientName] [varchar] (100) COLLATE Chinese_PRC_CI_AS NULL ,
	[IPaddress] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[ClientNumber] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[RecStatus] [int] NULL ,
	[Baudrate] [int] NULL ,
	[Floorid] [int] NULL ,
	[MachineType] [int] NULL ,
	[DeviceType] [int] NULL ,
	[CommPWD] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[CommPort] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Holiday] (
	[Holidayid] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[BDate] [smalldatetime] NOT NULL ,
	[Days] [smallint] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LeaveClass] (
	[Classid] [int] IDENTITY (1, 1) NOT NULL ,
	[Classname] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[ViewColor] [int] NOT NULL ,
	[Showas] [varchar] (2) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MemStat] (
	[Lsh] [int] IDENTITY (1, 1) NOT NULL ,
	[Userid] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Udept] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Uname] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Normal] [real] NOT NULL ,
	[Actual] [real] NOT NULL ,
	[Latetime] [real] NOT NULL ,
	[Earlytime] [real] NOT NULL ,
	[Absenttime] [real] NOT NULL ,
	[Overtime] [real] NOT NULL ,
	[Noin] [int] NOT NULL ,
	[Noout] [int] NOT NULL ,
	[Awaytime] [real] NOT NULL ,
	[BLeave] [real] NOT NULL ,
	[Leave] [real] NOT NULL ,
	[Freeovertime] [real] NOT NULL ,
	[Worktime] [real] NOT NULL ,
	[Attrate] [real] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[OPLog] (
	[Logid] [int] IDENTITY (1, 1) NOT NULL ,
	[OPid] [int] NOT NULL ,
	[Optime] [datetime] NOT NULL ,
	[OPlog] [varchar] (250) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[OPinfo] (
	[Opid] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Pwd] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Authority] [varchar] (32) COLLATE Chinese_PRC_CI_AS NULL ,
	[Deptpower] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[OutProg] (
	[Progid] [int] IDENTITY (1, 1) NOT NULL ,
	[ProgName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[ProgPath] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SchTime] (
	[Schid] [int] NOT NULL ,
	[BeginDay] [smallint] NOT NULL ,
	[Timeid] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Schedule] (
	[Schid] [int] IDENTITY (1, 1) NOT NULL ,
	[Schname] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Cycles] [smallint] NOT NULL ,
	[Units] [smallint] NOT NULL ,
	[AutoClass] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[StatItems] (
	[Itemid] [int] IDENTITY (1, 1) NOT NULL ,
	[ItemName] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Units] [smallint] NOT NULL ,
	[MinUnit] [numeric](18, 1) NOT NULL ,
	[SRControl] [smallint] NOT NULL ,
	[IsAddup] [bit] NOT NULL ,
	[IsTimes] [bit] NOT NULL ,
	[Showas] [varchar] (2) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Status] (
	[Statusid] [int] NOT NULL ,
	[StatusChar] [varchar] (2) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[StatusText] [varchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[T_Checkinout] (
	[Logid] [int] IDENTITY (1, 1) NOT NULL ,
	[Userid] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[CheckTime] [datetime] NULL ,
	[CheckType] [varchar] (2) COLLATE Chinese_PRC_CI_AS NULL ,
	[Sensorid] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[Checked] [bit] NOT NULL ,
	[WorkType] [int] NOT NULL ,
	[AttFlag] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[T_UpdateClient] (
	[Clientid] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TimeTable] (
	[Timeid] [int] IDENTITY (1, 1) NOT NULL ,
	[Timename] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Intime] [varchar] (5) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Outtime] [varchar] (5) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[BIntime] [varchar] (5) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[EIntime] [varchar] (5) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[BOuttime] [varchar] (5) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[EOuttime] [varchar] (5) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Latetime] [smallint] NOT NULL ,
	[Leavetime] [smallint] NOT NULL ,
	[WorkDays] [numeric](18, 1) NOT NULL ,
	[Longtime] [smallint] NOT NULL ,
	[MustIn] [bit] NULL ,
	[MustOut] [bit] NULL ,
	[IsFreeTime] [bit] NULL ,
	[IsOverTime] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserCtrLog] (
	[Logid] [int] IDENTITY (1, 1) NOT NULL ,
	[Clientid] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[Userid] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[CheckTime] [datetime] NULL ,
	[ULog] [varchar] (150) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserLeave] (
	[Lsh] [int] IDENTITY (1, 1) NOT NULL ,
	[Userid] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[BeginTime] [datetime] NOT NULL ,
	[EndTime] [datetime] NOT NULL ,
	[LeaveClassid] [int] NOT NULL ,
	[Whys] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserPower] (
	[Logid] [int] IDENTITY (1, 1) NOT NULL ,
	[Userid] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[ClientNumber] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[PowerFlag] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserShift] (
	[Userid] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Schid] [int] NOT NULL ,
	[BeginDate] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[EndDate] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserTempShift] (
	[Userid] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Timeid] [int] NOT NULL ,
	[WorkDate] [varchar] (10) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[IsOvertime] [bit] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Userinfo] (
	[Userid] [varchar] (20) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[Name] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Sex] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[Pwd] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Deptid] [int] NOT NULL ,
	[Nation] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Brithday] [smalldatetime] NULL ,
	[EmployDate] [smalldatetime] NULL ,
	[Telephone] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Duty] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[NativePlace] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[IDCard] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Address] [varchar] (150) COLLATE Chinese_PRC_CI_AS NULL ,
	[Mobile] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Educated] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Polity] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[Specialty] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[IsAtt] [bit] NULL ,
	[Isovertime] [bit] NULL ,
	[Isrest] [bit] NULL ,
	[Remark] [varchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[MgFlag] [smallint] NULL ,
	[CardNum] [varchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[Picture] [image] NULL ,
	[UserFlag] [int] NULL ,
	[Groupid] [int] NULL ,
	[workdaylong] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [dbo].[AddTimeSet] WITH NOCHECK ADD 
	CONSTRAINT [PK_AddTimeSet] PRIMARY KEY  CLUSTERED 
	(
		[Userid],
		[Date],
		[TimeID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BasePara] WITH NOCHECK ADD 
	CONSTRAINT [PK_BasePara] PRIMARY KEY  CLUSTERED 
	(
		[Company]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[CheckLog] WITH NOCHECK ADD 
	CONSTRAINT [PK_CheckLog] PRIMARY KEY  CLUSTERED 
	(
		[Logid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Checkinout] WITH NOCHECK ADD 
	CONSTRAINT [PK_Checkinout] PRIMARY KEY  CLUSTERED 
	(
		[Logid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ClientSet] WITH NOCHECK ADD 
	CONSTRAINT [PK_ClientSet] PRIMARY KEY  CLUSTERED 
	(
		[Clientid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[DefineField] WITH NOCHECK ADD 
	CONSTRAINT [PK_DefineField] PRIMARY KEY  CLUSTERED 
	(
		[Fieldid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dept] WITH NOCHECK ADD 
	CONSTRAINT [PK_Dept] PRIMARY KEY  CLUSTERED 
	(
		[Deptid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FingerClient] WITH NOCHECK ADD 
	CONSTRAINT [PK_FingerClient] PRIMARY KEY  CLUSTERED 
	(
		[Clientid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Holiday] WITH NOCHECK ADD 
	CONSTRAINT [PK_Holiday] PRIMARY KEY  CLUSTERED 
	(
		[Holidayid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[LeaveClass] WITH NOCHECK ADD 
	CONSTRAINT [PK_LeaveClass] PRIMARY KEY  CLUSTERED 
	(
		[Classid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MemStat] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[Lsh]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[OPLog] WITH NOCHECK ADD 
	CONSTRAINT [PK_OPLog] PRIMARY KEY  CLUSTERED 
	(
		[Logid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[OPinfo] WITH NOCHECK ADD 
	CONSTRAINT [PK_OPinfo] PRIMARY KEY  CLUSTERED 
	(
		[Opid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[OutProg] WITH NOCHECK ADD 
	CONSTRAINT [PK_OutProg] PRIMARY KEY  CLUSTERED 
	(
		[Progid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[SchTime] WITH NOCHECK ADD 
	CONSTRAINT [PK_SchTime] PRIMARY KEY  CLUSTERED 
	(
		[Schid],
		[BeginDay],
		[Timeid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Schedule] WITH NOCHECK ADD 
	CONSTRAINT [PK_Schedule] PRIMARY KEY  CLUSTERED 
	(
		[Schid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[StatItems] WITH NOCHECK ADD 
	CONSTRAINT [PK_StatItems] PRIMARY KEY  CLUSTERED 
	(
		[Itemid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Status] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[Statusid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[T_Checkinout] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[Logid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[T_UpdateClient] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[Clientid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[TimeTable] WITH NOCHECK ADD 
	CONSTRAINT [PK_TimeTable] PRIMARY KEY  CLUSTERED 
	(
		[Timeid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserCtrLog] WITH NOCHECK ADD 
	CONSTRAINT [PK_UserCtrLog] PRIMARY KEY  CLUSTERED 
	(
		[Logid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserLeave] WITH NOCHECK ADD 
	CONSTRAINT [PK_UserLeave] PRIMARY KEY  CLUSTERED 
	(
		[Lsh]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserPower] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[Logid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserShift] WITH NOCHECK ADD 
	CONSTRAINT [PK_UserShift] PRIMARY KEY  CLUSTERED 
	(
		[Userid],
		[Schid],
		[BeginDate],
		[EndDate]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[UserTempShift] WITH NOCHECK ADD 
	CONSTRAINT [PK_UserTempShift] PRIMARY KEY  CLUSTERED 
	(
		[Userid],
		[Timeid],
		[WorkDate]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Userinfo] WITH NOCHECK ADD 
	CONSTRAINT [PK_Userinfo] PRIMARY KEY  CLUSTERED 
	(
		[Userid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BasePara] ADD 
	CONSTRAINT [DF_BasePara_TwoDay] DEFAULT (0) FOR [TwoDay],
	CONSTRAINT [DF_BasePara_NoClockIn] DEFAULT (0) FOR [NoClockIn],
	CONSTRAINT [DF_BasePara_NoClockOut] DEFAULT (0) FOR [NoClockOut],
	CONSTRAINT [DF_BasePara_LateTime] DEFAULT (60) FOR [LateTime],
	CONSTRAINT [DF_BasePara_LeaveTime] DEFAULT (60) FOR [LeaveTime],
	CONSTRAINT [DF_BasePara_ISOverTime] DEFAULT (0) FOR [ISOverTime],
	CONSTRAINT [DF_BasePara_OverTime] DEFAULT (60) FOR [OverTime],
	CONSTRAINT [DF_BasePara_WorkDayLong] DEFAULT (480) FOR [WorkDayLong],
	CONSTRAINT [DF__BasePara__WOverT__1C873BEC] DEFAULT (1) FOR [WOverTime],
	CONSTRAINT [DF__BasePara__HOverT__1D7B6025] DEFAULT (1) FOR [HOverTime],
	CONSTRAINT [DF__BasePara__FOverT__1E6F845E] DEFAULT (1) FOR [FOverTime],
	CONSTRAINT [DF_BasePara_OvertimeIn] DEFAULT (0) FOR [OvertimeIn]
GO

ALTER TABLE [dbo].[CheckLog] ADD 
	CONSTRAINT [DF_CheckLog_DTime] DEFAULT (getdate()) FOR [DTime]
GO

ALTER TABLE [dbo].[Checkinout] ADD 
	CONSTRAINT [DF_Checkinout_CheckTime] DEFAULT (getdate()) FOR [CheckTime],
	CONSTRAINT [DF_Checkinout_CheckType] DEFAULT ('I') FOR [CheckType],
	CONSTRAINT [IX_Checkinout] UNIQUE  NONCLUSTERED 
	(
		[Userid],
		[CheckTime]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Dept] ADD 
	CONSTRAINT [DF_Dept_SupDeptid] DEFAULT (0) FOR [SupDeptid],
	CONSTRAINT [IX_Dept] UNIQUE  NONCLUSTERED 
	(
		[SupDeptid],
		[DeptName]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FingerClient] ADD 
	CONSTRAINT [DF_FingerClient_Linkmode] DEFAULT (1) FOR [Linkmode],
	CONSTRAINT [DF__FingerCli__RecSt__693CA210] DEFAULT (0) FOR [RecStatus],
	CONSTRAINT [DF_FingerClient_Baudrate] DEFAULT (5) FOR [Baudrate],
	CONSTRAINT [DF_FingerClient_Floorid] DEFAULT (0) FOR [Floorid],
	CONSTRAINT [DF_FingerClient_MachineType] DEFAULT (1) FOR [MachineType],
	CONSTRAINT [DF__FingerCli__Devic__5165187F] DEFAULT (0) FOR [DeviceType],
	CONSTRAINT [DF__FingerCli__CommP__52593CB8] DEFAULT (33302) FOR [CommPort]
GO

ALTER TABLE [dbo].[LeaveClass] ADD 
	CONSTRAINT [DF_LeaveClass_ViewColor] DEFAULT (0) FOR [ViewColor],
	CONSTRAINT [IX_LeaveClass] UNIQUE  NONCLUSTERED 
	(
		[Classname]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[MemStat] ADD 
	CONSTRAINT [DF__MemStat__Normal__10566F31] DEFAULT (0) FOR [Normal],
	CONSTRAINT [DF__MemStat__Actual__114A936A] DEFAULT (0) FOR [Actual],
	CONSTRAINT [DF__MemStat__Latetim__123EB7A3] DEFAULT (0) FOR [Latetime],
	CONSTRAINT [DF__MemStat__Earlyti__1332DBDC] DEFAULT (0) FOR [Earlytime],
	CONSTRAINT [DF__MemStat__Absentt__14270015] DEFAULT (0) FOR [Absenttime],
	CONSTRAINT [DF__MemStat__Overtim__151B244E] DEFAULT (0) FOR [Overtime],
	CONSTRAINT [DF__MemStat__Noin__160F4887] DEFAULT (0) FOR [Noin],
	CONSTRAINT [DF__MemStat__Noout__17036CC0] DEFAULT (0) FOR [Noout],
	CONSTRAINT [DF__MemStat__Awaytim__17F790F9] DEFAULT (0) FOR [Awaytime],
	CONSTRAINT [DF__MemStat__BLeave__18EBB532] DEFAULT (0) FOR [BLeave],
	CONSTRAINT [DF__MemStat__Leave__19DFD96B] DEFAULT (0) FOR [Leave],
	CONSTRAINT [DF__MemStat__Freeove__1AD3FDA4] DEFAULT (0) FOR [Freeovertime],
	CONSTRAINT [DF__MemStat__Worktim__1BC821DD] DEFAULT (0) FOR [Worktime],
	CONSTRAINT [DF__MemStat__Attrate__1CBC4616] DEFAULT (0) FOR [Attrate]
GO

ALTER TABLE [dbo].[OPLog] ADD 
	CONSTRAINT [DF_OPLog_Optime] DEFAULT (getdate()) FOR [Optime]
GO

ALTER TABLE [dbo].[OPinfo] ADD 
	CONSTRAINT [DF_OPinfo_Authority] DEFAULT (N'NNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN') FOR [Authority]
GO

ALTER TABLE [dbo].[SchTime] ADD 
	CONSTRAINT [DF_SchTime_BeginDay] DEFAULT (0) FOR [BeginDay],
	CONSTRAINT [DF_SchTime_Timeid] DEFAULT (0) FOR [Timeid]
GO

ALTER TABLE [dbo].[Schedule] ADD 
	CONSTRAINT [DF_Schedule_Cycles] DEFAULT (1) FOR [Cycles],
	CONSTRAINT [DF_Schedule_Units] DEFAULT (1) FOR [Units]
GO

ALTER TABLE [dbo].[StatItems] ADD 
	CONSTRAINT [DF_StatItems_Units] DEFAULT (1) FOR [Units],
	CONSTRAINT [DF_StatItems_MinUnit] DEFAULT (1) FOR [MinUnit],
	CONSTRAINT [DF_StatItems_SRControl] DEFAULT (2) FOR [SRControl],
	CONSTRAINT [DF_StatItems_IsAddup] DEFAULT (1) FOR [IsAddup],
	CONSTRAINT [DF_StatItems_IsTimes] DEFAULT (0) FOR [IsTimes]
GO

ALTER TABLE [dbo].[T_Checkinout] ADD 
	CONSTRAINT [DF__T_Checkin__Check__5535A963] DEFAULT (0) FOR [Checked],
	CONSTRAINT [DF__T_Checkin__WorkT__5629CD9C] DEFAULT (0) FOR [WorkType],
	CONSTRAINT [DF__T_Checkin__AttFl__571DF1D5] DEFAULT (0) FOR [AttFlag]
GO

ALTER TABLE [dbo].[TimeTable] ADD 
	CONSTRAINT [DF_TimeTable_Latetime] DEFAULT (0) FOR [Latetime],
	CONSTRAINT [DF_TimeTable_Leavetime] DEFAULT (0) FOR [Leavetime],
	CONSTRAINT [DF_TimeTable_WorkDays] DEFAULT (1) FOR [WorkDays],
	CONSTRAINT [DF_TimeTable_Longtime] DEFAULT (480) FOR [Longtime],
	CONSTRAINT [DF_TimeTable_MustIn] DEFAULT (1) FOR [MustIn],
	CONSTRAINT [DF_TimeTable_MustOut] DEFAULT (1) FOR [MustOut]
GO

ALTER TABLE [dbo].[UserLeave] ADD 
	CONSTRAINT [DF_UserLeave_LeaveClassid] DEFAULT (1) FOR [LeaveClassid]
GO

 CREATE  UNIQUE  INDEX [IX_UserLeave] ON [dbo].[UserLeave]([Userid], [BeginTime]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[UserTempShift] ADD 
	CONSTRAINT [DF_UserTempShift_IsOvertime] DEFAULT (0) FOR [IsOvertime]
GO

ALTER TABLE [dbo].[Userinfo] ADD 
	CONSTRAINT [DF_Userinfo_Deptid] DEFAULT (1) FOR [Deptid],
	CONSTRAINT [DF_Userinfo_IsAtt] DEFAULT (1) FOR [IsAtt],
	CONSTRAINT [DF_Userinfo_Isovertime] DEFAULT (1) FOR [Isovertime],
	CONSTRAINT [DF_Userinfo_Isrest] DEFAULT (1) FOR [Isrest],
	CONSTRAINT [DF_Userinfo_MgFlag] DEFAULT (0) FOR [MgFlag],
	CONSTRAINT [DF_Userinfo_UserFlag] DEFAULT (2) FOR [UserFlag],
	CONSTRAINT [DF__Userinfo__workda__0D7A0286] DEFAULT (0) FOR [workdaylong]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW V_Class AS 
SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units, Schedule.AutoClass, SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, 
TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, 
TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime
FROM (Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_Record
AS
SELECT dbo.Checkinout.Logid, dbo.Checkinout.Userid, dbo.Checkinout.CheckTime, 
      dbo.Checkinout.CheckType, dbo.Checkinout.Sensorid, dbo.Checkinout.WorkType, 
      dbo.Checkinout.AttFlag, dbo.Userinfo.Name, dbo.Userinfo.Deptid, dbo.Userinfo.Duty, 
      dbo.Dept.DeptName, dbo.FingerClient.Clientid, dbo.FingerClient.ClientName, 
      dbo.Status.StatusText
FROM dbo.Checkinout LEFT OUTER JOIN
      dbo.Userinfo ON dbo.Checkinout.Userid = dbo.Userinfo.Userid LEFT OUTER JOIN
      dbo.Dept ON dbo.Userinfo.Deptid = dbo.Dept.Deptid LEFT OUTER JOIN
      dbo.FingerClient ON 
      dbo.Checkinout.Sensorid = dbo.FingerClient.ClientNumber LEFT OUTER JOIN
      dbo.Status ON dbo.Checkinout.CheckType = dbo.Status.StatusChar

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.V_UserClient
AS
SELECT dbo.UserPower.Logid, dbo.UserPower.Userid, dbo.UserPower.ClientNumber, 
      dbo.UserPower.PowerFlag, dbo.Userinfo.Name, dbo.Userinfo.Pwd, 
      dbo.Userinfo.CardNum, dbo.Userinfo.Deptid, dbo.Userinfo.UserFlag, 
      dbo.Userinfo.MgFlag, dbo.Userinfo.Groupid
FROM dbo.UserPower INNER JOIN
      dbo.Userinfo ON dbo.UserPower.Userid = dbo.Userinfo.Userid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

