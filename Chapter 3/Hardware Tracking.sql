/****** Object:  Database Hardware Tracking    Script Date: 09/24/2000 9:40:39 AM ******/
CREATE DATABASE [Hardware Tracking]  ON (NAME = N'Hardware Tracking_Data', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL$SQL2000\Data\Hardware Tracking_Data.MDF' , SIZE = 5, FILEGROWTH = 10%) LOG ON (NAME = N'Hardware Tracking_Log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL$SQL2000\Data\Hardware Tracking_Log.LDF' , SIZE = 1, FILEGROWTH = 10%)
 COLLATE SQL_Latin1_General_CP1_CI_AS
GO

exec sp_dboption N'Hardware Tracking', N'autoclose', N'false'
GO

exec sp_dboption N'Hardware Tracking', N'bulkcopy', N'false'
GO

exec sp_dboption N'Hardware Tracking', N'trunc. log', N'true'
GO

exec sp_dboption N'Hardware Tracking', N'torn page detection', N'true'
GO

exec sp_dboption N'Hardware Tracking', N'read only', N'false'
GO

exec sp_dboption N'Hardware Tracking', N'dbo use', N'false'
GO

exec sp_dboption N'Hardware Tracking', N'single', N'false'
GO

exec sp_dboption N'Hardware Tracking', N'autoshrink', N'true'
GO

exec sp_dboption N'Hardware Tracking', N'ANSI null default', N'false'
GO

exec sp_dboption N'Hardware Tracking', N'recursive triggers', N'false'
GO

exec sp_dboption N'Hardware Tracking', N'ANSI nulls', N'false'
GO

exec sp_dboption N'Hardware Tracking', N'concat null yields null', N'false'
GO

exec sp_dboption N'Hardware Tracking', N'cursor close on commit', N'false'
GO

exec sp_dboption N'Hardware Tracking', N'default to local cursor', N'false'
GO

exec sp_dboption N'Hardware Tracking', N'quoted identifier', N'false'
GO

exec sp_dboption N'Hardware Tracking', N'ANSI warnings', N'false'
GO

exec sp_dboption N'Hardware Tracking', N'auto create statistics', N'true'
GO

exec sp_dboption N'Hardware Tracking', N'auto update statistics', N'true'
GO

use [Hardware Tracking]
GO

/****** Object:  Table [dbo].[CD_T]    Script Date: 09/24/2000 9:40:39 AM ******/
CREATE TABLE [dbo].[CD_T] (
	[CD_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[CD_Type_CH] [char] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Last_Update_DT] [datetime] NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Employee_T]    Script Date: 09/24/2000 9:40:40 AM ******/
CREATE TABLE [dbo].[Employee_T] (
	[Employee_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Location_ID] [int] NULL ,
	[First_Name_VC] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Last_Name_VC] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Phone_Number_VC] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Last_Update_DT] [datetime] NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Hardware_Notes_T]    Script Date: 09/24/2000 9:40:41 AM ******/
CREATE TABLE [dbo].[Hardware_Notes_T] (
	[Hardware_Notes_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Hardware_ID] [int] NOT NULL ,
	[Hardware_Notes_TX] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Last_Update_DT] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Hardware_T]    Script Date: 09/24/2000 9:40:41 AM ******/
CREATE TABLE [dbo].[Hardware_T] (
	[Hardware_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[CD_ID] [int] NOT NULL ,
	[Manufacturer_VC] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Model_VC] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Processor_Speed_VC] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Memory_VC] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[HardDrive_VC] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Sound_Card_VC] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Speakers_VC] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Video_Card_VC] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Monitor_VC] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Serial_Number_VC] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Lease_Expiration_DT] [datetime] NOT NULL ,
	[Last_Update_DT] [datetime] NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Location_T]    Script Date: 09/24/2000 9:40:42 AM ******/
CREATE TABLE [dbo].[Location_T] (
	[Location_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Location_Name_VC] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Last_Update_DT] [datetime] NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Software_Category_T]    Script Date: 09/24/2000 9:40:42 AM ******/
CREATE TABLE [dbo].[Software_Category_T] (
	[Software_Category_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Software_Category_VC] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Last_Update_DT] [datetime] NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[Software_T]    Script Date: 09/24/2000 9:40:43 AM ******/
CREATE TABLE [dbo].[Software_T] (
	[Software_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Software_Category_ID] [int] NOT NULL ,
	[Software_Name_VC] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Last_Update_DT] [datetime] NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[System_Assignment_T]    Script Date: 09/24/2000 9:40:43 AM ******/
CREATE TABLE [dbo].[System_Assignment_T] (
	[System_Assignment_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[Employee_ID] [int] NOT NULL ,
	[Hardware_ID] [int] NOT NULL ,
	[Last_Update_DT] [datetime] NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[System_Software_Relationship_T]    Script Date: 09/24/2000 9:40:44 AM ******/
CREATE TABLE [dbo].[System_Software_Relationship_T] (
	[System_Assignment_ID] [int] NOT NULL ,
	[Software_ID] [int] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[CD_T] WITH NOCHECK ADD 
	CONSTRAINT [PK_CD_T] PRIMARY KEY  CLUSTERED 
	(
		[CD_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Hardware_Notes_T] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[Hardware_Notes_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Hardware_T] WITH NOCHECK ADD 
	CONSTRAINT [PK_Hardware_T] PRIMARY KEY  CLUSTERED 
	(
		[Hardware_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Location_T] WITH NOCHECK ADD 
	CONSTRAINT [PK_Location_T] PRIMARY KEY  CLUSTERED 
	(
		[Location_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Software_Category_T] WITH NOCHECK ADD 
	CONSTRAINT [PK_Software_Category_T] PRIMARY KEY  CLUSTERED 
	(
		[Software_Category_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Software_T] WITH NOCHECK ADD 
	CONSTRAINT [PK_Software_T] PRIMARY KEY  CLUSTERED 
	(
		[Software_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[System_Assignment_T] WITH NOCHECK ADD 
	CONSTRAINT [PK_System_Assignment_T] PRIMARY KEY  CLUSTERED 
	(
		[System_Assignment_ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[System_Software_Relationship_T] WITH NOCHECK ADD 
	CONSTRAINT [PK_System_Software_Relationship_T] PRIMARY KEY  CLUSTERED 
	(
		[System_Assignment_ID],
		[Software_ID]
	)  ON [PRIMARY] 
GO

 CREATE  CLUSTERED  INDEX [IX_Employee_T] ON [dbo].[Employee_T]([Last_Name_VC], [First_Name_VC]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Employee_T] WITH NOCHECK ADD 
	CONSTRAINT [PK_Employee_T] PRIMARY KEY  NONCLUSTERED 
	(
		[Employee_ID]
	)  ON [PRIMARY] 
GO

 CREATE  INDEX [IX_Hardware_T] ON [dbo].[Hardware_T]([Manufacturer_VC]) ON [PRIMARY]
GO

GRANT  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON [dbo].[CD_T]  TO [Hardware Users]
GO

GRANT  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON [dbo].[Employee_T]  TO [Hardware Users]
GO

GRANT  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON [dbo].[Hardware_Notes_T]  TO [Hardware Users]
GO

GRANT  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON [dbo].[Hardware_T]  TO [Hardware Users]
GO

GRANT  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON [dbo].[Location_T]  TO [Hardware Users]
GO

GRANT  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON [dbo].[Software_Category_T]  TO [Hardware Users]
GO

GRANT  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON [dbo].[Software_T]  TO [Hardware Users]
GO

GRANT  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON [dbo].[System_Assignment_T]  TO [Hardware Users]
GO

GRANT  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON [dbo].[System_Software_Relationship_T]  TO [Hardware Users]
GO

ALTER TABLE [dbo].[Employee_T] ADD 
	CONSTRAINT [FK_Employee_T_Location_T] FOREIGN KEY 
	(
		[Location_ID]
	) REFERENCES [dbo].[Location_T] (
		[Location_ID]
	)
GO

ALTER TABLE [dbo].[Hardware_Notes_T] ADD 
	CONSTRAINT [FK_Hardware_Notes_T] FOREIGN KEY 
	(
		[Hardware_ID]
	) REFERENCES [dbo].[Hardware_T] (
		[Hardware_ID]
	) ON DELETE CASCADE 
GO

ALTER TABLE [dbo].[Hardware_T] ADD 
	CONSTRAINT [FK_Hardware_T_CD_T] FOREIGN KEY 
	(
		[CD_ID]
	) REFERENCES [dbo].[CD_T] (
		[CD_ID]
	)
GO

ALTER TABLE [dbo].[Software_T] ADD 
	CONSTRAINT [FK_Software_T_Software_Category_T] FOREIGN KEY 
	(
		[Software_Category_ID]
	) REFERENCES [dbo].[Software_Category_T] (
		[Software_Category_ID]
	)
GO

ALTER TABLE [dbo].[System_Assignment_T] ADD 
	CONSTRAINT [FK_System_Assignment_T_Employee_T] FOREIGN KEY 
	(
		[Employee_ID]
	) REFERENCES [dbo].[Employee_T] (
		[Employee_ID]
	) ON DELETE CASCADE ,
	CONSTRAINT [FK_System_Assignment_T_Hardware_T] FOREIGN KEY 
	(
		[Hardware_ID]
	) REFERENCES [dbo].[Hardware_T] (
		[Hardware_ID]
	)
GO

ALTER TABLE [dbo].[System_Software_Relationship_T] ADD 
	CONSTRAINT [FK_System_Software_Relationship_T_Software_T] FOREIGN KEY 
	(
		[Software_ID]
	) REFERENCES [dbo].[Software_T] (
		[Software_ID]
	),
	CONSTRAINT [FK_System_Software_Relationship_T_System_Assignment_T] FOREIGN KEY 
	(
		[System_Assignment_ID]
	) REFERENCES [dbo].[System_Assignment_T] (
		[System_Assignment_ID]
	) ON DELETE CASCADE 
GO
