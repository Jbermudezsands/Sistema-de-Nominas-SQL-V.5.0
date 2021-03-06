USE [SistemaNominasNorteak]
GO
/****** Objeto:  Table [dbo].[_Plantacion]    Fecha de la secuencia de comandos: 02/14/2012 07:31:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[_Plantacion](
	[IdPlantacion] [int] IDENTITY(1,1) NOT NULL,
	[Plantacion] [nvarchar](50) NOT NULL,
 CONSTRAINT [PK__Plantacion] PRIMARY KEY CLUSTERED 
(
	[IdPlantacion] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
