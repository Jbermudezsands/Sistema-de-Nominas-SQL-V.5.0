USE [SistemaNominasNorteak]
GO
/****** Objeto:  Table [dbo].[_Actividades]    Fecha de la secuencia de comandos: 02/14/2012 07:28:58 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[_Actividades](
	[IdActividad] [int] IDENTITY(1,1) NOT NULL,
	[IdSuperior] [int] NOT NULL,
	[Sufijo] [nchar](10) NOT NULL,
	[Codigo] [nchar](10) NOT NULL,
	[Actividad] [nvarchar](50) NOT NULL,
	[PagaCliente] [bit] NOT NULL,
 CONSTRAINT [PK__Actividad] PRIMARY KEY CLUSTERED 
(
	[IdActividad] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
ALTER TABLE [dbo].[_Actividades]  WITH CHECK ADD  CONSTRAINT [FK__Actividades__Actividades] FOREIGN KEY([IdSuperior])
REFERENCES [dbo].[_Actividades] ([IdActividad])
GO
ALTER TABLE [dbo].[_Actividades] CHECK CONSTRAINT [FK__Actividades__Actividades]