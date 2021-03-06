USE [SistemaNominasNorteak]
GO
/****** Objeto:  Table [dbo].[_ActividadesProgramacion]    Fecha de la secuencia de comandos: 02/14/2012 07:30:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[_ActividadesProgramacion](
	[IdActProgramacion] [int] IDENTITY(1,1) NOT NULL,
	[IdActividad] [int] NOT NULL,
	[IdFincaPlantacion] [int] NOT NULL,
	[FechaRegistro] [datetime] NOT NULL,
	[FechaInicio] [datetime] NOT NULL,
	[FechaFin] [datetime] NOT NULL,
	[HPOrdinaria] [int] NOT NULL,
	[HPExtra] [int] NOT NULL,
	[CerrarPlaneacion] [bit] NULL,
 CONSTRAINT [PK__ActividadesProgramacion] PRIMARY KEY CLUSTERED 
(
	[IdActProgramacion] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY],
 CONSTRAINT [IX__ActividadesProgramacion] UNIQUE NONCLUSTERED 
(
	[IdActividad] ASC,
	[IdFincaPlantacion] ASC,
	[FechaInicio] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
ALTER TABLE [dbo].[_ActividadesProgramacion]  WITH CHECK ADD  CONSTRAINT [FK__ActividadesProgramacion__Actividades] FOREIGN KEY([IdActividad])
REFERENCES [dbo].[_Actividades] ([IdActividad])
GO
ALTER TABLE [dbo].[_ActividadesProgramacion] CHECK CONSTRAINT [FK__ActividadesProgramacion__Actividades]
GO
ALTER TABLE [dbo].[_ActividadesProgramacion]  WITH CHECK ADD  CONSTRAINT [FK__ActividadesProgramacion__FincaPlantacion] FOREIGN KEY([IdFincaPlantacion])
REFERENCES [dbo].[_FincaPlantacion] ([IdFincaPlantacion])
GO
ALTER TABLE [dbo].[_ActividadesProgramacion] CHECK CONSTRAINT [FK__ActividadesProgramacion__FincaPlantacion]