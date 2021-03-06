USE [SistemaNominasNorteak]
GO
/****** Objeto:  Table [dbo].[_FincaPlantacion]    Fecha de la secuencia de comandos: 02/14/2012 07:30:43 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[_FincaPlantacion](
	[IdFincaPlantacion] [int] IDENTITY(1,1) NOT NULL,
	[IdFinca] [int] NOT NULL,
	[IdPlantacion] [int] NOT NULL,
	[Anio] [int] NOT NULL,
 CONSTRAINT [PK_Plantacion] PRIMARY KEY CLUSTERED 
(
	[IdFincaPlantacion] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY],
 CONSTRAINT [IX__FincaPlantacion] UNIQUE NONCLUSTERED 
(
	[IdFinca] ASC,
	[IdPlantacion] ASC,
	[Anio] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
ALTER TABLE [dbo].[_FincaPlantacion]  WITH CHECK ADD  CONSTRAINT [FK__FincaPlantacion__Finca] FOREIGN KEY([IdFinca])
REFERENCES [dbo].[_Finca] ([IdFinca])
GO
ALTER TABLE [dbo].[_FincaPlantacion] CHECK CONSTRAINT [FK__FincaPlantacion__Finca]
GO
ALTER TABLE [dbo].[_FincaPlantacion]  WITH CHECK ADD  CONSTRAINT [FK__FincaPlantacion__Plantacion] FOREIGN KEY([IdPlantacion])
REFERENCES [dbo].[_Plantacion] ([IdPlantacion])
GO
ALTER TABLE [dbo].[_FincaPlantacion] CHECK CONSTRAINT [FK__FincaPlantacion__Plantacion]