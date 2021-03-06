USE [SistemaNominasNorteak]
GO
/****** Objeto:  Table [dbo].[_ActividadesProduccion]    Fecha de la secuencia de comandos: 02/14/2012 07:29:33 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[_ActividadesProduccion](
	[IdActProduccion] [int] IDENTITY(1,1) NOT NULL,
	[IdActividad] [int] NOT NULL,
	[CodEmpleado] [numeric](18, 0) NOT NULL,
	[IdFincaPlantacion] [int] NOT NULL,
	[FechaRegistro] [datetime] NOT NULL,
	[Fecha] [datetime] NOT NULL,
	[CantidadHoras] [int] NULL CONSTRAINT [DF__ActividadesProduccion_CantidadHoras]  DEFAULT ((0)),
	[HEntrada] [datetime] NULL,
	[HSalida] [datetime] NULL,
	[HExtras] [bit] NOT NULL,
	[FechaAprobado] [datetime] NULL,
	[Eliminar] [bit] NOT NULL,
 CONSTRAINT [PK__ActividadesProduccion] PRIMARY KEY CLUSTERED 
(
	[IdActProduccion] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
ALTER TABLE [dbo].[_ActividadesProduccion]  WITH CHECK ADD  CONSTRAINT [FK__ActividadesProduccion__Actividades] FOREIGN KEY([IdActividad])
REFERENCES [dbo].[_Actividades] ([IdActividad])
GO
ALTER TABLE [dbo].[_ActividadesProduccion] CHECK CONSTRAINT [FK__ActividadesProduccion__Actividades]
GO
ALTER TABLE [dbo].[_ActividadesProduccion]  WITH CHECK ADD  CONSTRAINT [FK__ActividadesProduccion__FincaPlantacion] FOREIGN KEY([IdFincaPlantacion])
REFERENCES [dbo].[_FincaPlantacion] ([IdFincaPlantacion])
GO
ALTER TABLE [dbo].[_ActividadesProduccion] CHECK CONSTRAINT [FK__ActividadesProduccion__FincaPlantacion]
GO
ALTER TABLE [dbo].[_ActividadesProduccion]  WITH CHECK ADD  CONSTRAINT [FK__ActividadesProduccion_Empleado] FOREIGN KEY([CodEmpleado])
REFERENCES [dbo].[Empleado] ([CodEmpleado])
GO
ALTER TABLE [dbo].[_ActividadesProduccion] CHECK CONSTRAINT [FK__ActividadesProduccion_Empleado]