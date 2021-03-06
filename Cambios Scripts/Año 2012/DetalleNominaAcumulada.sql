USE [SistemaNominasUrmosa]
GO
/****** Objeto:  Table [dbo].[DetalleNomina]    Fecha de la secuencia de comandos: 02/13/2012 21:01:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[DetalleNominaAcumulada](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[NumNomina] [int] NULL,
	[CodEmpleado] [numeric](18, 0) NULL CONSTRAINT [DF_DetalleNominaAcumulada_CodEmpleado]  DEFAULT (0),
	[SalarioBasico] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_SalarioBasico]  DEFAULT (0),
	[Destajo] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_Destajo]  DEFAULT (0),
	[HE] [real] NULL CONSTRAINT [DF_DetalleNominaAcumulada_HE]  DEFAULT (0),
	[DD] [real] NULL CONSTRAINT [DF_DetalleNominaAcumulada_DD]  DEFAULT (0),
	[HorasExtras] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_HorasExtras]  DEFAULT (0),
	[Comisiones] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_Comisiones]  DEFAULT (0),
	[OtrosIngresos] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_OtrosIngresos]  DEFAULT (0),
	[DescripOtrIngre] [nvarchar](20) NULL,
	[Incentivos] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_Incentivos]  DEFAULT (0),
	[Deducciones] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_Deducciones]  DEFAULT (0),
	[Prestamo] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_Prestamo]  DEFAULT (0),
	[MontoINSS] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_MontoINSS]  DEFAULT (0),
	[MontoIR] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_MontoIR]  DEFAULT (0),
	[Vacaciones] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_Vacaciones]  DEFAULT (0),
	[INSSPatronal] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_INSSPatronal]  DEFAULT (0),
	[IRPatronal] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_IRPatronal]  DEFAULT (0),
	[INATEC] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_INATEC]  DEFAULT (0),
	[Mes13] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_Mes13]  DEFAULT (0),
	[DiasDescuento] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_DiasDescuento]  DEFAULT (0),
	[Adelantos] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_Adelantos]  DEFAULT (0),
	[TotalSubsidio] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_TotalSubsidio]  DEFAULT (0),
	[VacacionesPagadas] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_VacacionesPagadas]  DEFAULT (0),
	[DiasVacaciones] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_DiasVacaciones]  DEFAULT (0),
	[AdelantosVacaciones] [int] NULL CONSTRAINT [DF_DetalleNominaAcumulada_AdelantosVacaciones]  DEFAULT (0),
	[HTrabajada] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_HTrabajada]  DEFAULT (0),
	[SeptimoDia] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_SeptimoDia]  DEFAULT (0),
	[IncetivoProduccion] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_IncetivoProduccion]  DEFAULT (0),
	[TarifaHoraria] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_TarifaHoraria]  DEFAULT (0),
	[produjo] [nvarchar](1) NULL,
	[BonoProduccion] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_BonoProduccion]  DEFAULT (0),
	[Viaticos] [float] NULL,
	[Ajuste] [float] NULL,
	[TIngresos] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_TIngresos]  DEFAULT ((0)),
	[TGastos] [float] NULL CONSTRAINT [DF_DetalleNominaAcumulada_TGastos]  DEFAULT ((0)),
 CONSTRAINT [PK_DetalleNominaAcumulada] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
ALTER TABLE [dbo].[DetalleNominaAcumulada]  WITH NOCHECK ADD  CONSTRAINT [FK_DetalleNominaAcumulada_Empleado] FOREIGN KEY([CodEmpleado])
REFERENCES [dbo].[Empleado] ([CodEmpleado])
ON UPDATE CASCADE
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[DetalleNominaAcumulada] CHECK CONSTRAINT [FK_DetalleNominaAcumulada_Empleado]
GO
ALTER TABLE [dbo].[DetalleNominaAcumulada]  WITH NOCHECK ADD  CONSTRAINT [FK_DetalleNominaAcumulada_Nomina] FOREIGN KEY([NumNomina])
REFERENCES [dbo].[NominaAcumulada] ([NumNomina])
ON UPDATE CASCADE
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[DetalleNominaAcumulada] CHECK CONSTRAINT [FK_DetalleNominaAcumulada_Nomina]