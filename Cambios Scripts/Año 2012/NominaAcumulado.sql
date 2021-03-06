USE [SistemaNominasUrmosa]
GO
/****** Objeto:  Table [dbo].[Nomina]    Fecha de la secuencia de comandos: 02/13/2012 20:51:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[NominaAcumulada](
	[NumNomina] [int] NOT NULL,
	[CodTipoNomina] [nvarchar](10) NULL,
	[TotalSalarioBasico] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalSalarioBasico]  DEFAULT (0),
	[TotalDestajo] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalDestajo]  DEFAULT (0),
	[TotalHorasExtras] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalHorasExtras]  DEFAULT (0),
	[TotalComisiones] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalComisiones]  DEFAULT (0),
	[TotalIncentivos] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalIncentivos]  DEFAULT (0),
	[TotalDeducciones] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalDeducciones]  DEFAULT (0),
	[TotalOtrosIngresos] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalOtrosIngresos]  DEFAULT (0),
	[TotalPrestamo] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalPrestamo]  DEFAULT (0),
	[TotalMontoINSS] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalMontoINSS]  DEFAULT (0),
	[TotalMontoIR] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalMontoIR]  DEFAULT (0),
	[TotalVacaciones] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalVacaciones]  DEFAULT (0),
	[TotalINSSPatronal] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalINSSPatronal]  DEFAULT (0),
	[TotalIRPatronal] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalIRPatronal]  DEFAULT (0),
	[TotalINATEC] [float] NULL CONSTRAINT [DF_NominaAcumulada_TotalINATEC]  DEFAULT (0),
	[Totalmes13] [float] NULL CONSTRAINT [DF_NominaAcumulada_Totalmes13]  DEFAULT (0),
	[FechaNominaINI] [smalldatetime] NULL CONSTRAINT [DF_NominaAcumulada_FechaNominaINI]  DEFAULT (0),
	[FechaNomina] [smalldatetime] NULL CONSTRAINT [DF_NominaAcumulada_FechaNomina]  DEFAULT (0),
	[Activa] [bit] NOT NULL CONSTRAINT [DF_NominaAcumulada_Activa]  DEFAULT (0),
	[Procesada] [bit] NOT NULL CONSTRAINT [DF_NominaAcumulada_Procesada]  DEFAULT (0),
	[Cerrada] [bit] NOT NULL CONSTRAINT [DF_NominaAcumulada_Cerrada]  DEFAULT (0),
	[Anulada] [bit] NOT NULL CONSTRAINT [DF_NominaAcumulada_Anulada]  DEFAULT (0),
	[Mes] [numeric](18, 0) NULL,
	[Ano] [numeric](18, 0) NULL,
	[Periodo] [numeric](18, 0) NULL,
 CONSTRAINT [PK_NominaAcumulada] PRIMARY KEY CLUSTERED 
(
	[NumNomina] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
ALTER TABLE [dbo].[NominaAcumulada]  WITH CHECK ADD  CONSTRAINT [FK_NominaAcumulada_TipoNomina] FOREIGN KEY([CodTipoNomina])
REFERENCES [dbo].[TipoNomina] ([CodTipoNomina])
ON UPDATE CASCADE
ON DELETE CASCADE
GO
ALTER TABLE [dbo].[NominaAcumulada] CHECK CONSTRAINT [FK_NominaAcumulada_TipoNomina]