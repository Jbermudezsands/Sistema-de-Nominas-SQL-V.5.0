USE [SistemaNominasNorteak]
GO
/****** Objeto:  Table [dbo].[_ActCuentas]    Fecha de la secuencia de comandos: 02/14/2012 07:28:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[_ActCuentas](
	[IdCuenta] [int] IDENTITY(1,1) NOT NULL,
	[IdActividad] [int] NOT NULL,
	[IdNomina] [nvarchar](10) NOT NULL,
	[Cuenta] [nchar](10) NOT NULL,
 CONSTRAINT [PK__ActCuentas] PRIMARY KEY CLUSTERED 
(
	[IdCuenta] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
ALTER TABLE [dbo].[_ActCuentas]  WITH CHECK ADD  CONSTRAINT [FK__ActCuentas__Actividades] FOREIGN KEY([IdActividad])
REFERENCES [dbo].[_Actividades] ([IdActividad])
GO
ALTER TABLE [dbo].[_ActCuentas] CHECK CONSTRAINT [FK__ActCuentas__Actividades]
GO
ALTER TABLE [dbo].[_ActCuentas]  WITH CHECK ADD  CONSTRAINT [FK__ActCuentas_TipoNomina] FOREIGN KEY([IdNomina])
REFERENCES [dbo].[TipoNomina] ([CodTipoNomina])
GO
ALTER TABLE [dbo].[_ActCuentas] CHECK CONSTRAINT [FK__ActCuentas_TipoNomina]