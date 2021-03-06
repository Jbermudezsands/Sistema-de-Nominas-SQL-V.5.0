/*
   lunes, 07 de abril de 201402:49:06 p.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaNominasMetro
   Aplicación: 
*/

/* Para evitar posibles problemas de pérdida de datos, debe revisar esta secuencia de comandos detalladamente antes de ejecutarla fuera del contexto del diseñador de base de datos.*/
BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.DetalleNomVaca ADD CONSTRAINT
	DF_DetalleNomVaca_SalarioMensual DEFAULT 0 FOR SalarioMensual
GO
ALTER TABLE dbo.DetalleNomVaca ADD CONSTRAINT
	DF_DetalleNomVaca_DiasAPagar DEFAULT 0 FOR DiasAPagar
GO
ALTER TABLE dbo.DetalleNomVaca ADD CONSTRAINT
	DF_DetalleNomVaca_DiasDescuento DEFAULT 0 FOR DiasDescuento
GO
ALTER TABLE dbo.DetalleNomVaca ADD CONSTRAINT
	DF_DetalleNomVaca_AdelantoVacaciones DEFAULT 0 FOR AdelantoVacaciones
GO
ALTER TABLE dbo.DetalleNomVaca ADD CONSTRAINT
	DF_DetalleNomVaca_Inss DEFAULT 0 FOR Inss
GO
ALTER TABLE dbo.DetalleNomVaca ADD CONSTRAINT
	DF_DetalleNomVaca_TarifaHoraria DEFAULT 0 FOR TarifaHoraria
GO
ALTER TABLE dbo.DetalleNomVaca ADD CONSTRAINT
	DF_DetalleNomVaca_TotalDevengado DEFAULT 0 FOR TotalDevengado
GO
ALTER TABLE dbo.DetalleNomVaca ADD CONSTRAINT
	DF_DetalleNomVaca_Ir DEFAULT 0 FOR Ir
GO
COMMIT
