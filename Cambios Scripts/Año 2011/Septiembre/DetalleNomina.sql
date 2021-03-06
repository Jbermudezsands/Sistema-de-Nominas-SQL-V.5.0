/*
   domingo, 18 de septiembre de 201110:29:11 p.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaNominasNorteak
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
ALTER TABLE dbo.DetalleNomina ADD
	TIngresos float(53) NULL,
	TGastos float(53) NULL
GO
ALTER TABLE dbo.DetalleNomina ADD CONSTRAINT
	DF_DetalleNomina_TIngresos DEFAULT 0 FOR TIngresos
GO
ALTER TABLE dbo.DetalleNomina ADD CONSTRAINT
	DF_DetalleNomina_TGastos DEFAULT 0 FOR TGastos
GO
COMMIT
