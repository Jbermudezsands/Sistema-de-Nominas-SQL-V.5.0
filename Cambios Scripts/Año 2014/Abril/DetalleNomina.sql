/*
   domingo, 06 de abril de 201402:52:46 p.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaNominasYanber
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
	HTurno decimal(18, 2) NULL,
	HorasTurno decimal(18, 2) NULL
GO
ALTER TABLE dbo.DetalleNomina ADD CONSTRAINT
	DF_DetalleNomina_HTurno DEFAULT 0 FOR HTurno
GO
ALTER TABLE dbo.DetalleNomina ADD CONSTRAINT
	DF_DetalleNomina_HorasTurno DEFAULT 0 FOR HorasTurno
GO
COMMIT
