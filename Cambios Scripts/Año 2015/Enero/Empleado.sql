/*
   martes, 06 de enero de 201509:26:31 a.m.
   Usuario: zeus
   Servidor: JUAN\SQL2012
   Base de datos: SistemaNominasMEDISUT
   Aplicación: 
*/

/* Para evitar posibles problemas de pérdida de datos, debe revisar este script detalladamente antes de ejecutarlo fuera del contexto del diseñador de base de datos.*/
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
ALTER TABLE dbo.Empleado ADD
	DiasBasico float(53) NULL
GO
ALTER TABLE dbo.Empleado ADD CONSTRAINT
	DF_Empleado_DiasBasico DEFAULT 0 FOR DiasBasico
GO
ALTER TABLE dbo.Empleado SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
