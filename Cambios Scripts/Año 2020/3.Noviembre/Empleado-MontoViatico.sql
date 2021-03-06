/*
   jueves, 05 de noviembre de 202011:59:25 a.m.
   Usuario: sa
   Servidor: RRHH2\SQL2014
   Base de datos: SistemaNominaGraceFashion
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
	MontoViatico float(53) NULL
GO
ALTER TABLE dbo.Empleado ADD CONSTRAINT
	DF_Empleado_MontoViatico DEFAULT 0 FOR MontoViatico
GO
ALTER TABLE dbo.Empleado SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
