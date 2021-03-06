/*
   miércoles, 26 de diciembre de 201809:53:02 a.m.
   Usuario: 
   Servidor: JUANBERMUDEZ\SQL2014
   Base de datos: SNTextilesValidos
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
	Reembolso float(53) NULL
GO
ALTER TABLE dbo.Empleado ADD CONSTRAINT
	DF_Empleado_Reembolso DEFAULT 0 FOR Reembolso
GO
ALTER TABLE dbo.Empleado SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
select Has_Perms_By_Name(N'dbo.Empleado', 'Object', 'ALTER') as ALT_Per, Has_Perms_By_Name(N'dbo.Empleado', 'Object', 'VIEW DEFINITION') as View_def_Per, Has_Perms_By_Name(N'dbo.Empleado', 'Object', 'CONTROL') as Contr_Per 