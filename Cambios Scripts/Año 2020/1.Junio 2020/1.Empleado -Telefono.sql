/*
   viernes, 26 de junio de 202011:44:12 a.m.
   Usuario: 
   Servidor: JUANBERMUDEZ-PC\SQL2014
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
	Telefono nvarchar(50) NULL
GO
ALTER TABLE dbo.Empleado SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
