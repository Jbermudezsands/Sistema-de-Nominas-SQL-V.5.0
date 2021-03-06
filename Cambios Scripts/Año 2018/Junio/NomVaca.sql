/*
   miércoles, 13 de junio de 201804:21:32 p.m.
   Usuario: sa
   Servidor: JUANBERMUDEZ\SQL2014
   Base de datos: SistemaNominasMonteFresco
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
ALTER TABLE dbo.NomVaca ADD
	Transfereir bit NULL
GO
ALTER TABLE dbo.NomVaca ADD CONSTRAINT
	DF_NomVaca_Transfereir DEFAULT 0 FOR Transfereir
GO
ALTER TABLE dbo.NomVaca SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
