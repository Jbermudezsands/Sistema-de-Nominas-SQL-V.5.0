/*
   domingo, 12 de agosto de 201810:46:53 a.m.
   Usuario: sa
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
ALTER TABLE dbo.Bajas ADD
	Calculada bit NULL,
	Procesada bit NULL
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	DF_Bajas_Calculada DEFAULT 1 FOR Calculada
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	DF_Bajas_Procesada DEFAULT 0 FOR Procesada
GO
ALTER TABLE dbo.Bajas SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
