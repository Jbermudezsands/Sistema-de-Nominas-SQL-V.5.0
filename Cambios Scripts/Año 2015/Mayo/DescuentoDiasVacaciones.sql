/*
   miércoles, 27 de mayo de 201503:45:29 p.m.
   Usuario: sa
   Servidor: JUAN\SQL2012
   Base de datos: SistemasNominasEmtrides
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
ALTER TABLE dbo.DescuentoDiasVacaciones ADD
	NumeroSolicitud nvarchar(50) NULL
GO
ALTER TABLE dbo.DescuentoDiasVacaciones SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
