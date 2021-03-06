/*
   viernes, 26 de octubre de 201812:07:35 p.m.
   Usuario: 
   Servidor: JUANBERMUDEZ\SQL2005
   Base de datos: SistemaNominaEmtrides
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
CREATE TABLE dbo.DeduccionPorciento
	(
	IdDeduccionPorciento int NOT NULL IDENTITY (1, 1),
	Rango1 float(53) NULL,
	Rango2 float(53) NULL,
	Porciento float(53) NULL
	)  ON [PRIMARY]
GO
COMMIT
select Has_Perms_By_Name(N'dbo.DeduccionPorciento', 'Object', 'ALTER') as ALT_Per, Has_Perms_By_Name(N'dbo.DeduccionPorciento', 'Object', 'VIEW DEFINITION') as View_def_Per, Has_Perms_By_Name(N'dbo.DeduccionPorciento', 'Object', 'CONTROL') as Contr_Per 