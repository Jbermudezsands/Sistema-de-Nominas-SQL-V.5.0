/*
   viernes, 15 de enero de 201612:39:41 p.m.
   Usuario: sa
   Servidor: 192.168.1.2\SQL2005
   Base de datos: SistemaNominasDistelsa
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
ALTER TABLE dbo.DetalleNomina ADD
	Antiguedad numeric(18, 2) NULL,
	AñoAntiguedad float(53) NULL
GO
ALTER TABLE dbo.DetalleNomina ADD CONSTRAINT
	DF_DetalleNomina_Antiguedad DEFAULT 0 FOR Antiguedad
GO
ALTER TABLE dbo.DetalleNomina ADD CONSTRAINT
	DF_DetalleNomina_AñoAntiguedad DEFAULT 0 FOR AñoAntiguedad
GO
COMMIT
