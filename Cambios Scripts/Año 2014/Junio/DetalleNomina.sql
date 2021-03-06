/*
   miércoles, 09 de julio de 201404:27:51 p.m.
   Usuario: sa
   Servidor: JUAN\SQL2005
   Base de datos: SistemaNominaDATATEX
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
	Antiguedad numeric(18, 2) NULL,
	AñoAntiguedad float(53) NULL
GO
ALTER TABLE dbo.DetalleNomina ADD CONSTRAINT
	DF_DetalleNomina_Antiguedad DEFAULT 0 FOR Antiguedad
GO
ALTER TABLE dbo.DetalleNomina ADD CONSTRAINT
	DF_DetalleNomina_AñosAntiguedad DEFAULT 0 FOR AñoAntiguedad
GO
COMMIT
