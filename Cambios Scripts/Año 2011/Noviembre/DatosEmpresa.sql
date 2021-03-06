/*
   viernes, 04 de noviembre de 201105:34:16 p.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaNominasNorteak
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
ALTER TABLE dbo.DatosEmpresa ADD
	CalcularPuntos bit NULL
GO
ALTER TABLE dbo.DatosEmpresa ADD CONSTRAINT
	DF_DatosEmpresa_CalcularPuntos DEFAULT 0 FOR CalcularPuntos
GO
COMMIT
