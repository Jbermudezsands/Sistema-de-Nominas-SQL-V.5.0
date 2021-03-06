/*
   martes, 17 de julio de 201211:40:10 a.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaNominasIpemsa
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
ALTER TABLE dbo.Nomina ADD
	Contabilizado bit NULL,
	Marca bit NULL
GO
ALTER TABLE dbo.Nomina ADD CONSTRAINT
	DF_Nomina_Contabilizado DEFAULT 0 FOR Contabilizado
GO
ALTER TABLE dbo.Nomina ADD CONSTRAINT
	DF_Nomina_Marca DEFAULT 1 FOR Marca
GO
COMMIT
