/*
   martes, 20 de agosto de 201311:16:03 a.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaNominaNitacsa
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
ALTER TABLE dbo.Controles ADD
	SalarioPromedioReal bit NULL
GO
ALTER TABLE dbo.Controles ADD CONSTRAINT
	DF_Controles_SalarioPromedioReal DEFAULT 0 FOR SalarioPromedioReal
GO
COMMIT
