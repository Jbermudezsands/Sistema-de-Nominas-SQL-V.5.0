/*
   lunes, 29 de agosto de 201110:58:47 a.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaNominasMetro
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
COMMIT
BEGIN TRANSACTION
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.Puntos ADD CONSTRAINT
	FK_Puntos_PuntosGrupo FOREIGN KEY
	(
	Grupo
	) REFERENCES dbo.PuntosGrupo
	(
	Id
	) ON UPDATE  NO ACTION 
	 ON DELETE  NO ACTION 
	
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.PuntosEmpleado ADD CONSTRAINT
	FK_PuntosEmpleado_Puntos FOREIGN KEY
	(
	Puntos
	) REFERENCES dbo.Puntos
	(
	Id
	) ON UPDATE  NO ACTION 
	 ON DELETE  NO ACTION 
	
GO
ALTER TABLE dbo.PuntosEmpleado ADD CONSTRAINT
	FK_PuntosEmpleado_Empleado FOREIGN KEY
	(
	Empleado
	) REFERENCES dbo.Empleado
	(
	CodEmpleado
	) ON UPDATE  NO ACTION 
	 ON DELETE  NO ACTION 
	
GO
COMMIT
