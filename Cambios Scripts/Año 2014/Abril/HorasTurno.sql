/*
   domingo, 06 de abril de 201404:46:34 p.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaNominasYanber
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
CREATE TABLE dbo.HorasTurno
	(
	CodEmpleado numeric(18, 0) NOT NULL,
	NumNomina int NOT NULL,
	CantHoras float(53) NULL,
	Pagada bit NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.HorasTurno ADD CONSTRAINT
	DF_HorasTurno_Pagada DEFAULT 0 FOR Pagada
GO
ALTER TABLE dbo.HorasTurno ADD CONSTRAINT
	PK_HorasTurno PRIMARY KEY CLUSTERED 
	(
	CodEmpleado,
	NumNomina
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
