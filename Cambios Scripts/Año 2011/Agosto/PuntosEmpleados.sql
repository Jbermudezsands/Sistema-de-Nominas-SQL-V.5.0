/*
   lunes, 29 de agosto de 201110:04:45 a.m.
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
CREATE TABLE dbo.PuntosEmpleado
	(
	Empleado numeric(18, 0) NOT NULL,
	Puntos int NOT NULL,
	Aprobado bit NOT NULL,
	Justificacion nvarchar(200) NOT NULL,
	Documento nvarchar(50) NULL,
	DirDocumento nvarchar(50) NULL,
	FechaSolicitud datetime NOT NULL,
	FechaAprobado datetime NULL,
	Eliminar bit NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.PuntosEmpleado ADD CONSTRAINT
	PK_PuntosEmpleado PRIMARY KEY CLUSTERED 
	(
	Empleado,
	Puntos
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
