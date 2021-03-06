/*
   domingo, 04 de mayo de 201408:35:02 a.m.
   Usuario: sa
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
CREATE TABLE dbo.SolicitudVacaciones
	(
	FechaSolicitud smalldatetime NOT NULL,
	NumeroSolicitud nvarchar(50) NOT NULL,
	TipoSolicitud nvarchar(50) NOT NULL,
	CodigoEmpleado nvarchar(50) NOT NULL,
	DiasVacaciones float(53) NULL,
	DiasDisfrutados float(53) NULL,
	FechaInicio smalldatetime NULL,
	FechaFin smalldatetime NULL,
	DiasDisfrutar float(53) NULL,
	Observaciones nvarchar(MAX) NULL
	)  ON [PRIMARY]
	 TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE dbo.SolicitudVacaciones ADD CONSTRAINT
	PK_SolicitudVacaciones PRIMARY KEY CLUSTERED 
	(
	FechaSolicitud,
	NumeroSolicitud,
	TipoSolicitud,
	CodigoEmpleado
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
