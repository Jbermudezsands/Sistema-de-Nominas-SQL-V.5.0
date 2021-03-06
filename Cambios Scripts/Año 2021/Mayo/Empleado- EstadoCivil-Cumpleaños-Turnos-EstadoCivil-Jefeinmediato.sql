/*
   miércoles, 26 de mayo de 202116:23:19
   Usuario: 
   Servidor: JUANBERMUDEZ
   Base de datos: SistemaNominaSpinningMills
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
ALTER TABLE dbo.Empleado ADD
	Turno nvarchar(50) NULL,
	Numerocelular nvarchar(50) NULL,
	CelularEmergencia nvarchar(50) NULL,
	Profesion nvarchar(50) NULL,
	EstadoCivil nvarchar(50) NULL,
	JefeInmediato nvarchar(50) NULL,
	Incentivo float(53) NULL
GO
ALTER TABLE dbo.Empleado SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
