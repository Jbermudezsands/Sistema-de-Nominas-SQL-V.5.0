/*
   Lunes, 23 de Noviembre de 2009 06:54:45 p.m.
   Usuario: 
   Servidor: JUAN\SQL2000
   Base de datos: SistemaNominasMetro
   Aplicación: MS SQLEM - Data Tools
*/

BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
ALTER TABLE dbo.Historico
	DROP CONSTRAINT FK_Historico_Empleado
GO
COMMIT
BEGIN TRANSACTION
CREATE TABLE dbo.Tmp_Historico
	(
	Id int NOT NULL,
	Codempleado numeric(18, 0) NULL,
	FechaBaja smalldatetime NULL,
	MotivoBaja nvarchar(150) NULL,
	FechaAumento smalldatetime NULL,
	MotivoAumento nvarchar(150) NULL,
	FechaInicialSusp smalldatetime NULL,
	FechaFinalSusp smalldatetime NULL,
	MotivoSuspencion nvarchar(150) NULL,
	FechaNacimiento smalldatetime NULL,
	FechaContrato smalldatetime NULL,
	FechaContratoVac smalldatetime NULL,
	CargoInicial nvarchar(35) NULL,
	CargoActual nvarchar(35) NULL,
	CargoAnterior nvarchar(35) NULL,
	SueldoInicial money NOT NULL,
	SueldoAnterior money NOT NULL,
	SueldoActual money NOT NULL,
	CuentaDebito nvarchar(20) NULL,
	CuentaCredito nvarchar(20) NULL
	)  ON [PRIMARY]
GO
IF EXISTS(SELECT * FROM dbo.Historico)
	 EXEC('INSERT INTO dbo.Tmp_Historico (Id, Codempleado, FechaBaja, MotivoBaja, FechaAumento, MotivoAumento, FechaInicialSusp, FechaFinalSusp, MotivoSuspencion, FechaNacimiento, FechaContrato, CargoInicial, CargoActual, CargoAnterior, SueldoInicial, SueldoAnterior, SueldoActual, CuentaDebito, CuentaCredito)
		SELECT Id, Codempleado, FechaBaja, MotivoBaja, FechaAumento, MotivoAumento, FechaInicialSusp, FechaFinalSusp, MotivoSuspencion, FechaNacimiento, FechaContrato, CargoInicial, CargoActual, CargoAnterior, SueldoInicial, SueldoAnterior, SueldoActual, CuentaDebito, CuentaCredito FROM dbo.Historico TABLOCKX')
GO
DROP TABLE dbo.Historico
GO
EXECUTE sp_rename N'dbo.Tmp_Historico', N'Historico', 'OBJECT'
GO
ALTER TABLE dbo.Historico ADD CONSTRAINT
	PK_Historico PRIMARY KEY CLUSTERED 
	(
	Id
	) ON [PRIMARY]

GO
ALTER TABLE dbo.Historico WITH NOCHECK ADD CONSTRAINT
	FK_Historico_Empleado FOREIGN KEY
	(
	Codempleado
	) REFERENCES dbo.Empleado
	(
	CodEmpleado
	) ON UPDATE CASCADE
	 ON DELETE CASCADE
	
GO
COMMIT
