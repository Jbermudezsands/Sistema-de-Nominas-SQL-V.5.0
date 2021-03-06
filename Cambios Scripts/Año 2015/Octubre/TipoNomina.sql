/*
   viernes, 15 de enero de 201612:36:28 p.m.
   Usuario: sa
   Servidor: 192.168.1.2\SQL2005
   Base de datos: SistemaNominasDistelsa
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
ALTER TABLE dbo.TipoNomina
	DROP CONSTRAINT DF_TipoNomina_PorcientoInss
GO
ALTER TABLE dbo.TipoNomina
	DROP CONSTRAINT DF_TipoNomina_TasaInss
GO
ALTER TABLE dbo.TipoNomina
	DROP CONSTRAINT DF_TipoNomina_PorcientoIr
GO
ALTER TABLE dbo.TipoNomina
	DROP CONSTRAINT DF_TipoNomina_TasaIr
GO
ALTER TABLE dbo.TipoNomina
	DROP CONSTRAINT DF_TipoNomina_TasaInssPatronal
GO
ALTER TABLE dbo.TipoNomina
	DROP CONSTRAINT DF_TipoNomina_CalcularHoraTrabajada
GO
ALTER TABLE dbo.TipoNomina
	DROP CONSTRAINT DF_TipoNomina_IrUltimaSemana
GO
CREATE TABLE dbo.Tmp_TipoNomina
	(
	CodTipoNomina nvarchar(10) NOT NULL,
	Nomina nvarchar(25) NULL,
	Periodo nvarchar(25) NULL,
	UltFecha smalldatetime NULL,
	TipoPago nvarchar(35) NULL,
	Moneda nvarchar(2) NULL,
	MantValor bit NOT NULL,
	Activa bit NOT NULL,
	PorcientoInss bit NULL,
	TasaInss decimal(18, 4) NULL,
	PorcientoIr bit NULL,
	TasaIr decimal(18, 4) NULL,
	TasaInssPatronal decimal(18, 4) NULL,
	HEntrada datetime NULL,
	HSalida datetime NULL,
	TarifaHoraria float(53) NULL,
	CalcularHoraTrabajada bit NULL,
	IrUltimaSemana bit NULL,
	Horas float(53) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Tmp_TipoNomina ADD CONSTRAINT
	DF_TipoNomina_PorcientoInss DEFAULT ((0)) FOR PorcientoInss
GO
ALTER TABLE dbo.Tmp_TipoNomina ADD CONSTRAINT
	DF_TipoNomina_TasaInss DEFAULT ((0)) FOR TasaInss
GO
ALTER TABLE dbo.Tmp_TipoNomina ADD CONSTRAINT
	DF_TipoNomina_PorcientoIr DEFAULT ((0)) FOR PorcientoIr
GO
ALTER TABLE dbo.Tmp_TipoNomina ADD CONSTRAINT
	DF_TipoNomina_TasaIr DEFAULT ((0)) FOR TasaIr
GO
ALTER TABLE dbo.Tmp_TipoNomina ADD CONSTRAINT
	DF_TipoNomina_TasaInssPatronal DEFAULT ((0)) FOR TasaInssPatronal
GO
ALTER TABLE dbo.Tmp_TipoNomina ADD CONSTRAINT
	DF_TipoNomina_CalcularHoraTrabajada DEFAULT ((0)) FOR CalcularHoraTrabajada
GO
ALTER TABLE dbo.Tmp_TipoNomina ADD CONSTRAINT
	DF_TipoNomina_IrUltimaSemana DEFAULT ((0)) FOR IrUltimaSemana
GO
ALTER TABLE dbo.Tmp_TipoNomina ADD CONSTRAINT
	DF_TipoNomina_Horas DEFAULT 8 FOR Horas
GO
IF EXISTS(SELECT * FROM dbo.TipoNomina)
	 EXEC('INSERT INTO dbo.Tmp_TipoNomina (CodTipoNomina, Nomina, Periodo, UltFecha, TipoPago, Moneda, MantValor, Activa, PorcientoInss, TasaInss, PorcientoIr, TasaIr, TasaInssPatronal, HEntrada, HSalida, TarifaHoraria, CalcularHoraTrabajada, IrUltimaSemana)
		SELECT CodTipoNomina, Nomina, Periodo, UltFecha, TipoPago, Moneda, MantValor, Activa, PorcientoInss, TasaInss, PorcientoIr, TasaIr, TasaInssPatronal, HEntrada, HSalida, TarifaHoraria, CalcularHoraTrabajada, IrUltimaSemana FROM dbo.TipoNomina WITH (HOLDLOCK TABLOCKX)')
GO
ALTER TABLE dbo._ActCuentas
	DROP CONSTRAINT FK__ActCuentas_TipoNomina
GO
ALTER TABLE dbo.Nomina
	DROP CONSTRAINT FK_Nomina_TipoNomina
GO
ALTER TABLE dbo.ActividadesCuenta
	DROP CONSTRAINT FK_TareasCuenta_TipoNomina
GO
ALTER TABLE dbo.Empleado
	DROP CONSTRAINT FK_Empleado_TipoNomina
GO
ALTER TABLE dbo.NominaAcumulada
	DROP CONSTRAINT FK_NominaAcumulada_TipoNomina
GO
ALTER TABLE dbo.ActividadesProgramacion
	DROP CONSTRAINT FK_TareasPrograma_TipoNomina
GO
ALTER TABLE dbo.AsistenciaDiaria
	DROP CONSTRAINT FK_AsistenciaDiaria_TipoNomina
GO
DROP TABLE dbo.TipoNomina
GO
EXECUTE sp_rename N'dbo.Tmp_TipoNomina', N'TipoNomina', 'OBJECT' 
GO
ALTER TABLE dbo.TipoNomina ADD CONSTRAINT
	PK_TipoNomina PRIMARY KEY CLUSTERED 
	(
	CodTipoNomina
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.AsistenciaDiaria ADD CONSTRAINT
	FK_AsistenciaDiaria_TipoNomina FOREIGN KEY
	(
	CodTipoNomina
	) REFERENCES dbo.TipoNomina
	(
	CodTipoNomina
	) ON UPDATE  NO ACTION 
	 ON DELETE  NO ACTION 
	
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.ActividadesProgramacion ADD CONSTRAINT
	FK_TareasPrograma_TipoNomina FOREIGN KEY
	(
	TipoNomina
	) REFERENCES dbo.TipoNomina
	(
	CodTipoNomina
	) ON UPDATE  CASCADE 
	 ON DELETE  NO ACTION 
	
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.NominaAcumulada ADD CONSTRAINT
	FK_NominaAcumulada_TipoNomina FOREIGN KEY
	(
	CodTipoNomina
	) REFERENCES dbo.TipoNomina
	(
	CodTipoNomina
	) ON UPDATE  CASCADE 
	 ON DELETE  CASCADE 
	
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.Empleado ADD CONSTRAINT
	FK_Empleado_TipoNomina FOREIGN KEY
	(
	CodTipoNomina
	) REFERENCES dbo.TipoNomina
	(
	CodTipoNomina
	) ON UPDATE  NO ACTION 
	 ON DELETE  NO ACTION 
	
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.ActividadesCuenta ADD CONSTRAINT
	FK_TareasCuenta_TipoNomina FOREIGN KEY
	(
	TipoNomina
	) REFERENCES dbo.TipoNomina
	(
	CodTipoNomina
	) ON UPDATE  CASCADE 
	 ON DELETE  CASCADE 
	
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo.Nomina ADD CONSTRAINT
	FK_Nomina_TipoNomina FOREIGN KEY
	(
	CodTipoNomina
	) REFERENCES dbo.TipoNomina
	(
	CodTipoNomina
	) ON UPDATE  CASCADE 
	 ON DELETE  CASCADE 
	
GO
COMMIT
BEGIN TRANSACTION
GO
ALTER TABLE dbo._ActCuentas ADD CONSTRAINT
	FK__ActCuentas_TipoNomina FOREIGN KEY
	(
	IdNomina
	) REFERENCES dbo.TipoNomina
	(
	CodTipoNomina
	) ON UPDATE  NO ACTION 
	 ON DELETE  NO ACTION 
	
GO
COMMIT
