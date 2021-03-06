/*
   domingo, 28 de julio de 201301:58:26 p.m.
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
ALTER TABLE dbo.Bajas
	DROP CONSTRAINT FK_Bajas_Empleado
GO
COMMIT
BEGIN TRANSACTION
GO
CREATE TABLE dbo.Tmp_Bajas
	(
	Id int NOT NULL,
	CodEmpleado numeric(18, 0) NULL,
	FechaBaja smalldatetime NULL,
	AnnosTrabajados float(53) NULL,
	MesesTrabajados float(53) NULL,
	DiasTrabajados float(53) NULL,
	MontoNomPropor float(53) NULL,
	MontoVaca float(53) NULL,
	Monto13Mes float(53) NULL,
	MontoAnosTrab float(53) NULL,
	MontoCargoConfianza float(53) NULL,
	MontoAntiguedad float(53) NULL,
	MotivoBaja nvarchar(50) NULL,
	TipoBaja nvarchar(10) NULL,
	Otro nvarchar(25) NULL,
	MontoOtro float(53) NULL,
	Prestamo float(53) NULL,
	Deducciones float(53) NULL,
	SalarioMensual money NULL,
	MontoINSS money NULL,
	MontoIR money NULL,
	FechaIniAgui smalldatetime NULL,
	FechaFinAgui smalldatetime NULL,
	DiasAguinaldo numeric(18, 2) NULL,
	DiasVacaciones numeric(18, 2) NULL,
	DiasMenosVaca numeric(18, 2) NULL,
	FechaIniVaca smalldatetime NULL,
	FechaFinVaca smalldatetime NULL,
	HorasExtra money NULL,
	Viaticos money NULL,
	MontoHorasExtra money NULL
	)  ON [PRIMARY]
GO
IF EXISTS(SELECT * FROM dbo.Bajas)
	 EXEC('INSERT INTO dbo.Tmp_Bajas (Id, CodEmpleado, FechaBaja, AnnosTrabajados, MesesTrabajados, DiasTrabajados, MontoNomPropor, MontoVaca, Monto13Mes, MontoAnosTrab, MontoCargoConfianza, MontoAntiguedad, MotivoBaja, TipoBaja, Otro, MontoOtro, Prestamo, Deducciones, SalarioMensual, MontoINSS, MontoIR, FechaIniAgui, FechaFinAgui, DiasAguinaldo, DiasVacaciones, DiasMenosVaca, FechaIniVaca, FechaFinVaca, HorasExtra, Viaticos, MontoHorasExtra)
		SELECT Id, CodEmpleado, FechaBaja, AnnosTrabajados, MesesTrabajados, DiasTrabajados, MontoNomPropor, MontoVaca, Monto13Mes, MontoAnosTrab, MontoCargoConfianza, MontoAntiguedad, MotivoBaja, TipoBaja, Otro, MontoOtro, Prestamo, Deducciones, SalarioMensual, MontoINSS, MontoIR, FechaIniAgui, FechaFinAgui, CONVERT(numeric(18, 2), DiasAguinaldo), CONVERT(numeric(18, 2), DiasVacaciones), CONVERT(numeric(18, 2), DiasMenosVaca), FechaIniVaca, FechaFinVaca, HorasExtra, Viaticos, MontoHorasExtra FROM dbo.Bajas WITH (HOLDLOCK TABLOCKX)')
GO
DROP TABLE dbo.Bajas
GO
EXECUTE sp_rename N'dbo.Tmp_Bajas', N'Bajas', 'OBJECT' 
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	PK_Bajas PRIMARY KEY CLUSTERED 
	(
	Id
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
ALTER TABLE dbo.Bajas WITH NOCHECK ADD CONSTRAINT
	FK_Bajas_Empleado FOREIGN KEY
	(
	CodEmpleado
	) REFERENCES dbo.Empleado
	(
	CodEmpleado
	) ON UPDATE  CASCADE 
	 ON DELETE  CASCADE 
	
GO
COMMIT
