/*
   lunes, 08 de febrero de 202110:01:12 a.m.
   Usuario: sa
   Servidor: XIOMARA\SQL2005
   Base de datos: SistemaNominas
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
ALTER TABLE dbo.Historico ADD
	CuentaPrestamo nvarchar(20) NULL,
	CuentaOtrosIngresos nvarchar(20) NULL,
	CuentaINSS nvarchar(20) NULL,
	CuentaIR nvarchar(20) NULL,
	CuentaSueldos nvarchar(20) NULL,
	ProvAguinaldo nvarchar(20) NULL,
	ProvVacaciones nvarchar(20) NULL,
	INSSPatronal nvarchar(20) NULL,
	INATEC nvarchar(20) NULL,
	AguinaldoxPagar nvarchar(20) NULL,
	VacacionesxPagar nvarchar(20) NULL,
	INSSxPagar nvarchar(20) NULL,
	INATECxPagar nvarchar(20) NULL,
	IRxPagar nvarchar(20) NULL,
	PrestamoxPagar nvarchar(20) NULL,
	NominaxPagar nvarchar(20) NULL,
	CuentaHorasExtra nvarchar(20) NULL,
	INSSPatronalPagar nvarchar(20) NULL,
	CuentaBanco nvarchar(20) NULL,
	CuentaSubsidio nvarchar(20) NULL
GO
COMMIT
