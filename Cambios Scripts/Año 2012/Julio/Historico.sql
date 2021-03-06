/*
   domingo, 29 de julio de 201204:35:14 p.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaNominasIpemsa
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
ALTER TABLE dbo.Historico ADD
	CuentaSueldos nvarchar(50) NULL,
	ProvAguinaldo nvarchar(50) NULL,
	ProvVacaciones nvarchar(50) NULL,
	INSSPatronal nvarchar(50) NULL,
	INATEC nvarchar(50) NULL,
	AguinaldoxPagar nvarchar(50) NULL,
	VacacionesxPagar nvarchar(50) NULL,
	INSSxPagar nvarchar(50) NULL,
	INATECxPagar nvarchar(50) NULL,
	IRxPagar nvarchar(50) NULL,
	PrestamoxPagar nvarchar(50) NULL,
	NominaxPagar nvarchar(50) NULL,
    CuentaHorasExtra nvarchar(50) NULL
GO
COMMIT
