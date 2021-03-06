/*
   jueves, 10 de mayo de 201206:25:34 a.m.
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
CREATE TABLE dbo.Reportes
	(
	Campo1 nvarchar(50) NULL,
	Campo2 nvarchar(50) NULL,
	Campo3 nvarchar(50) NULL,
	Campo4 nvarchar(50) NULL,
	Campo5 nvarchar(50) NULL,
	Campo6 nvarchar(50) NULL,
	Campo7 nvarchar(50) NULL,
	Campo8 nvarchar(50) NULL,
	Campo9 nvarchar(50) NULL,
	Campo10 nvarchar(50) NULL,
	Num1 float(53) NULL,
	Num2 float(53) NULL,
	Num3 float(53) NULL,
	Num4 float(53) NULL,
	Num5 float(53) NULL,
	Num6 float(53) NULL,
	Num7 float(53) NULL,
	Num8 float(53) NULL,
	Num9 float(53) NULL,
	Num10 float(53) NULL,
	Fecha1 datetime NULL,
	Fecha2 datetime NULL,
	Fecha3 datetime NULL,
	Fecha4 datetime NULL,
	Fecha5 datetime NULL,
	Fecha6 datetime NULL,
	Fecha7 datetime NULL,
	Fecha8 datetime NULL,
	Fecha9 datetime NULL,
	Fecha10 datetime NULL
	)  ON [PRIMARY]
GO
COMMIT
