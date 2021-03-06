/*
   lunes, 29 de julio de 201303:41:34 p.m.
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
CREATE TABLE dbo.Table_1
	(
	Campo1 nvarchar(MAX) NULL,
	Campo2 nvarchar(MAX) NULL,
	Campo3 nvarchar(MAX) NULL,
	Campo4 nvarchar(MAX) NULL,
	Campo5 nvarchar(MAX) NULL,
	Campo6 nvarchar(MAX) NULL,
	Campo7 nvarchar(MAX) NULL,
	Campo8 nvarchar(MAX) NULL,
	Campo9 nvarchar(MAX) NULL,
	Campo10 nvarchar(MAX) NULL,
	Campo11 nvarchar(MAX) NULL,
	Campo12 nvarchar(MAX) NULL,
	Campo13 nvarchar(MAX) NULL,
	CampoNum1 numeric(18, 2) NULL,
	CampoNum2 numeric(18, 2) NULL,
	CampoNum3 numeric(18, 2) NULL,
	CampoNum4 numeric(18, 2) NULL,
	CampoNum5 numeric(18, 2) NULL,
	CampoNum6 numeric(18, 2) NULL,
	CampoNum7 numeric(18, 2) NULL,
	CampoNum8 numeric(18, 2) NULL,
	CampoNum9 numeric(18, 2) NULL,
	CampoNum10 numeric(18, 2) NULL
	)  ON [PRIMARY]
	 TEXTIMAGE_ON [PRIMARY]
GO
COMMIT
