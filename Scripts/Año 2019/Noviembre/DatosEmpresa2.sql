/*
   Lunes, 16 de Noviembre de 2009 07:25:04 p.m.
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
CREATE TABLE dbo.Tmp_DatosEmpresa
	(
	Numero int NOT NULL,
	NombreEmpresa nvarchar(50) NULL,
	NumeroRUC nvarchar(50) NULL,
	Direccion nvarchar(50) NULL,
	Telefono nvarchar(50) NULL,
	Fax nvarchar(50) NULL,
	Email nvarchar(50) NULL,
	RutaLogo nvarchar(250) NULL,
	MetodoVacaciones nvarchar(50) NULL,
	FormatoColilla nvarchar(50) NULL,
	FormatoNomina nvarchar(50) NULL,
	MetodoCalculo nvarchar(50) NULL,
	RutaFoto nvarchar(250) NULL
	)  ON [PRIMARY]
GO
IF EXISTS(SELECT * FROM dbo.DatosEmpresa)
	 EXEC('INSERT INTO dbo.Tmp_DatosEmpresa (Numero, NombreEmpresa, NumeroRUC, Direccion, Telefono, Fax, Email, RutaLogo, FormatoColilla, FormatoNomina, MetodoCalculo, RutaFoto)
		SELECT Numero, NombreEmpresa, NumeroRUC, Direccion, Telefono, Fax, Email, RutaLogo, FormatoColilla, FormatoNomina, MetodoCalculo, RutaFoto FROM dbo.DatosEmpresa TABLOCKX')
GO
DROP TABLE dbo.DatosEmpresa
GO
EXECUTE sp_rename N'dbo.Tmp_DatosEmpresa', N'DatosEmpresa', 'OBJECT'
GO
COMMIT
