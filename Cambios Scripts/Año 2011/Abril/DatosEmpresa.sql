/*
   Jueves, 28 de Abril de 201110:31:57 a.m.
   Usuario: 
   Servidor: CONSULTOR\SQL2005
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
ALTER TABLE dbo.DatosEmpresa
	DROP CONSTRAINT DF_DatosEmpresa_MetodoVacaciones
GO
ALTER TABLE dbo.DatosEmpresa
	DROP CONSTRAINT DF_DatosEmpresa_TipoCalculoIR
GO
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
	RutaFoto nvarchar(250) NULL,
	ConexionSistemaContable nvarchar(250) NULL,
	TipoCalculoIR nvarchar(50) NULL,
	Calcular7mo bit NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Tmp_DatosEmpresa ADD CONSTRAINT
	DF_DatosEmpresa_MetodoVacaciones DEFAULT (N'Vacaciones Semestrales') FOR MetodoVacaciones
GO
ALTER TABLE dbo.Tmp_DatosEmpresa ADD CONSTRAINT
	DF_DatosEmpresa_TipoCalculoIR DEFAULT (N'Calcular Ajustando IR') FOR TipoCalculoIR
GO
ALTER TABLE dbo.Tmp_DatosEmpresa ADD CONSTRAINT
	DF_DatosEmpresa_Calcular7mo DEFAULT 0 FOR Calcular7mo
GO
IF EXISTS(SELECT * FROM dbo.DatosEmpresa)
	 EXEC('INSERT INTO dbo.Tmp_DatosEmpresa (Numero, NombreEmpresa, NumeroRUC, Direccion, Telefono, Fax, Email, RutaLogo, MetodoVacaciones, FormatoColilla, FormatoNomina, MetodoCalculo, RutaFoto, ConexionSistemaContable, TipoCalculoIR, Calcular7mo)
		SELECT Numero, NombreEmpresa, NumeroRUC, Direccion, Telefono, Fax, Email, RutaLogo, MetodoVacaciones, FormatoColilla, FormatoNomina, MetodoCalculo, RutaFoto, ConexionSistemaContable, TipoCalculoIR, CONVERT(bit, Calcular7mo) FROM dbo.DatosEmpresa WITH (HOLDLOCK TABLOCKX)')
GO
DROP TABLE dbo.DatosEmpresa
GO
EXECUTE sp_rename N'dbo.Tmp_DatosEmpresa', N'DatosEmpresa', 'OBJECT' 
GO
COMMIT
