/*
   lunes, 29 de agosto de 201110:17:09 a.m.
   Usuario: 
   Servidor: JUAN\SQL2005
   Base de datos: SistemaNominasNorteak
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
CREATE TABLE dbo.ActividadesCuenta
	(
	Raiz int NOT NULL,
	Codigo int NOT NULL,
	TipoNomina nvarchar(10) NOT NULL,
	Cuenta nvarchar(50) NOT NULL
	)  ON [PRIMARY]
GO
COMMIT
