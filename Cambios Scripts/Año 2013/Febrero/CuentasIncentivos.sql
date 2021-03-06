/*
   viernes, 22 de febrero de 201307:48:31 p.m.
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
CREATE TABLE dbo.CuentasIncentivos
	(
	CodEmpleado nvarchar(50) NOT NULL,
	CodIncentivo nvarchar(50) NOT NULL,
	CodCuentas nvarchar(50) NOT NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.CuentasIncentivos ADD CONSTRAINT
	PK_CuentasIncentivos PRIMARY KEY CLUSTERED 
	(
	CodEmpleado,
	CodIncentivo,
	CodCuentas
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
