/*
   jueves, 13 de octubre de 201110:12:34 p.m.
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
CREATE TABLE dbo.Reembolso
	(
	NumNomina float(53) NOT NULL,
	CodEmpleado nvarchar(50) NOT NULL,
	Monto float(53) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Reembolso ADD CONSTRAINT
	DF_Reembolso_Monto DEFAULT 0 FOR Monto
GO
ALTER TABLE dbo.Reembolso ADD CONSTRAINT
	PK_Reembolos PRIMARY KEY CLUSTERED 
	(
	NumNomina,
	CodEmpleado
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
