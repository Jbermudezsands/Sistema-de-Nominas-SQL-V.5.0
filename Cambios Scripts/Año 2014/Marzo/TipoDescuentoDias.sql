/*
   miércoles, 19 de marzo de 201406:32:28 p.m.
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
CREATE TABLE dbo.TipoDescuentoDias
	(
	TipoAusencia nvarchar(50) NOT NULL,
	Color nvarchar(50) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.TipoDescuentoDias ADD CONSTRAINT
	PK_TipoDescuentoDias PRIMARY KEY CLUSTERED 
	(
	TipoAusencia
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
