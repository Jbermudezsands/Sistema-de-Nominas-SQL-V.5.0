/*
   lunes, 29 de agosto de 201110:46:10 a.m.
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
CREATE TABLE dbo.Tmp_PuntosGrupo
	(
	Id int NOT NULL IDENTITY (1, 1),
	Grupo nvarchar(50) NULL
	)  ON [PRIMARY]
GO
SET IDENTITY_INSERT dbo.Tmp_PuntosGrupo ON
GO
IF EXISTS(SELECT * FROM dbo.PuntosGrupo)
	 EXEC('INSERT INTO dbo.Tmp_PuntosGrupo (Id, Grupo)
		SELECT Id, Grupo FROM dbo.PuntosGrupo WITH (HOLDLOCK TABLOCKX)')
GO
SET IDENTITY_INSERT dbo.Tmp_PuntosGrupo OFF
GO
DROP TABLE dbo.PuntosGrupo
GO
EXECUTE sp_rename N'dbo.Tmp_PuntosGrupo', N'PuntosGrupo', 'OBJECT' 
GO
ALTER TABLE dbo.PuntosGrupo ADD CONSTRAINT
	PK_PuntosGrupo PRIMARY KEY CLUSTERED 
	(
	Id
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
