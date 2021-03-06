/*
   lunes, 29 de agosto de 201110:47:42 a.m.
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
CREATE TABLE dbo.Tmp_Actividades
	(
	Llave int NOT NULL IDENTITY (1, 1),
	Raiz int NOT NULL,
	Codigo int NOT NULL,
	Actividad nvarchar(50) NOT NULL,
	PagaCliente bit NOT NULL
	)  ON [PRIMARY]
GO
SET IDENTITY_INSERT dbo.Tmp_Actividades ON
GO
IF EXISTS(SELECT * FROM dbo.Actividades)
	 EXEC('INSERT INTO dbo.Tmp_Actividades (Llave, Raiz, Codigo, Actividad, PagaCliente)
		SELECT Llave, Raiz, Codigo, Actividad, PagaCliente FROM dbo.Actividades WITH (HOLDLOCK TABLOCKX)')
GO
SET IDENTITY_INSERT dbo.Tmp_Actividades OFF
GO
DROP TABLE dbo.Actividades
GO
EXECUTE sp_rename N'dbo.Tmp_Actividades', N'Actividades', 'OBJECT' 
GO
ALTER TABLE dbo.Actividades ADD CONSTRAINT
	PK_Actividades PRIMARY KEY CLUSTERED 
	(
	Raiz,
	Codigo
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
