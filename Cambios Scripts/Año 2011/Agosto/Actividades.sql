/*
   lunes, 29 de agosto de 201110:14:19 a.m.
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
CREATE TABLE dbo.Actividades
	(
	Llave int NOT NULL,
	Raiz int NOT NULL,
	Codigo int NOT NULL,
	Actividad nvarchar(50) NOT NULL,
	PagaCliente bit NOT NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.Actividades ADD CONSTRAINT
	PK_Actividades PRIMARY KEY CLUSTERED 
	(
	Raiz,
	Codigo
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
