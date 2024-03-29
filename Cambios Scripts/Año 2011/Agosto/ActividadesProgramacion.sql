/*
   lunes, 29 de agosto de 201111:08:34 a.m.
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
CREATE TABLE dbo.ActividadesProgramacion
	(
	Raiz int NOT NULL,
	Codigo int NOT NULL,
	TipoNomina nvarchar(50) NOT NULL,
	Fecha datetime NOT NULL,
	HPOrdinaria int NOT NULL,
	HPExtra int NOT NULL,
	CerrarPlaneacion bit NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.ActividadesProgramacion ADD CONSTRAINT
	PK_ActividadesProgramacion PRIMARY KEY CLUSTERED 
	(
	Raiz,
	Codigo,
	TipoNomina,
	Fecha
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
