/*
   lunes, 29 de agosto de 201110:21:34 a.m.
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
CREATE TABLE dbo.ActividadesProduccion
	(
	Raiz int NOT NULL,
	Codigo int NOT NULL,
	Empleado numeric(18, 0) NOT NULL,
	Fecha datetime NOT NULL,
	HEntrada datetime NOT NULL,
	HSalida datetime NOT NULL,
	HExtras bit NOT NULL,
	FechaAprobado datetime NULL,
	Eliminar bit NOT NULL
	)  ON [PRIMARY]
GO
ALTER TABLE dbo.ActividadesProduccion ADD CONSTRAINT
	PK_ActividadesProduccion PRIMARY KEY CLUSTERED 
	(
	Raiz,
	Codigo,
	Empleado,
	Fecha
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
COMMIT
