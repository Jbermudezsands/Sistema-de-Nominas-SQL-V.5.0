/*
   Jueves, 07 de Enero de 2010 09:12:29 a.m.
   Usuario: 
   Servidor: JUAN\SQL2000
   Base de datos: SistemaNominasMetro
   Aplicación: MS SQLEM - Data Tools
*/

BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET TRANSACTION ISOLATION LEVEL SERIALIZABLE
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
ALTER TABLE dbo.TipoSubsidio ADD
	CuentaContable nvarchar(200) NULL
GO
COMMIT
