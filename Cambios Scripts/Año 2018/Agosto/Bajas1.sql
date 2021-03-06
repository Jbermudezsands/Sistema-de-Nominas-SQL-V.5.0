/*
   sábado, 25 de agosto de 201811:15:44 a.m.
   Usuario: sa
   Servidor: JUANBERMUDEZ\SQL2014
   Base de datos: SNTextilesValidos
   Aplicación: 
*/

/* Para evitar posibles problemas de pérdida de datos, debe revisar este script detalladamente antes de ejecutarlo fuera del contexto del diseñador de base de datos.*/
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
ALTER TABLE dbo.Bajas ADD
	FechaEgreso smalldatetime NULL,
	FechaHistorial smalldatetime NULL,
	SueldoActualBasicoLiquida bit NULL,
	PrestamoOpt bit NULL,
	DeduccionesOpt bit NULL,
	AguinaldoOpt bit NULL,
	AntiguedadOpt bit NULL,
	ViaticosOpt bit NULL,
	VacacionesOpt bit NULL,
	HorasExtraOpt bit NULL,
	OtrosIngresosOpt bit NULL,
	OtrosIngresosPlanillaOpt bit NULL,
	Prestacion float(53) NULL,
	MontoPagarPrestacion bit NULL
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	DF_Bajas_SueldoActualBasico DEFAULT 0 FOR SueldoActualBasicoLiquida
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	DF_Bajas_PrestamoOpt DEFAULT 0 FOR PrestamoOpt
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	DF_Bajas_DeduccionesOpt DEFAULT 0 FOR DeduccionesOpt
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	DF_Bajas_13VoOpt DEFAULT 0 FOR AguinaldoOpt
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	DF_Bajas_AntiguedadOpt DEFAULT 0 FOR AntiguedadOpt
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	DF_Bajas_ViaticosOpt DEFAULT 0 FOR ViaticosOpt
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	DF_Bajas_VacacionesOpt DEFAULT 0 FOR VacacionesOpt
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	DF_Bajas_HorasExtraOpt DEFAULT 0 FOR HorasExtraOpt
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	DF_Bajas_OtrosIngresosOpt DEFAULT 0 FOR OtrosIngresosOpt
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	DF_Bajas_OtrosIngresosPlanillaOpt DEFAULT 0 FOR OtrosIngresosPlanillaOpt
GO
ALTER TABLE dbo.Bajas ADD CONSTRAINT
	DF_Bajas_MontaPagarPrestacion DEFAULT 0 FOR MontoPagarPrestacion
GO
ALTER TABLE dbo.Bajas SET (LOCK_ESCALATION = TABLE)
GO
COMMIT
