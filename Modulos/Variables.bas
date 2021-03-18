Attribute VB_Name = "Variables"
Public bExportar As Boolean, CodigoDepartamento As String, CuentaContable As String
Public pActualiza As Boolean
Public DisfrutadosConsulta As Double
Public NumNom13Mes As Integer, Conexion As String, ConexionContable As String
Public VerCia As Boolean, CopiaCia As Boolean, BorraCia As Boolean, SelecCia As Boolean, VerEditaNiveles As Boolean, VerEmpleado As Boolean, GEmpleado As Boolean, BEmpleado As Boolean
Public VerAnotaciones As Boolean, GAnotaciones As Boolean, BAnotaciones As Boolean, VerDepartamento As Boolean, GDepartamento As Boolean, BDepartamento As Boolean, VerCargo As Boolean, GCargo As Boolean, BCargo As Boolean, VerTipoIncapacidad As Boolean, GTipoIncapacidad As Boolean, BTipoIncapacidad As Boolean
Public VerIncapacidad As Boolean, GIncapacidad As Boolean, BIncapacidad As Boolean, VerPrestamos As Boolean, GPrestamos As Boolean, BPrestamos As Boolean, VerTipoNomina As Boolean, GTipoNomina As Boolean, BTipoNomina As Boolean, VerInss As Boolean, GInss As Boolean, BInss As Boolean, VerRespaldo As Boolean, VerUsuarios As Boolean, VerRegistroMoneda As Boolean, GRegistroMoneda As Boolean, BRegistroMoneda As Boolean
Public TieneDatos As Variant, BaseEntrada As Boolean, Cancela As Boolean
Public ruta As String, Fecha As Boolean, Server As String
Public SaldoPrestamo As String, Monto As String, MontoPagar As String
Public VarClave As Variant, Controlador As Variant, Guia As Variant, Origen As Variant, Destino As String, Guarda As Variant, Prueba As Variant, Valida As Integer, Evaluar As Boolean, Salida As Boolean, Active As Boolean, Lectura As Variant, Respuesta As Variant, Contesta As Boolean, error As Integer, Usuario As Integer, NivelAcceso As Variant, CodPasword As Variant, NombreUsuario As String
Public NumReport As Integer, MontoBruto As Double
Public NumNomina As Long, MontoBrutoMensual As Double
Public NumNominaSubsidio As Long, SalarioBasico As Double
Public NumNomVaca As Long, MesIni As Integer, NumNominaAnterior As Long
Public CodEmpleado As String, AnoIni As Integer
Public NumPrestamo As Long, CodTipoNomina As String
Public RutaFoto As String, TotalDevengadoAnterior As Double
Public MontoInssAnterior As Double, MontoIrAnterior As Double
Public FechaPde As Date, FechaPHasta As Date, FechaSde As Date, FechaShasta As Date
Public Quien As String, Meses, SalarioBasicoAnterior As Double, HorasExtraAnterior As Double
Public Inicio As Integer, Fin As Integer, IncentivoAnterior As Double, OtrosIngresosAnterior As Double
Public Tasa As Double, DiaFin As Integer, DestajoAnterior As Double, ComisionesAnterior As Double
Public Titulo As String, MontoInssMensual As Double, TasaCambioR As Double
Public SubTitulo As String, MontoIrMensual As Double
Public RutaLogo As String, MontoInssPatronalMensual As Double
Public QuienLlama As String, MontoIrPatronalAnterior As Double
Public ConexionReporte As String, MontoInssPatronalAnterior As Double
Public Total As Double, Contador As Integer
Public NumFecha1 As Long, NumFecha2 As Long
Public Directorio As String, RutaIconos As String, SQlReportes As String
Public SueldoFijo As Boolean, Factor As Double
Public CantMes As Double, DiasMes As Double, Criterio As String
Public MontoAdelanto13 As Double, MontoAdelantoVaca As Double
Public FechaSubsidio As Long, MontoVacaciones As Double, Convertir As String
Public FechaIniAgui As Date, FechaFinAgui As Date, DiasDescuento As Double
Public FechaIniVaca As Date, FechaFinVaca As Date
Public TotalIngreso As Double, TotalEgreso As Double, NetoPagar As Double
Public CodigoUsuario As Integer, Orden As Boolean, QueProducto As String
Public RutaServer As String, CodLinea As Double, TotalHoras As Double
Public TarifaHoraria As Double, Exportar As Boolean
Public Fecha1Reporte As String, Fecha2Reporte As String, FechaInicio As String, FechaFinal As String
Public saPeriodos(1 To 5) As Integer, ConteoEmpleados As Double
Public FechaIngreso As Date
Public pTotalVacacionesDisfrutadas, pTotalDiasVacaciones, pTotalDiasDisponibles As Double
'Public objExcel As Excel.Application
Public ConexionReloj As String, RutaServerReloj As String
Public ConexionEasy As String, RutaServerEasy As String
Public RutaArchivo As String
Public i As Double
Public CodigoH As String
Public MontoInssRegistros(3) As Double
Public PeriodoReporte As String, tempDias As Double, RegistrarBitacora As Boolean
