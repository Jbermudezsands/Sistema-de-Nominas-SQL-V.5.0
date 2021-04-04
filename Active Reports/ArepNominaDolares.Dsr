VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepNominaDolares 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ArepNominaDolares.dsx":0000
End
Attribute VB_Name = "ArepNominaDolares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
Dim FechaNomina As Date, TasaCambio As Double, Salario As Double
Dim MontoSubsidio As Double, CodigoEmpleado As Double, Septimo As Double, TotalOtrosIngresos As Double
Dim Comisiones As Double, Incentivos As Double, OtrosIngresos As Double, HorasExtra As Double
Dim TotalDevengado As Double, Inss As Double, TotalDeducciones As Double, Ir As Double

Ir = 0
Inss = 0
'//////////////////////////////////////////////////////////////////////////////////////
'/////////////BUSCO LOS OTROS INGRESOS//////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////
If Val(Me.FldSeptimo.Text) <> 0 Then
 Septimo = Me.FldSeptimo.Text
End If
If Val(Me.FldComisiones.Text) <> 0 Then
Comisiones = Me.FldComisiones.Text
End If
If Val(Me.FldIncentivoProduccion.Text) <> 0 Then
 Incentivos = Me.FldIncentivoProduccion.Text
End If
If Val(Me.FldIncentivos.Text) <> 0 Then
 Incentivos = Incentivos + Me.FldIncentivos.Text
End If
If Val(Me.FldOtrosIngresos.Text) <> 0 Then
 OtrosIngresos = Me.FldOtrosIngresos.Text
End If
If Val(Me.FldHorasExtra.Text) <> 0 Then
 HorasExtra = Me.FldHorasExtra.Text
End If

TotalOtrosIngresos = Septimo + Comisiones + Incentivos + OtrosIngresos + HorasExtra
Me.LblOtrosIngresos.Caption = Format(TotalOtrosIngresos, "##,##0.00")
If Me.FldCodEmpleado.Text <> "" Then
 CodigoEmpleado = Me.FldCodEmpleado.Text
End If
'//////////////////////////////////////////////////////////////////////////////////////
'/////////////BUSCO LA TASA DE CAMBIO///////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////

If Not Me.FldFecha.Text = "" Then
FechaNomina = Me.FldFecha.Text
Else
 FechaNomina = Format(Now, "dd/mm/yyyy")
End If
If Not Me.FldSalario.Text = "" Then
Salario = Me.FldSalario.Text
End If
Me.AdoTasas.RecordSource = "SELECT FechaDia, MontoDia From CambioMoneda WHERE (FechaDia = '" & Format(FechaNomina, "yyyymmdd") & "')"
Me.AdoTasas.Refresh
If Not Me.AdoTasas.Recordset.EOF Then
   TasaCambio = Me.AdoTasas.Recordset("MontoDia")
Else
   TasaCambio = 0
End If

 Me.LblDevengado.Caption = Format(TasaCambio * Salario, "##,##0.00")
 Me.LblMensual1.Caption = Format(Salario * 2, "##,##0.00")
 Me.LblMensual2.Caption = Format((TasaCambio * Salario) * 2, "##,##0.00")
 
'//////////////////////////////////////////////////////////////////////////////////////
'/////////////BUSCO EL MONTO DEL SUBSIDIO//////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////
MontoSubsidio = 0
Me.AdoConsulta.RecordSource = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Subsidio.NumSubsidio, TipoSubsidio.Subsidio, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.Descripcion, DetalleSubsidio.NumNominaSubsidio,DetalleSubsidio.Pagado FROM TipoSubsidio INNER JOIN Empleado INNER JOIN Subsidio ON Empleado.CodEmpleado = Subsidio.CodEmpleado INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio ON TipoSubsidio.CodTipoSubsidio = Subsidio.CodTipoSubsidio Where (DetalleSubsidio.Pagado = 0) And (Empleado.CodEmpleado = " & CodigoEmpleado & ") ORDER BY Empleado.CodEmpleado"
Me.AdoConsulta.Refresh
If Not Me.AdoConsulta.Recordset.EOF Then
 If Not IsNull(Me.AdoConsulta.Recordset("Valor")) Then
   MontoSubsidio = Me.AdoConsulta.Recordset("Valor")
 End If
End If
Me.LblSubsidio.Caption = Format(MontoSubsidio, "##,##0.00")

'//////////////////////////////////////////////////////////////////////////////////////
'/////////////SUMO EL TOTAL DEVENGADO//////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////

 TotalDevengado = MontoSubsidio + Salario + TotalOtrosIngresos
 Me.LblDevengado1.Caption = Format(TotalDevengado, "##,##0.00")
 Me.LblDevengado2.Caption = Format(TotalDevengado * TasaCambio, "##,##0.00")
 
 If Val(Me.FldInss.Text) <> 0 Then
 Inss = Me.FldInss.Text
 Me.LblInss.Caption = Format(Inss * TasaCambio, "##,##0.00")
 End If

 
 If Val(Me.FldIr.Text) <> 0 Then
  Ir = Me.FldIr.Text
  Me.LblIr.Caption = Format(Ir * TasaCambio, "##,##0.00")
 Else
  Me.LblIr.Caption = "0.00"
 End If

 If Val(Me.FldTotalDeducir.Text) <> 0 Then
 TotalDeducciones = Me.FldTotalDeducir.Text
 End If
 Me.LblNeto1.Caption = Format(TotalDevengado - TotalDeducciones, "##,##0.00")
 Me.LblNeto2.Caption = Format((TotalDevengado - TotalDeducciones) * TasaCambio, "##,##0.00")
 
 
End Sub

Private Sub PageHeader_Format()


With Me.AdoTasas
   .ConnectionString = Conexion
End With

With Me.AdoConsulta
   .ConnectionString = Conexion
End With
 
 
 
End Sub


Private Sub ActiveReport_ReportEnd()
If Exportar = True Then
    Dim myExportObject As ActiveReportsExcelExport.ARExportExcel
    Dim Nombre As String
    
'    Nombre = InputBox("Digite el Nombre del Archivo", "Sistema de Nominas")
    Set myExportObject = CreateObject("ActiveReportsExcelExport.ARExportExcel")
    myExportObject.FileName = Directorio
    myExportObject.Export Me.Pages
    Set myExportObject = Nothing
 End If
End Sub
