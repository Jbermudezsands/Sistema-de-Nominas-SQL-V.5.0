VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepNomVacaciones 
   Caption         =   "Reporte de Vacaciones"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepNomVacaciones.dsx":0000
End
Attribute VB_Name = "ArepNomVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Dim Titulo As String, SubTitulo As String

MDIPrimero.DtaEmpresa.Refresh
Titulo = MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa")
SubTitulo = MDIPrimero.DtaEmpresa.Recordset("Direccion") + " RUC: " + MDIPrimero.DtaEmpresa.Recordset("numeroruc")


      Me.lbltitulo.Caption = Titulo
      Me.LblSubtitulo.Caption = SubTitulo
      Me.ImgLogo.Picture = LoadPicture(RutaLogo)
      
      Me.LblFecha.Caption = "Desde " + Format(Frm13Vaca.TxtFINIVaca.Value, "dd/mm/yyyy") + " Hasta " + Format(Frm13Vaca.TxtFFinVaca.Value, "dd/mm/yyyy")
      Me.LblFechaHoy = Format(Now, "dddddd")
End Sub

Private Sub Detail_Format()
Dim CodEmpleado As Double, FechaIni As Date, FechaFin As Date, CodTipoNomina As String, SubTotal As Double
Dim DiasPagar As Double, DiasDescuento As Double, DiasNeto As Double

CodEmpleado = Me.FldCodEmpleado.Text
FechaIni = FechaIniVaca
FechaFin = FechaFinVaca
CodTipoNomina = Me.FldCodTipoNomina.Text

DiasPagar = Me.FldDiasPagar.Text
DiasDescuento = Me.FldDiasDescuento.Text
DiasNeto = Format(DiasPagar - DiasDescuento, "##,##0.00")

Me.LblDiasNeto.Caption = DiasNeto


MDIPrimero.DtaConsulta.RecordSource = "SELECT  * From Empleado Where (CodEmpleado = " & CodEmpleado & ")"
MDIPrimero.DtaConsulta.Refresh
If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
  Me.LblCedula.Caption = MDIPrimero.DtaConsulta.Recordset("NumCedula")
  Me.LblInss.Caption = MDIPrimero.DtaConsulta.Recordset("NumeroInss")
End If
 
MDIPrimero.DtaConsulta.RecordSource = "SELECT  Nomina.CodTipoNomina, DetalleNomina.CodEmpleado, SUM(DetalleNomina.BonoProduccion) AS BonoProduccion, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.HorasExtras) AS HorasExtras, SUM(DetalleNomina.Comisiones) AS Comisiones, SUM(DetalleNomina.DiasDescuento) AS DiasDescuento, SUM(DetalleNomina.Adelantos) AS Adelantos, SUM(DetalleNomina.Incentivos) AS Incentivos,SUM(DetalleNomina.Deducciones) AS Deducciones, SUM(DetalleNomina.DiasVacaciones) AS DiasVacaciones, SUM(DetalleNomina.VacacionesPagadas) AS VacacionesPagadas, SUM(DetalleNomina.Prestamo) AS Prestamo, SUM(DetalleNomina.MontoINSS) AS MontoINSS, SUM(DetalleNomina.MontoIR) AS MontoIR, SUM(DetalleNomina.Vacaciones) AS Vacaciones, SUM(DetalleNomina.OtrosIngresos) AS OtrosIngresos, SUM(DetalleNomina.INSSPatronal) AS INSSPatronal,  " & _
                                  "SUM(DetalleNomina.IRPatronal) AS IRPatronal, SUM(DetalleNomina.Mes13) AS Mes13, SUM(DetalleNomina.TotalSubsidio) AS TotalSubsidio, SUM(DetalleNomina.HE) AS HE, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) AS TotalDevengado, SUM(DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS TotalDeducir,SUM((DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) " & _
                                  ") AS NetoPagar, SUM(DetalleNomina.TarifaHoraria) AS TarifaHoraria, SUM(DetalleNomina.HTrabajada) AS HTrabajada, SUM(DetalleNomina.SeptimoDia) AS SeptimoDia, SUM(DetalleNomina.IncetivoProduccion) AS IncetivoProduccion, SUM(DetalleNomina.AjusteINSS) AS AjusteINSS FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina  " & _
                                  "WHERE (Nomina.FechaNominaINI >= CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102)) AND (Nomina.FechaNomina <= CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) GROUP BY Nomina.CodTipoNomina, DetalleNomina.CodEmpleado HAVING (Nomina.CodTipoNomina = '" & CodTipoNomina & "') AND (DetalleNomina.CodEmpleado = " & CodEmpleado & ") AND (SUM((DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones)) <> 0) ORDER BY DetalleNomina.CodEmpleado"
MDIPrimero.DtaConsulta.Refresh
If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
  Me.LblSalarioMensual.Caption = Format(MDIPrimero.DtaConsulta.Recordset("NetoPagar"), "##,##0.00")
  Me.LblIRMensual.Caption = Format(MDIPrimero.DtaConsulta.Recordset("MontoIR"), "##,##0.00")
  SubTotal = CDbl(MDIPrimero.DtaConsulta.Recordset("NetoPagar")) + CDbl(Me.Field14.Text)
  Me.LblSubTotal.Caption = Format(SubTotal, "##,##0.00")
End If


End Sub

Private Sub GroupFooter2_Format()
 Dim TotalDias As Double, TotalDiasDescuento As Double, TotalNeto As Double
 
 TotalDias = Me.FldTotalDiasPagar.Text
 TotalDiasDescuento = Me.FldTotalDiasDescuento.Text
 
 TotalNeto = Format(TotalDias - TotalDiasDescuento, "##,##0.00")
 
 Me.LblTotalNeto.Caption = Format(TotalNeto, "##,##0.00")
 


End Sub

