VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepProvicion 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepProvicion.dsx":0000
End
Attribute VB_Name = "ArepProvicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalDiasVacaAcumulado As Double, TotalDiasVacaNomina As Double, TotalProvicionVacaciones As Double, TotalDiasAguinaldo As Double, TotalProvicionAguinaldo As Double, TotalProvicionNomina As Double


Private Sub ActiveReport_ReportStart()
 
         TotalDiasVacaAcumulado = 0
         TotalDiasVacaNomina = 0
         TotalProvicionVacaciones = 0
         TotalDiasAguinaldo = 0
         TotalProvicionAguinaldo = 0
         TotalProvicionNomina = 0
         
         'Me.LblTitulo3.Caption = "Nomina # " & FrmReportes.TxtNNomina.Text & " Periodo Desde " & FrmReportes.MtxtFechaini.Value & " Hasta " & FrmReportes.MtxtFecha.Value
         
End Sub

Private Sub Detail_Format()
 Dim sql As String, CodigoEmpleado As Double, FechaVaca As Date
 Dim FechaIncioVaca As Date, FechaContrato As Date
 Dim FechaNomina As Date, FechaFin As Date
 Dim FechaFinNomina As Date, ValorDia As Double
 Dim DiasAcumuladoVaca As Double, DiasVacaNomina As Double
 Dim DiasMes As Double, DiasAguinaldo As Double
 Dim ProvicionNomina As Double

   MDIPrimero.DtaControles.Refresh
   If Not MDIPrimero.DtaControles.Recordset.EOF Then
      DiasMes = MDIPrimero.DtaControles.Recordset("DiasMes")
   End If
   
   


FechaContrato = Me.FldFechaIngreso.Text
FechaNomina = Me.FldFechaFinNomina.Text    'Me.FldFechaIniNomina.Text
FechaFinNomina = Me.FldFechaFinNomina.Text

CodigoEmpleado = Me.FldCodigoEmpleado.Text

If CodigoEmpleado = "13200" Then
 CodigoEmpleado = Me.FldCodigoEmpleado.Text
End If
'------------------------------BUSCO LA ULTIMA NOMINA DE VACACIONES-------------------
sql = "SELECT DetalleNomVaca.Id, DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, DetalleNomVaca.Inss, DetalleNomVaca.TarifaHoraria, DetalleNomVaca.TotalDevengado, DetalleNomVaca.IR , NomVaca.FechaFin, NomVaca.FechaIni, NomVaca.FechaAplica FROM DetalleNomVaca INNER JOIN  NomVaca ON DetalleNomVaca.NumNomVaca = NomVaca.NumNomVaca " & _
      "WHERE (DetalleNomVaca.CodEmpleado = " & CodigoEmpleado & ") AND (NomVaca.FechaFin <= CONVERT(DATETIME, '" & Format(FechaNomina, "yyyy-mm-dd") & "', 102))"

FrmReportes.AdoConsulta.RecordSource = sql
FrmReportes.AdoConsulta.Refresh
If Not FrmReportes.AdoConsulta.Recordset.EOF Then
  FechaVaca = FrmReportes.AdoConsulta.Recordset("FechaFin")
Else
  FechaVaca = Me.FldFechaIngreso.Text
End If

Me.LblFechaVacaciones.Caption = FechaVaca


FechaFin = FechaNomina
FechaInicioVaca = FechaVaca

ValorDia = Me.FldValorDia.Text



'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////BUSCO LOS DIAS DE VACACIONES ///////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
         MesesAguinaldo = 0
         FechaInicioVaca = FechaVaca
         If CDate(FechaContrato) > CDate(FechaInicioVaca) Then FechaInicioVaca = FechaContrato
'         DiasAguinaldo = (CDbl(FechaFin) - CDbl(FechaInicioVaca) + 1) / DiasMes '
         DiasAguinaldo = (DateDiff("d", FechaInicioVaca, FechaFin) + 1) * 0.083333

'         DiasVacaciones = DateDiff("d", CDbl(FechaInicioVaca), CDbl(FechaFin))
          DiasVacaciones = CalcularDiasVaca(CDate(FechaInicioVaca), FechaFin) + 1
         If MesesAguinaldo < 12 Then
          DiasProporcional = MesesAguinaldo - Int(MesesAguinaldo)
          DiasVacaciones = DiasVacaciones * 0.08333333
         Else
          DiasVacaciones = 12 * 2.5

         End If
         
         
         DiasAcumuladoVaca = Format(DiasVacaciones, "##,##0.00")
         Me.LblDiasAcumuladoVacaciones.Caption = DiasAcumuladoVaca
         DiasVacaNomina = CalcularDiasVaca(CDate(Me.FldFechaIniNomina.Text), FechaFinNomina) + 1
         Me.LblDiasVacaNomina.Caption = Format(DiasVacaNomina * 0.0833333, "##,##0.00")
         

         ProvicionNomina = Format(ValorDia * CDbl(Format(DiasVacaNomina * 0.083333, "##,##0.00")), "##,##0.00")
         Me.LblProvicionNomina.Caption = Format(ProvicionNomina, "##,##0.00")
         
         Me.LblProvicionVacaciones.Caption = Format(DiasAcumuladoVaca * ValorDia, "##,##0.00")
         
         Me.LblDiasAguinaldo.Caption = Format(DiasAguinaldo, "##,##0.00")
         DiasAguinaldo = Format(DiasAguinaldo, "##,##0.00")
         Me.LblProvicionAguinaldo.Caption = Format(DiasAguinaldo * ValorDia, "##,##0.00")
         
         Me.LblDiasLiquida.Caption = Format(DiasAguinaldo, "##,##0.00")
         Me.LblProvicionLiquida.Caption = Format(DiasAguinaldo * ValorDia, "##,##0.00")
         
         TotalDiasVacaAcumulado = TotalDiasVacaAcumulado + DiasAcumuladoVaca
         TotalDiasVacaNomina = TotalDiasVacaNomina + DiasVacaNomina
         TotalProvicionVacaciones = TotalProvicionVacaciones + DiasVacaNomina + (DiasAguinaldo * ValorDia)
         TotalDiasAguinaldo = TotalDiasAguinaldo + DiasAguinaldo
         TotalProvicionAguinaldo = TotalProvicionAguinaldo + (DiasAguinaldo * ValorDia)
         TotalProvicionNomina = TotalProvicionNomina + ProvicionNomina
         
         


End Sub

Private Sub GroupFooter1_Format()
  Me.LblTotalDiasVacaAcumulado.Caption = Format(TotalDiasVacaAcumulado, "##,##0.00")
  Me.LblTotalDiasNomina.Caption = Format(TotalDiasVacaNomina, "##,##0.00")
  Me.LblTotalProvicionVacaciones.Caption = Format(TotalProvicionVacaciones, "##,##0.00")
  Me.LblTotalDiasAcumuladoAguinaldo.Caption = Format(TotalDiasAguinaldo, "##,##0.00")
  Me.LblTotalProvicionAguinaldo.Caption = Format(TotalProvicionAguinaldo, "##,##0.00")
  Me.LblTotalDiasAcumuladoLiquida.Caption = Format(TotalDiasAguinaldo, "##,##0.00")
  Me.LblTotalProvicionLiquida.Caption = Format(TotalProvicionAguinaldo, "##,##0.00")
    Me.LblTotalProvicionNomina.Caption = Format(TotalProvicionNomina, "##,##0.00")
End Sub


