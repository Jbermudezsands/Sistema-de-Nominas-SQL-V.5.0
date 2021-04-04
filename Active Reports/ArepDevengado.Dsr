VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepDevengado 
   Caption         =   "Reporte de Devengados"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   19420
   SectionData     =   "ArepDevengado.dsx":0000
End
Attribute VB_Name = "ArepDevengado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
    Me.LblTitulo.Caption = Titulo
    Me.LblSubtitulo.Caption = SubTitulo
    Me.Label5.Caption = "REPORTE ACUMULADO VACACIONES,AGUINALDO,INDEMNIZACION"
    Me.LblFecha.Caption = "Impreso desde: " & FrmReportes.TxtFecha1.Value & " Hasta: " & FrmReportes.TxtFecha2.Value
    Me.LblFechaHoy.Caption = Format(Now, "Long Date")
End Sub

Private Sub Detail_Format()
 Dim SqlString As String, CodEmpleado As Double
 Dim FechaNominaVaca As Date, FechaFin As Date
 Dim DiasVacaciones As Double, FechaNomina13vo As Date, DiasAguinaldo As Double
 Dim DiasAntiguedad As Double, FechaContrato As Date, DiasMes As Double, DiasLiquidacion As Double
 Dim DiasAntiguedadPro As Double, MesesAguinaldo As Double, DiasProporcional As Double, FechaInicioAgui As Date
 Dim FechaInicioVaca As Date, DiasAcumulados As Double
 
   MDIPrimero.DtaControles.Refresh
   If Not MDIPrimero.DtaControles.Recordset.EOF Then
      DiasMes = MDIPrimero.DtaControles.Recordset("DiasMes")
   End If
 
 If Me.FldCodEmpleados.Text <> "" Then
  CodEmpleado = Me.FldCodEmpleados.Text
 End If
 FechaFin = FrmReportes.TxtFecha2.Value
 If Me.FldFechaContrato.Text <> "" Then
  FechaContrato = Me.FldFechaContrato.Text
 End If
 DiasVacaciones = 0
 
 '/////////////////////////////////////////////////////////////////////////////////////
 '////////////////////////BUSCO LA ULTIMA NOMINA DE VACACIONES /////////////////////////////////////////
 '//////////////////////////////////////////////////////////////////////////////////////////
' SqlString = "SELECT  * FROM DetalleNomVaca INNER JOIN NomVaca ON DetalleNomVaca.NumNomVaca = NomVaca.NumNomVaca WHERE (DetalleNomVaca.CodEmpleado = " & CodEmpleado & ") AND (NomVaca.FechaFin <= CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-MM-dd") & "', 102))"
' MDIPrimero.DtaConsulta.RecordSource = SqlString
' MDIPrimero.DtaConsulta.Refresh
' If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
'  MDIPrimero.DtaConsulta.Recordset.MoveLast
'  FechaNominaVaca = MDIPrimero.DtaConsulta.Recordset("FechaFin")
'  DiasVacaciones = CDbl(FechaFin) - CDbl(FechaNominaVaca) + 1
'  DiasVacaciones = DiasVacaciones * (1 / 12)
' End If
 
 
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////BUSCO LA ULTIMA NOMINA DE 13VO /////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////
'SqlString = "SELECT  * FROM  DetalleNom13Mes INNER JOIN Nom13Mes ON DetalleNom13Mes.NumNom13Mes = Nom13Mes.NumNom13Mes  Where (DetalleNom13Mes.CodEmpleado = " & CodEmpleado & ")"
'MDIPrimero.DtaConsulta.RecordSource = SqlString
'MDIPrimero.DtaConsulta.Refresh
'If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
' MDIPrimero.DtaConsulta.Recordset.MoveLast
' FechaNomina13vo = MDIPrimero.DtaConsulta.Recordset("FechaFin")
' DiasAguinaldo = CDbl(FechaFin) - CDbl(FechaNomina13vo) + 1
' DiasAguinaldo = DiasAguinaldo * (1 / 12)
'End If

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////BUSCO LOS DIAS DE VACACIONES ///////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
         MesesAguinaldo = 0
         FechaInicioVaca = DateSerial(Year(FechaFin), 1, 1)
         If CDate(FechaContrato) > CDate(FechaInicioVaca) Then FechaInicioVaca = FechaContrato
         MesesAguinaldo = (CDbl(FechaFin) - CDbl(FechaInicioVaca) + 1) / DiasMes '
'         DiasVacaciones = DateDiff("d", CDbl(FechaInicioVaca), CDbl(FechaFin))
          DiasVacaciones = CalcularDiasVaca(CDate(FechaInicioVaca), FechaFin)
         If MesesAguinaldo < 12 Then
          DiasProporcional = MesesAguinaldo - Int(MesesAguinaldo)
          DiasVacaciones = DiasVacaciones * 0.08333333
         Else
          DiasVacaciones = 12 * 2.5

         End If
         
     '////////////////////////////////////////////////////////////////////////////////////////////////////////
     '/////////////////////////////CALCULO DE LOS DIAS ACUMULADO EN OTRAS VACACIONES /////////////////////////
     '////////////////////////////////////////////////////////////////////////////////////////////////////////
     MDIPrimero.DtaConsulta.RecordSource = "SELECT  DetalleNomVaca.CodEmpleado, SUM(DetalleNomVaca.DiasAPagar) AS DiasAPagar, SUM(DetalleNomVaca.DiasDescuento) AS DiasDescuento, SUM(DetalleNomVaca.AdelantoVacaciones) As AdelantoVacaciones FROM  NomVaca INNER JOIN  DetalleNomVaca ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca Where (NomVaca.Activa = 0) GROUP BY DetalleNomVaca.CodEmpleado Having (DetalleNomVaca.CodEmpleado = " & CodEmpleado & ")"
     MDIPrimero.DtaConsulta.Refresh
     If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
       DiasAcumulados = MDIPrimero.DtaConsulta.Recordset("DiasAPagar")
     Else
       DiasAcumulados = 0
     End If
     
      '///////////////////SI LOS DIAS A PAGAR ES MAYOR QUE LOS QUE YA SE HAN PAGADO LOS HAGO CERO //////////////////////////
      If DiasAcumulados > DiasVacaciones Then
        DiasVacaciones = 0
      Else
        DiasVacaciones = DiasVacaciones - DiasAcumulados
      End If

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////BUSCO LOS DIAS DE AGUINALDO ///////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////
         MesesAguinaldo = 0
         FechaInicioAgui = DateSerial(Year(FechaFin) - 1, 12, 1)
         If FechaContrato > FechaInicioAgui Then FechaInicioAgui = FechaContrato
         DiasProporcional = Day(FechaFin)
         MesesAguinaldo = (CDbl(FechaFin) - CDbl(FechaInicioAgui) + 1) / DiasMes 'DateDiff("m", CDbl(FechaInicioAgui), CDbl(FechaFin))
'         DiasAguinaldo = DateDiff("d", CDbl(FechaInicioAgui), CDbl(FechaFin))
         DiasAguinaldo = CalcularDiasVaca(CDate(FechaInicioAgui), FechaFin)
         If MesesAguinaldo < 12 Then
          DiasAguinaldo = DiasAguinaldo * 0.08333333
          DiasAguinaldoReal = DiasAguinaldo * 0.08333333
'          TotalAguinaldo = DiasAguinaldo * SalarioDiarioAgui
         Else
          DiasAguinaldo = 12 * 2.5
          DiasAguinaldoReal = 12 * 2.5
         End If


'//////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////DIAS LIQUIDACION/////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////

        
       DiasAntiguedad = CDbl(FechaFin) - CDbl(FechaContrato) + 1
      
       If Int(DiasAntiguedad / 365) > 3 Then
          
          If Int(DiasAntiguedad / 365) >= 6 Then
            DiasLiquidacion = DiasMes * 5
          Else
            DiasLiquidacion = DiasMes * 3
            DiasAntiguedadPro = 20 * ((DiasAntiguedad / 365) - 3)
            DiasLiquidacion = DiasLiquidacion + DiasAntiguedadPro
          End If
       ElseIf Int(DiasAntiguedad / 365) >= 1 Then
          DiasLiquidacion = DiasMes * (DiasAntiguedad / 365)
        
       Else
          TotalAntiguedad = 0
       End If



Me.LblDiasVacaciones.Caption = Format(CDbl(DiasVacaciones), "##,##0.00")
Me.LblDiasAguinaldo.Caption = Format(CDbl(DiasAguinaldo), "##,##0.00")
Me.LblDiasLiquidacion.Caption = Format(CDbl(DiasLiquidacion), "##,##0.00")
 



















End Sub


