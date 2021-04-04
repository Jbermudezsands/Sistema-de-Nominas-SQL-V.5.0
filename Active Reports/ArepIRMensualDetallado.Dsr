VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepIRMensualDetallado 
   Caption         =   "Reporte de Vacaciones"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19420
   SectionData     =   "ArepIRMensualDetallado.dsx":0000
End
Attribute VB_Name = "ArepIRMensualDetallado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalSalarioVaca As Double, TotalIRVaca As Double, TotalINSSVaca As Double, GranTotalIngresos As Double, GranTotalInss As Double, GranTotalIR As Double, TotalSalarioBase As Double


Private Sub ActiveReport_ReportStart()
Dim Titulo As String, SubTitulo As String

MDIPrimero.DtaEmpresa.Refresh
Titulo = MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa")
SubTitulo = MDIPrimero.DtaEmpresa.Recordset("Direccion") + " RUC: " + MDIPrimero.DtaEmpresa.Recordset("numeroruc")


      Me.LblTitulo.Caption = Titulo
      Me.LblSubtitulo.Caption = SubTitulo
      Me.ImgLogo.Picture = LoadPicture(RutaLogo)
      
      Me.lblFecha.Caption = "Desde " + Format(Frm13Vaca.TxtFINIVaca.Value, "mm/dd/yyyy") + " Hasta " + Format(Frm13Vaca.TxtFFinVaca.Value, "mm/dd/yyyy")
      Me.LblFechaHoy = Format(Now, "dddddd")

  TotalSalarioVaca = 0
  TotalIRVaca = 0
  TotalINSSVaca = 0
  GranTotalIngresos = 0
  GranTotalInss = 0
  GranTotalIR = 0
  TotalSalarioBase = 0
End Sub

Private Sub Detail_Format()
Dim CodEmpleado As Double, FechaIni As Date, FechaFin As Date, CodTipoNomina As String, SubTotal As Double
Dim TotalIngresos As Double, SalarioMensual As Double, SalarioVacaciones As Double, InssVaca As Double, Inss As Double, TotalInss As Double
Dim IrVaca As Double, IR As Double, TotalIr As Double, SalarioBase As Double, Fecha1 As Date, Fecha2 As Date
Dim SqlSalarios As String, IrAcumulado As Double, InssAcumulado As Double, SalarioAcumulado As Double

If Me.FldCodEmpleado.Text = "" Then
 Exit Sub
End If
CodEmpleado = Me.FldCodEmpleado.Text
FechaIni = MesIni(FrmReportes.Combo1.Text, FrmReportes.DBCAño.Text)
FechaFin = MesIni(FrmReportes.Combo2.Text, FrmReportes.DBAño2.Text)
FechaFin = DateSerial(Year(FechaFin), Month(FechaFin) + 1, 0)

CodTipoNomina = Me.FldCodTipoNomina.Text


MDIPrimero.DtaConsulta.RecordSource = "SELECT  * From Empleado Where (CodEmpleado = " & CodEmpleado & ")"
MDIPrimero.DtaConsulta.Refresh
If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
  Me.LblCedula.Caption = MDIPrimero.DtaConsulta.Recordset("NumCedula")
  Me.LblInss.Caption = MDIPrimero.DtaConsulta.Recordset("NumeroInss")
End If
 
MDIPrimero.DtaConsulta.RecordSource = "SELECT SUM(DetalleNomVaca.Inss) AS Inss, SUM(DetalleNomVaca.Ir) AS Ir, SUM(DetalleNomVaca.TotalDevengado) AS TotalDevengado, DetalleNomVaca.CodEmpleado FROM DetalleNomVaca INNER JOIN  NomVaca ON DetalleNomVaca.NumNomVaca = NomVaca.NumNomVaca WHERE (NomVaca.FechaAplica BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) GROUP BY DetalleNomVaca.CodEmpleado Having (DetalleNomVaca.CodEmpleado = " & CodEmpleado & ")"
MDIPrimero.DtaConsulta.Refresh
If Not MDIPrimero.DtaConsulta.Recordset.EOF Then

    '--------------------------------------------------------------
    '---------------------BUSCO LA FECHA DEL PERIODO ---------------
    '---------------------------------------------------------------
    
    
    
    Fecha1 = BuscaIncioPeriodo(Year(FechaIni), Format(Month(FechaIni), "0#"), CodTipoNomina)
    Fecha2 = BuscaFinPeriodo(Year(FechaFin), Format(Month(FechaFin), "0#"), CodTipoNomina)
    
    If CodEmpleado = "11440" Then
      CodEmpleado = "11440"
    End If
      
 
    If Not IsNull(MDIPrimero.DtaConsulta.Recordset("Ir")) Then
      IrVaca = MDIPrimero.DtaConsulta.Recordset("Ir")
      Me.LblIRVacaciones.Caption = Format(IrVaca, "##,##0.00")
     
    End If
    If Not IsNull(MDIPrimero.DtaConsulta.Recordset("Inss")) Then
      InssVaca = MDIPrimero.DtaConsulta.Recordset("Inss")
      Me.LblInssVacaciones.Caption = Format(InssVaca, "##,##0.00")
    
    End If
    If Not IsNull(MDIPrimero.DtaConsulta.Recordset("TotalDevengado")) Then
     SalarioVacaciones = MDIPrimero.DtaConsulta.Recordset("TotalDevengado")
    End If
    Me.LblSalarioVacaciones.Caption = Format(SalarioVacaciones, "##,##0.00")
    
    
    '--------------------------------------------------------------------------------------------
    '----------------------------------BUSCO LA FECHA DEL SALARIO -------------------------------
    '--------------------------------------------------------------------------------------------
    Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
    Dim FechaV1 As Date, FechaV2 As Date
    
    FMes (FrmReportes.Combo1.Text)
    Mes1 = Format(Nmes - 1, "0#")
    FMes (FrmReportes.Combo2.Text)
    Mes2 = Format(Nmes - 1, "0#")
    Año1 = val(FrmReportes.DBCAño.Text)
    Año2 = val(FrmReportes.DBAño2.Text)
    CodTipoNomina = FrmReportes.DBTipoNominas.Columns(0).Text
    
    
    FrmReportes.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
    FrmReportes.AdoBusca.Refresh
     If Not FrmReportes.AdoBusca.Recordset.EOF Then
       FechaV1 = FrmReportes.AdoBusca.Recordset("Inicio")
     End If
     
    FrmReportes.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
    FrmReportes.AdoBusca.Refresh
     If Not FrmReportes.AdoBusca.Recordset.EOF Then
       FrmReportes.AdoBusca.Recordset.MoveLast
       FechaV2 = FrmReportes.AdoBusca.Recordset("Final")
     End If
    
    
    SqlSalarios = "SELECT DISTINCT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.HorasExtras) AS HorasExttras, SUM(DetalleNomina.BonoProduccion) AS BonoProduccion, SUM(DetalleNomina.SeptimoDia) AS SeptimoDia, SUM(DetalleNomina.OtrosIngresos) AS OtrosIngresos, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.Incentivos) AS Incentivos, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.Comisiones + DetalleNomina.HorasExtras + DetalleNomina.BonoProduccion + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos+DetalleNomina.IncetivoProduccion) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin,Nomina.Mes AS MES,Nomina.Ano AS AÑO,SUM(DetalleNomina.Comisiones) AS Comisiones,SUM(DetalleNomina.MontoIR) AS MontoIR,SUM(DetalleNomina.MontoINSS) As MontoINSS  " & _
                "FROM  DetalleNomina INNER JOIN  Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = " & CodEmpleado & ") AND (MAX(Nomina.FechaNomina) <= CONVERT(DATETIME, '" & Format(FechaV2, "yyyy-mm-dd") & "', 102)) AND (MIN(Nomina.FechaNominaINI) >= CONVERT(DATETIME, '" & Format(FechaV1, "yyyy-mm-dd") & "', 102))"
                           
    MDIPrimero.DtaConsulta2.RecordSource = SqlSalarios
    MDIPrimero.DtaConsulta2.Refresh
    If Not MDIPrimero.DtaConsulta2.Recordset.EOF Then
                                  
       IrAcumulado = Format(MDIPrimero.DtaConsulta2.Recordset("MontoIR"), "####0.00")
       InssAcumulado = Format(MDIPrimero.DtaConsulta2.Recordset("MontoINSS"), "####0.00")
       SalarioAcumulado = Format(MDIPrimero.DtaConsulta2.Recordset("TotalIngresos"), "####0.00") + SalarioVacaciones
                                    
    End If

Else
   IrVaca = 0
   InssVaca = 0
   SalarioVacaciones = 0
   TotalInss = 0
   TotalIngresos = 0
   TotalIr = 0
   SalarioAcumulado = 0
   
   Me.LblSalarioVacaciones.Caption = Format(SalarioVacaciones, "##,##0.00")
   Me.LblIRVacaciones.Caption = Format(IrVaca, "##,##0.00")
   Me.LblInssVacaciones.Caption = Format(InssVaca, "##,##0.00")
   Me.LblSalarioBaseVaca.Caption = Format(SalarioAcumulado, "##,##0.00")
End If

  IR = Me.FldIr
  Inss = Me.FldInss.Text
  SalarioMensual = Me.FldSalario.Text
  
  TotalInss = Inss + InssVaca
  TotalIngresos = SalarioMensual + SalarioVacaciones
  TotalIr = IrVaca + IR
'  SalarioBase = TotalIngresos - TotalInss
  SalarioBase = SalarioMensual - Inss
  
  Me.LblTotalInss.Caption = Format(TotalInss, "##,##0.00")
  Me.LblTotalIngresos.Caption = Format(TotalIngresos, "##,##0.00")
  Me.LblTotalIr.Caption = Format(TotalIr, "##,##0.00")
  Me.LblSalarioBase.Caption = Format(SalarioBase, "##,##0.00")
  Me.LblSalarioBaseVaca.Caption = Format(SalarioAcumulado, "##,##0.00")
  
  '/////////////////TOTALES ///////////////////////////////////////////////
  TotalSalarioVaca = TotalSalarioVaca + SalarioVacaciones
  TotalINSSVaca = TotalINSSVaca + InssVaca
  TotalIRVaca = TotalIRVaca + IrVaca
  GranTotalIngresos = GranTotalIngresos + TotalIngresos
  GranTotalInss = GranTotalInss + TotalInss
  GranTotalIR = GranTotalIR + TotalIr
  TotalSalarioBase = TotalSalarioBase + SalarioBase

End Sub

Public Function MesIni(MesLetra As String, Ano As Double) As Date
  Select Case MesLetra
     Case "Enero": MesIni = "01/01/" & Ano
     Case "Febrero": MesIni = "01/02/" & Ano
     Case "Marzo": MesIni = "01/03/" & Ano
     Case "Abril": MesIni = "01/04/" & Ano
     Case "Mayo": MesIni = "01/05/" & Ano
     Case "Junio": MesIni = "01/06/" & Ano
     Case "Julio": MesIni = "01/07/" & Ano
     Case "Agosto": MesIni = "01/08/" & Ano
     Case "Septiembre": MesIni = "01/09/" & Ano
     Case "Octubre": MesIni = "01/10/" & Ano
     Case "Noviembre": MesIni = "01/11/" & Ano
     Case "Diciembre": MesIni = "01/12/" & Ano

  End Select

End Function





Private Sub ReportFooter_Format()
  '/////////////////TOTALES ///////////////////////////////////////////////
  Me.LblTotalSalarioVacaciones.Caption = Format(TotalSalarioVaca, "##,##0.00")
  Me.LblTotalInssVacaciones.Caption = Format(TotalINSSVaca, "##,##0.00")
  Me.LblTotalIRVacaciones.Caption = Format(TotalIRVaca, "##,##0.00")
  Me.LblGranTotalIngresos = Format(GranTotalIngresos, "##,##0.00")
  Me.LblGranTotalInss.Caption = Format(GranTotalInss, "##,##0.00")
  Me.LblGranTotalIr.Caption = Format(GranTotalIR, "##,##0.00")
  Me.LblGranSalarioBase.Caption = Format(TotalSalarioBase, "##,##0.00")
  

   
End Sub
