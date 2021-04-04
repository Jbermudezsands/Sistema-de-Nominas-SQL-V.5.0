VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepInssIr 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepIRINSS.dsx":0000
End
Attribute VB_Name = "ArepInssIr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
Dim Viaticos As Double, Fecha1 As Date, Fecha2 As Date, CodigoEmpleado As String, TotalDevengado As Double

Me.LblNombres.Caption = Me.Field2.Text & " " & Me.Field11.Text & " " & Me.Field12.Text

'Fecha1 = Year(FrmReportes.TxtFecha1.Value) & "-" & Month(FrmReportes.TxtFecha1.Value) & "-" & Day(FrmReportes.TxtFecha1.Value)
Fecha1 = FrmReportes.TxtFecha1.Value '& "-" & Month(FrmReportes.TxtFecha1.Value) & "-" & 15
Fecha2 = Year(FrmReportes.TxtFecha2.Value) & "-" & Month(FrmReportes.TxtFecha2.Value) & "-" & Day(FrmReportes.TxtFecha2.Value)
CodigoEmpleado = Me.Field1.Text

    If CodigoEmpleado = "S117100014" Then
      CodigoEmpleado = "S117100014"
   
    End If
 

  '/////////////////////////////BUSCO LOS INCENTIVOS /////////////////////////////////////////////
 MDIPrimero.AdoConsulta.ConnectionString = Conexion
 MDIPrimero.AdoConsulta.RecordSource = "SELECT MAX(DetalleIncentivo.NumIncentivo) AS NumIncentivo, SUM(DetalleIncentivo.Valor) AS Valor FROM  DetalleIncentivo INNER JOIN  Incentivo ON DetalleIncentivo.NumIncentivo = Incentivo.NumIncentivo INNER JOIN  Empleado ON Incentivo.CodEmpleado = Empleado.CodEmpleado INNER JOIN  Nomina ON DetalleIncentivo.NumNomina = Nomina.NumNomina  WHERE (Incentivo.CodTipoIncentivo = '14') AND (Empleado.CodEmpleado1 = '" & CodigoEmpleado & "') AND (Nomina.FechaNominaINI <= CONVERT(DATETIME, '" & Format(Fecha2, "yyyy-mm-dd") & "',102)) AND (Nomina.FechaNomina >= CONVERT(DATETIME, '" & Format(Fecha1, "yyyy-mm-dd") & "', 102)) "
 MDIPrimero.AdoConsulta.Refresh
 If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
  If Not IsNull(MDIPrimero.AdoConsulta.Recordset("Valor")) Then
     Viaticos = Format(MDIPrimero.AdoConsulta.Recordset("Valor"), "##,##0.00")
  End If
 End If
 
 
 TotalDevengado = Me.Field3.Text
 TotalDevengado = TotalDevengado - Viaticos
 Me.LblDevengado.Caption = Format(TotalDevengado, "##,##0.00")


End Sub

Private Sub GroupHeader1_Format()
 Select Case val(Me.FldMes.Text)
   Case "1"
      Me.LBLMES.Caption = "SALARIO MENSUAL PARA EL MES DE ENERO DEL " & Me.FldAño.Text
   Case "2"
       Me.LBLMES.Caption = "SALARIO MENSUAL PARA EL MES DE FEBRERO DEL " & Me.FldAño.Text
   Case "3"
       Me.LBLMES.Caption = "SALARIO MENSUAL PARA EL MES DE MARZO DEL " & Me.FldAño.Text
   Case "4"
       Me.LBLMES.Caption = "SALARIO MENSUAL PARA EL MES DE ABRIL DEL " & Me.FldAño.Text
   Case "5"
       Me.LBLMES.Caption = "SALARIO MENSUAL PARA EL MES DE MAYO DEL " & Me.FldAño.Text
    Case "6"
       Me.LBLMES.Caption = "SALARIO MENSUAL PARA EL MES DE JUNIO DEL " & Me.FldAño.Text
   Case "7"
       Me.LBLMES.Caption = "SALARIO MENSUAL PARA EL MES DE JULIO DEL " & Me.FldAño.Text
   Case "8"
       Me.LBLMES.Caption = "SALARIO MENSUAL PARA EL MES DE AGOSTO DEL " & Me.FldAño.Text
   Case "9"
       Me.LBLMES.Caption = "SALARIO MENSUAL PARA EL MES DE SEPTIEMBRE DEL " & Me.FldAño.Text
   Case "10"
       Me.LBLMES.Caption = "SALARIO MENSUAL PARA EL MES DE OCTUBRE DEL " & Me.FldAño.Text
   Case "11"
       Me.LBLMES.Caption = "SALARIO MENSUAL PARA EL MES DE NOVIEMBRE DEL " & Me.FldAño.Text
   Case "12"
       Me.LBLMES.Caption = "SALARIO MENSUAL PARA EL MES DE DICIEMBRE DEL " & Me.FldAño.Text
 End Select
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

