VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepTotalSemanaPago 
   Caption         =   "Resumen - Nomina Pago Mensual"
   ClientHeight    =   10950
   ClientLeft      =   165
   ClientTop       =   615
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19315
   SectionData     =   "ArepTotalSemanalPago.dsx":0000
End
Attribute VB_Name = "ArepTotalSemanaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public sSemanasCotizadas As String
Public sSemana1 As String
Public sSemana2 As String
Public sSemana3 As String
Public sSemana4 As String
Public sSemana5 As String





Private Sub ActiveReport_PageStart()

'Dim cnDB As New ADODB.Connection
'Dim rsBD As New ADODB.Recordset
'Dim iCont As Integer
'
'cnDB.ConnectionString = Conexion
'cnDB.Open
'
'rsBD.Open "SELECT * FROM Fecha_Planilla WHERE Inicio >= CONVERT(DATETIME, '" & Format(Me.lblFechaDesde.Caption, "yyyy-mm-dd") & " 00:00:00', 102) AND Final <=CONVERT(DATETIME, '" & Format(Me.lblFechaHasta.Caption, "yyyy-mm-dd") & " 00:00:00', 102) AND CodTipoNomina ='" & Me.lblNoNomina.Caption & "' ORDER BY Periodo ASC", cnDB
'
'iCont = 1
'
' Do While Not rsBD.EOF
'   saPeriodos(iCont) = rsBD.Fields("Periodo")
'
'   rsBD.MoveNext
'Loop
'
'rsBD.Close
'cnDB.Close


End Sub

Private Sub ActiveReport_ReportEnd()

If Exportar = True Then
    Dim myExportObject As ActiveReportsExcelExport.ARExportExcel
    Dim Nombre As String
    
    Set myExportObject = CreateObject("ActiveReportsExcelExport.ARExportExcel")
    myExportObject.FileName = Directorio
    myExportObject.Export Me.Pages
    Set myExportObject = Nothing
 End If


End Sub

Private Sub ActiveReport_ReportStart()

Dim cnDB As New ADODB.Connection
Dim rsBD As New ADODB.Recordset
Dim iCont As Integer

cnDB.ConnectionString = Conexion
cnDB.Open

rsBD.Open "SELECT * FROM Fecha_Planilla WHERE Inicio >= CONVERT(DATETIME, '" & Format(Me.lblFechaDesde.Caption, "yyyy-mm-dd") & " 00:00:00', 102) AND Final <=CONVERT(DATETIME, '" & Format(Me.lblFechaHasta.Caption, "yyyy-mm-dd") & " 00:00:00', 102) AND CodTipoNomina ='" & Me.lblNoNomina.Caption & "' ORDER BY Periodo ASC", cnDB

iCont = 1

 Do While Not rsBD.EOF
   saPeriodos(iCont) = rsBD.Fields("Periodo")
   iCont = iCont + 1
   
   rsBD.MoveNext
Loop

rsBD.Close
cnDB.Close

sSemana1 = "0"
sSemana2 = "0"
sSemana3 = "0"
sSemana4 = "0"
sSemana5 = "0"


End Sub

Private Sub FinEmpleado_BeforePrint()

Dim cnDB As New ADODB.Connection
Dim rsBD As New ADODB.Recordset
Dim iCont As Integer

Me.txtTotalNetoBruto.Text = Format(CDbl(Me.txtSalBasico.Text) + CDbl(Me.txtSalDestajo.Text) + CDbl(Me.txtSeptimo.Text) + CDbl(Me.txtHorasExtras.Text) + CDbl(Me.txtOtrosIngresos.Text) + CDbl(Me.txtAntiguedad.Text), "##,####.##")

cnDB.ConnectionString = Conexion
cnDB.Open

SQL = "SELECT  Empleado.CodEmpleado1, Empleado.CodEmpleado," & vbLf
    SQL = SQL & "Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Departamento.Departamento," & vbLf
    SQL = SQL & "Empleado.Numeroinss, Empleado.Activo, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Historico.FechaContrato, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.OtrosIngresos, DetalleNomina.SeptimoDia," & vbLf
    SQL = SQL & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo AS Sueldo," & vbLf
    SQL = SQL & "DetalleNomina.HorasExtras, DetalleNomina.MontoINSS, Nomina.FechaNominaINI, Nomina.FechaNomina," & vbLf
    SQL = SQL & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo + DetalleNomina.HorasExtras - DetalleNomina.MontoInss" & vbLf
    SQL = SQL & "AS Neto, Nomina.Mes, Nomina.Ano, Nomina.Periodo" & vbLf
    SQL = SQL & "FROM         Empleado INNER JOIN" & vbLf
    SQL = SQL & "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN" & vbLf
    SQL = SQL & "Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN" & vbLf
    SQL = SQL & "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN" & vbLf
    SQL = SQL & "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina" & vbLf
    SQL = SQL & "WHERE (Nomina.FechaNomina BETWEEN '" & Format(Me.lblFechaDesde.Caption, "yyyymmdd") & "' AND '" & Format(Me.lblFechaHasta.Caption, "yyyymmdd") & "') AND (Nomina.CodTipoNomina = '" & Me.lblNoNomina.Caption & "') AND (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo <> 0) AND (Empleado.CodEmpleado1 LIKE '" & Me.txtCodEmpleado1.Text & "') " & vbLf
    SQL = SQL & "ORDER BY Nomina.Periodo ASC"
' _
'           "WHERE Inicio >= CONVERT(DATETIME, '" & Format(Me.lblFechaDesde.Caption, "yyyy-mm-dd") & " 00:00:00', 102) AND Final <=CONVERT(DATETIME, '" & Format(Me.lblFechaHasta.Caption, "yyyy-mm-dd") & " 00:00:00', 102) AND CodTipoNomina ='" & Me.lblNoNomina.Caption & "' AND Empleado.CodEmpleado1 ='" & Me.txtCodEmpleado1.Text & "' AND Empleado.Activo =1 ORDER BY Nomina.Periodo ASC", cnDB

iCont = 1




rsBD.Open SQL, cnDB

Do While Not rsBD.EOF

If Me.lblNoNomina.Caption = "02" Then

Select Case rsBD.Fields("Periodo")


Case saPeriodos(1):

     sSemana1 = "1"
     
     
Case saPeriodos(2):
    sSemana2 = "1"
    
Case saPeriodos(3):
    sSemana3 = "1"

Case saPeriodos(4):
    sSemana4 = "1"


Case saPeriodos(5):
    sSemana5 = "1"

End Select


Else

Select Case rsBD.Fields("Periodo")


Case saPeriodos(1):

     sSemana1 = "1"
     sSemana2 = "1"
     
Case saPeriodos(2):
     sSemana3 = "1"
     sSemana4 = "1"

End Select


End If





rsBD.MoveNext

Loop


Me.lblSemanasCotizadas.Caption = sSemana1 & sSemana2 & sSemana3 & sSemana4 & sSemana5


rsBD.Close
cnDB.Close





sSemana1 = "0"
sSemana2 = "0"
sSemana3 = "0"
sSemana4 = "0"
sSemana5 = "0"


If IsNumeric(Me.txtTarifaHoraria.Text) Then

   Me.txtSalarioBasico.Text = CDbl(Me.txtTarifaHoraria.Text) * 30.4167 * 8
   
Else
   
   Me.txtSalarioBasico.Text = "0"

End If





End Sub





Private Sub MesNomina_Format()
 Select Case Me.Mes.Text
   Case "1"
      Me.LblMes.Caption = "Enero"
   Case "2"
       Me.LblMes.Caption = "Febrero"
   Case "3"
       Me.LblMes.Caption = "Marzo"
   Case "4"
       Me.LblMes.Caption = "Abril"
   Case "5"
       Me.LblMes.Caption = "Mayo"
    Case "6"
       Me.LblMes.Caption = "Junio"
   Case "7"
       Me.LblMes.Caption = "Julio"
   Case "8"
       Me.LblMes.Caption = "Agosto"
   Case "9"
       Me.LblMes.Caption = "Septiembre"
   Case "10"
       Me.LblMes.Caption = "Octubre"
   Case "11"
       Me.LblMes.Caption = "Noviembre"
   Case "12"
       Me.LblMes.Caption = "Diciembre"
 End Select

End Sub

