VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepMensualIR 
   Caption         =   "Reporte de Ir Mensual"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   19368
   SectionData     =   "ArepMesualIR.dsx":0000
End
Attribute VB_Name = "ArepMensualIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()

Dim Viaticos As Double, Fecha1 As Date, Fecha2 As Date, CodigoEmpleado As String, TotalDevengado As Double


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

