VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepProduccionBasico 
   Caption         =   "ArepoProduccionBasico"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "ArepProduccionBasico.dsx":0000
End
Attribute VB_Name = "ArepProduccionBasico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
 Dim Basico As Double, CodigoEmpleado As String, NumeroNomina As Double
 Dim HorasExtras As Double, Produccion As Double
 
 NumeroNomina = FrmReportes.TxtNNomina.Text
 CodigoEmpleado = Me.FldCodEmpleado1.Text
 Produccion = Me.FldProduccion.Text
 
 '----------------CONSULTO LOS SALARIOS DE LA NOMINA --------------------------------
 MDIPrimero.DtaConsulta.RecordSource = "SELECT DetalleNomina.NumNomina, Empleado.CodEmpleado1, Empleado.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.HorasExtras, DetalleNomina.BonoProduccion FROM DetalleNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado " & _
                                       "WHERE (DetalleNomina.NumNomina = " & NumeroNomina & ") AND (Empleado.CodEmpleado1 = '" & CodigoEmpleado & "') ORDER BY Empleado.CodEmpleado1"
 
 MDIPrimero.DtaConsulta.Refresh
 If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
   Basico = MDIPrimero.DtaConsulta.Recordset("SalarioBasico")
   HorasExtras = MDIPrimero.DtaConsulta.Recordset("HorasExtras")
 Else
    Basico = 0
   HorasExtras = 0
 End If
 
 Me.LblHorasExtras.Caption = Format(HorasExtras, "##,##0.00")
 Me.LblBasico.Caption = Format(Basico, "##,##0.00")
 Me.LblTotal.Caption = Format(Basico + HorasExtras, "##,##0.00")
 Me.LblDiferencia.Caption = Format(Produccion - Basico - HorasExtras, "##,##0.00")
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
