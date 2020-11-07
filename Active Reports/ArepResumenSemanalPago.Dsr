VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepResumen 
   Caption         =   "Resumen - Nomina Pago Mensual"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   16642
   SectionData     =   "ArepResumenSemanalPago.dsx":0000
End
Attribute VB_Name = "ArepResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub FinEmpleado_BeforePrint()

Me.txtTotalNetoBruto.Text = Format(CDbl(Me.txtSalBasico.Text) + CDbl(Me.txtSalDestajo.Text) + CDbl(Me.txtSeptimo.Text) + CDbl(Me.txtHorasExtras.Text) + CDbl(Me.txtOtrosIngresos.Text), "##,####.##")

End Sub

Private Sub MesNomina_Format()
 Select Case Me.Mes.Text
   Case "1"
      Me.LBLMES.Caption = "Enero"
   Case "2"
       Me.LBLMES.Caption = "Febrero"
   Case "3"
       Me.LBLMES.Caption = "Marzo"
   Case "4"
       Me.LBLMES.Caption = "Abril"
   Case "5"
       Me.LBLMES.Caption = "Mayo"
    Case "6"
       Me.LBLMES.Caption = "Junio"
   Case "7"
       Me.LBLMES.Caption = "Julio"
   Case "8"
       Me.LBLMES.Caption = "Agosto"
   Case "9"
       Me.LBLMES.Caption = "Septiembre"
   Case "10"
       Me.LBLMES.Caption = "Octubre"
   Case "11"
       Me.LBLMES.Caption = "Noviembre"
   Case "12"
       Me.LBLMES.Caption = "Diciembre"
 End Select

End Sub
