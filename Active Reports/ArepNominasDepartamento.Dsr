VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepNominaDepartamento 
   Caption         =   "Reporte de las Nominas "
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ArepNominasDepartamento.dsx":0000
End
Attribute VB_Name = "ArepNominaDepartamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub ActiveReport_ReportStart()
Me.LblDesde.Caption = "Planilla de Pago # " & FrmCalcularNomina.NumeroNominas & ", Correspondiente del " & FrmCalcularNomina.LblFecha1.Caption & " al " & FrmCalcularNomina.LblFecha2.Caption
End Sub

