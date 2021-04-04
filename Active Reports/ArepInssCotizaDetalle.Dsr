VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepInssCotizaDetalle 
   Caption         =   "Reporte de las Nominas "
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19420
   SectionData     =   "ArepInssCotizaDetalle.dsx":0000
End
Attribute VB_Name = "ArepInssCotizaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_ReportEnd()
If Exportar = True Then
    Dim myExportObject As ActiveReportsExcelExport.ARExportExcel
    Dim Nombre As String
    
    Nombre = InputBox("Digite el Nombre del Archivo", "Sistema de Nominas")
    Set myExportObject = CreateObject("ActiveReportsExcelExport.ARExportExcel")
    myExportObject.FileName = "C:\" & Nombre & ".xls"
    myExportObject.Export Me.Pages
    Set myExportObject = Nothing
 End If
End Sub

