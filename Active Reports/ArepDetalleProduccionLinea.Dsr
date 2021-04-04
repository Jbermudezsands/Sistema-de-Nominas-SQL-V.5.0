VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepProduccionLinea 
   Caption         =   "reporte de Produccion"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "ArepDetalleProduccionLinea.dsx":0000
End
Attribute VB_Name = "ArepProduccionLinea"
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

