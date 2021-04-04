VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepNominaDptoResumen 
   Caption         =   "Reporte de las Nominas "
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ArepNominasDptoResumen.dsx":0000
End
Attribute VB_Name = "ArepNominaDptoResumen"
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

Private Sub GroupHeader1_Format()

End Sub

Private Sub Produjo_Format()

ConteoEmpleados = 0

 If Me.Fldprodujo.Text = "S" Then
   Me.LblProdujo.Caption = "Empleados x Produccion"
   Me.LblSubTotal.Caption = "Empleados x Produccion"
 Else
   Me.LblProdujo.Caption = "Empleados x Salario Basico"
   Me.LblSubTotal.Caption = "Empleados x Salario Basico"
 End If
End Sub

