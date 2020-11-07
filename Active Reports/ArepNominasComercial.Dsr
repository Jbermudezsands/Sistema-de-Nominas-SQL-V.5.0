VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepNominaComercial 
   Caption         =   "Reporte de las Nominas "
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepNominasComercial.dsx":0000
End
Attribute VB_Name = "ArepNominaComercial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_ReportEnd()
If Exportar = True Then
    Dim myExportObject As ActiveReportsExcelExport.ARExportExcel
    Dim Nombre As String
    
'   Nombre = InputBox("Digite el Nombre del Archivo", "Sistema de Nominas")
    Set myExportObject = CreateObject("ActiveReportsExcelExport.ARExportExcel")
    myExportObject.FileName = Directorio
    myExportObject.Export Me.Pages
    Set myExportObject = Nothing
 End If
End Sub

Private Sub ActiveReport_ReportStart()
  
  MDIPrimero.DtaEmpresa.Refresh
 

    Me.LblTitulo.Caption = MDIPrimero.DtaEmpresa.Recordset("nombreempresa")
    Me.LblSubtitulo.Caption = MDIPrimero.DtaEmpresa.Recordset("Direccion") + " RUC: " + MDIPrimero.DtaEmpresa.Recordset("numeroruc")
    Me.ImgLogo.Picture = LoadPicture(MDIPrimero.DtaEmpresa.Recordset("RutaLogo"))
    Me.LblFecha.Caption = Format(Now, "dddddd")
   Me.LblDesde = FrmCalcularNomina.LblFecha1.Caption
   Me.LblHasta = FrmCalcularNomina.LblFecha2.Caption
    

End Sub

