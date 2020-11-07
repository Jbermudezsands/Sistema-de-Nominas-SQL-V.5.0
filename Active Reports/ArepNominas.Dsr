VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepNomina 
   Caption         =   "Reporte de las Nominas "
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepNominas.dsx":0000
End
Attribute VB_Name = "ArepNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NumeroNomina As Double, Fecha1 As Date, Fecha2 As Date


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
Dim Fecha1 As Date, Fecha2 As Date

If Quien = "ListadoNomina" Then
  NumeroNomina = FrmListNomina.DtaNominas.Recordset("NumNomina")
Else
  NumeroNomina = FrmCalcularNomina.DtaNomina.Recordset("NumNomina")
End If

  MDIPrimero.DtaEmpresa.Refresh
 

    Me.LblTitulo.Caption = MDIPrimero.DtaEmpresa.Recordset("nombreempresa")
    Me.LblSubtitulo.Caption = MDIPrimero.DtaEmpresa.Recordset("Direccion") '+ " RUC: " + MDIPrimero.DtaEmpresa.Recordset("numeroruc")
    If Dir(RutaLogo) <> "" Then
    Me.ImgLogo.Picture = LoadPicture(MDIPrimero.DtaEmpresa.Recordset("RutaLogo"))
    End If
    Me.LblFecha.Caption = Format(Now, "dddddd")
    Me.LblDesde.Caption = "Planilla de Pago # " & NumeroNomina & ", Corespondiente del " & Format(FrmCalcularNomina.LblFecha1.Caption, "dd/mm/yyyy") & " al " & Format(FrmCalcularNomina.LblFecha2.Caption, "dd/mm/yyyy")
'    Me.LblDesde = FrmCalcularNomina.LblFecha1.Caption
'    Me.LblHasta = FrmCalcularNomina.LblFecha2.Caption
    
    
End Sub

Private Sub ReportFooter_Format()
If Quien = "ListadoNomina" Then
  NumeroNomina = FrmListNomina.DtaNominas.Recordset("NumNomina")
Else
  NumeroNomina = FrmCalcularNomina.DtaNomina.Recordset("NumNomina")
End If
 Me.LblCanasta.Caption = Format(MontoDeduccionTotal(NumeroNomina, "03"), "##,##0.00")
 Me.LblAlimentos.Caption = Format(MontoDeduccionTotal(NumeroNomina, "04"), "##,##0.00")
 Me.LbllSindicatos.Caption = Format(MontoDeduccionTotal(NumeroNomina, "05"), "##,##0.00")
End Sub
