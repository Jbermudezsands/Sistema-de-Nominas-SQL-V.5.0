VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepNominaProduccionLegal 
   Caption         =   "Reporte de las Nominas "
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19420
   SectionData     =   "ArepNominasProduccionLegal.dsx":0000
End
Attribute VB_Name = "ArepNominaProduccionLegal"
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

If Quien = "ListadoNomina" Then
  NumeroNomina = FrmListNomina.DtaNominas.Recordset("NumNomina")
  Me.LblDesde.Caption = "Planilla de Pago # " & FrmListNomina.DtaNominas.Recordset("NumNomina") & ", Correspondiente del " & Format(FrmListNomina.DtaNominas.Recordset("FechaNominaINI"), "dddddd") & " al " & Format(FrmListNomina.DtaNominas.Recordset("FechaNomina"), "dddddd")

        FechaIni = Format(FrmListNomina.DtaNominas.Recordset("FechaNominaINI"), "dddddd")
        FechaFin = Format(FrmListNomina.DtaNominas.Recordset("FechaNomina"), "dddddd")
        Me.LblDesde.Caption = FechaIni
        Me.LblHasta.Caption = FechaFin
        Me.lblFecha = Format(Now, "dddddd")
        Me.LblTitulo.Caption = Titulo
        Me.LblSubtitulo.Caption = SubTitulo
        If Dir(RutaLogo) <> "" Then
          ArepNominaProduccionLegal.ImgLogo.Picture = LoadPicture(RutaLogo)
        End If



Else
  Me.LblDesde.Caption = "Planilla de Pago # " & FrmCalcularNomina.NumeroNominas & ", Correspondiente del " & FrmCalcularNomina.LblFecha1.Caption & " al " & FrmCalcularNomina.LblFecha2.Caption

        Me.LblDesde.Caption = FrmCalcularNomina.LblFecha1.Caption
        Me.LblHasta.Caption = FrmCalcularNomina.LblFecha2.Caption
        Me.lblFecha = Format(Now, "dddddd")
        Me.LblTitulo.Caption = Titulo
        Me.LblSubtitulo.Caption = SubTitulo
        If Dir(RutaLogo) <> "" Then
          Me.ImgLogo.Picture = LoadPicture(RutaLogo)
        End If
        
        
End If


End Sub

