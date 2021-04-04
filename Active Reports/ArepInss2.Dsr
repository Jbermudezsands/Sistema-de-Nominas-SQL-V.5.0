VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepInss2 
   Caption         =   "Reporte del Inss 2"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ArepInss2.dsx":0000
End
Attribute VB_Name = "ArepInss2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportEnd()
Unload Me.SubDetalle.object
Set Me.SubDetalle.object = Nothing

If Exportar = True Then
    Dim myExportObject As ActiveReportsExcelExport.ARExportExcel
    Dim Nombre As String
    
    Set myExportObject = CreateObject("ActiveReportsExcelExport.ARExportExcel")
    myExportObject.FileName = Directorio
    myExportObject.Export Me.Pages
    Set myExportObject = Nothing
 End If
 
End Sub

Private Sub ActiveReport_ReportStart()

Set Me.SubDetalle.object = New ArepSubInss2

End Sub


Private Sub GroupFooter1_Format()
Dim sql As String
Dim CodEmpleado As Integer
If Not Me.FldCodEmpledo.Text = "CodEmpleado" Then

CodEmpleado = Me.FldCodEmpledo.Text
sql = "SELECT     TOP 100 PERCENT dbo.Empleado.Nombre1 + N' ' + dbo.Empleado.Nombre2 + N' ' + dbo.Empleado.Apellido1 + N' ' + dbo.Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "                       dbo.DetalleNomina.CodEmpleado, dbo.DetalleNomina.MontoINSS, dbo.DetalleNomina.INSSPatronal," & vbLf
sql = sql & "                dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
sql = sql & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
sql = sql & "                       dbo.DetalleNomina.INATEC, dbo.Empleado.CodInss, dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.INSSPatronal AS TotalInss," & vbLf
sql = sql & "                      dbo.Empleado.CodEmpleado1 , dbo.DetalleNomina.NumNomina, dbo.Nomina.FechaNomina, dbo.Cargo.Cargo" & vbLf
sql = sql & "FROM         dbo.Nomina INNER JOIN" & vbLf
sql = sql & "                      dbo.Grupo INNER JOIN" & vbLf
sql = sql & "                      dbo.Cargo INNER JOIN" & vbLf
sql = sql & "                      dbo.TipoNomina INNER JOIN" & vbLf
sql = sql & "                      dbo.Empleado ON dbo.TipoNomina.CodTipoNomina = dbo.Empleado.CodTipoNomina ON dbo.Cargo.CodCargo = dbo.Empleado.CodCargo INNER JOIN" & vbLf
sql = sql & "                      dbo.DetalleNomina ON dbo.Empleado.CodEmpleado = dbo.DetalleNomina.CodEmpleado ON dbo.Grupo.CodGrupo = dbo.Empleado.CodGrupo ON" & vbLf
sql = sql & "                      dbo.TipoNomina.CodTipoNomina = dbo.Nomina.CodTipoNomina And dbo.Nomina.NumNomina = dbo.DetalleNomina.NumNomina" & vbLf
sql = sql & "WHERE(Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1Reporte & "', 102) AND CONVERT(DATETIME, '" & Fecha2Reporte & "', 102))AND(dbo.DetalleNomina.CodEmpleado = '" & CodEmpleado & "')"
sql = sql & "ORDER BY dbo.Empleado.CodEmpleado1, dbo.Nomina.FechaNomina"

Me.SubDetalle.object.AdoNomina.ConnectionString = ConexionReporte
SubDetalle.object.AdoNomina.Source = sql
End If
End Sub

