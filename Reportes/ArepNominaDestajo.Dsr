VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepNominaDestajo 
   Caption         =   "Reporte Detalle Salario"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19368
   SectionData     =   "ArepNominaDestajo.dsx":0000
End
Attribute VB_Name = "ArepNominaDestajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public TotalNeto1 As Double, TotalNeto2 As Double, NumeroNomina As Double

Public TotalViatico As Double, TotalIncentivo As Double

Private Sub ActiveReport_ReportStart()
If Quien = "ListadoNomina" Then
   Me.LblTitulo.Caption = Titulo
   Me.LblSubtitulo.Caption = SubTitulo & "  Reimpresion"
   If Dir(RutaLogo) <> "" Then
      Me.ImgLogo.Picture = LoadPicture(RutaLogo)
   End If
   Me.LblDesde.Caption = "Desde " & FechaInicio & " Hasta " & FechaFinal
End If

End Sub

Private Sub Detail_Format()
  Dim NetoPagar As Double, NumNomina As Double, CodEmpleado As String

    CodEmpleado = Me.Codigo.Text
    NumNomina = Me.FldNumeroNomina.Text
    

    NetoPagar = Me.FldNeto.Text
    Me.FldNetoDolar.Text = Format(NetoPagar / TasaCambioR, "##,##0.00")
    
    TotalNeto1 = TotalNeto1 + (NetoPagar / TasaCambioR)
    TotalNeto2 = TotalNeto2 + (NetoPagar / TasaCambioR)
   
 
    If Field21.Text = "" Then
        Field21.Text = "0.00"
    End If
    
' '/////////////////////////////BUSCO TODOS LOS INCENTIVOS QUE NO SON  EXCENTO /////////////////////////////////////////////
' MDIPrimero.AdoConsulta.ConnectionString = Conexion
' MDIPrimero.AdoConsulta.RecordSource = "SELECT MAX(DetalleIncentivo.NumIncentivo) AS NumIncentivo, SUM(DetalleIncentivo.Valor) AS Valor FROM DetalleIncentivo INNER JOIN Incentivo ON DetalleIncentivo.NumIncentivo = Incentivo.NumIncentivo INNER JOIN Empleado ON Incentivo.CodEmpleado = Empleado.CodEmpleado  " & _
'                                       "WHERE (Incentivo.CodTipoIncentivo <> N'14') AND  (Empleado.CodEmpleado1 = '" & CodEmpleado & "') AND (DetalleIncentivo.NumNomina = " & NumNomina & ")"
' MDIPrimero.AdoConsulta.Refresh
' If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
'    Me.LblMontoIncentivo.Caption = Format(MDIPrimero.AdoConsulta.Recordset("Valor"), "##,##0.00")
'    If Not IsNull(MDIPrimero.AdoConsulta.Recordset("Valor")) Then
'     TotalIncentivo = TotalIncentivo + MDIPrimero.AdoConsulta.Recordset("Valor")
'    End If
' End If
'
'
'  '/////////////////////////////BUSCO LOS INCENTIVOS /////////////////////////////////////////////
' MDIPrimero.AdoConsulta.ConnectionString = Conexion
' MDIPrimero.AdoConsulta.RecordSource = "SELECT MAX(DetalleIncentivo.NumIncentivo) AS NumIncentivo, SUM(DetalleIncentivo.Valor) AS Valor FROM DetalleIncentivo INNER JOIN Incentivo ON DetalleIncentivo.NumIncentivo = Incentivo.NumIncentivo INNER JOIN Empleado ON Incentivo.CodEmpleado = Empleado.CodEmpleado  " & _
'                                       "WHERE (Incentivo.CodTipoIncentivo = '14') AND (Empleado.CodEmpleado1 = '" & CodEmpleado & "') AND (DetalleIncentivo.NumNomina = " & NumNomina & ")"
' MDIPrimero.AdoConsulta.Refresh
' If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
'  Me.LblViaticos.Caption = Format(MDIPrimero.AdoConsulta.Recordset("Valor"), "##,##0.00")
'  If Not IsNull(MDIPrimero.AdoConsulta.Recordset("Valor")) Then
'   TotalViatico = TotalViatico + MDIPrimero.AdoConsulta.Recordset("Valor")
'  End If
' End If
 
    
End Sub

Private Sub GroupFooter1_Format()
Dim NetoPagar As Double

NetoPagar = Me.Field71.Text


Me.FldNetoDolar2.Text = Format(NetoPagar / TasaCambioR, "##,##0.00")
Field50.Text = Format(Field50.Text, "###,##0.00")
Field51.Text = Format(Field51.Text, "###,##0.00")
Field77.Text = Format(Field77.Text, "###,##0.00")
Field56.Text = Format(Field56.Text, "###,##0.00")
Field57.Text = Format(Field57.Text, "###,##0.00")
Field58.Text = Format(Field58.Text, "###,##0.00")
Field59.Text = Format(Field59.Text, "###,##0.00")
Field60.Text = Format(Field60.Text, "###,##0.00")
Field61.Text = Format(Field61.Text, "###,##0.00")
Field62.Text = Format(Field62.Text, "###,##0.00")
Field63.Text = Format(Field63.Text, "###,##0.00")
Field64.Text = Format(Field64.Text, "###,##0.00")
Field65.Text = Format(Field65.Text, "###,##0.00")
Field66.Text = Format(Field66.Text, "###,##0.00")
'Field67.Text = Format(Field67.Text, "###,##0.00")
'Field68.Text = Format(Field68.Text, "###,##0.00")
Field69.Text = Format(Field69.Text, "###,##0.00")
Field71.Text = Format(Field71.Text, "###,##0.00")
End Sub

Private Sub GroupFooter2_Format()
Me.LblTotalIncentivo.Caption = Format(TotalIncentivo, "##,##0.00")
Me.LblTotalViaticos.Caption = Format(TotalViatico, "##,##0.00")
Me.FldNetoDolar1.Text = Format(TotalNeto1, "##,##0.00")
TotalNeto1 = 0

End Sub

Private Sub GroupHeader2_Format()
TotalViatico = 0
TotalIncentivo = 0
End Sub

Private Sub PageHeader_Format()
Dim FechaNomina As Date


FechaNomina = Me.FldFechaNomina.Text
TasaCambioR = BuscaTasaCambio(FechaNomina)
Me.LblFecha.Caption = Format(DateTime.Now, "dddddd")
Me.LblTasaCambio.Caption = TasaCambioR


TotalNeto1 = 0
TotalNeto2 = 0

End Sub

Private Sub ReportFooter_Format()
Field94.Text = Format(Field94.Text, "###,##0.00")
Field93.Text = Format(Field93.Text, "###,##0.00")
Field99.Text = Format(Field99.Text, "###,##0.00")
Field98.Text = Format(Field98.Text, "###,##0.00")
Field91.Text = Format(Field91.Text, "###,##0.00")
Field79.Text = Format(Field79.Text, "###,##0.00")
Field80.Text = Format(Field80.Text, "###,##0.00")
Field83.Text = Format(Field83.Text, "###,##0.00")
Field84.Text = Format(Field84.Text, "###,##0.00")
Field52.Text = Format(Field52.Text, "###,##0.00")
Field97.Text = Format(Field97.Text, "###,##0.00")
Field100.Text = Format(Field100.Text, "###,##0.00")
Field88.Text = Format(Field88.Text, "###,##0.00")
Field101.Text = Format(Field101.Text, "###,##0.00")
Field102.Text = Format(Field102.Text, "###,##0.00")
Field103.Text = Format(Field103.Text, "###,##0.00")


MDIPrimero.DtaConsulta.RecordSource = "SELECT  Nomina.NumNomina, SUM(DetalleNomina.INSSPatronal) AS INSSPatronal, SUM(DetalleNomina.IRPatronal) AS IRPatronal, SUM(DetalleNomina.INATEC) AS INATEC FROM Nomina INNER JOIN Grupo INNER JOIN Cargo INNER JOIN TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN  DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON  TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado  GROUP BY Nomina.NumNomina  Having (Nomina.NumNomina = " & NumeroNomina & ")"
MDIPrimero.DtaConsulta.Refresh
If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
   Me.Field52.Text = Format(MDIPrimero.DtaConsulta.Recordset("INSSPatronal"), "##,##0.00")
   Me.Field151.Text = Format(MDIPrimero.DtaConsulta.Recordset("INATEC"), "##,##0.00")

End If

'* TasaCambioR
End Sub

