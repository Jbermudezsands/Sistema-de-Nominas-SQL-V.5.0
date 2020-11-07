VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepNominasHM 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepNominasHM.dsx":0000
End
Attribute VB_Name = "ArepNominasHM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public TotalNeto1 As Double, TotalNeto2 As Double, NumeroNomina As Double

Private Sub Detail_Format()
  Dim NetoPagar As Double

    If Me.FldNeto.Text <> "" Then
     NetoPagar = Me.FldNeto.Text
     Me.FldNetoDolar.Text = Format(NetoPagar / TasaCambioR, "##,##0.00")
      TotalNeto1 = TotalNeto1 + (NetoPagar / TasaCambioR)
      TotalNeto2 = TotalNeto2 + (NetoPagar / TasaCambioR)
    End If
    

    
 
    If Field21.Text = "" Then
        Field21.Text = "0.00"
    End If
End Sub

Private Sub GroupFooter1_Format()
Dim NetoPagar As Double

NetoPagar = Me.Field71.Text

If TasaCambioR <> 0 Then
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
    Field67.Text = Format(Field67.Text, "###,##0.00")
    Field68.Text = Format(Field68.Text, "###,##0.00")
    Field69.Text = Format(Field69.Text, "###,##0.00")
    Field71.Text = Format(Field71.Text, "###,##0.00")
End If
End Sub

Private Sub GroupFooter2_Format()
Me.FldNetoDolar1.Text = Format(TotalNeto1, "##,##0.00")
TotalNeto1 = 0

End Sub

Private Sub PageHeader_Format()
Dim FechaNomina As Date

If Me.FldFechaNomina.Text <> "" Then
    FechaNomina = Me.FldFechaNomina.Text
    TasaCambioR = BuscaTasaCambio(FechaNomina)
    Me.LblFecha.Caption = Format(DateTime.Now, "dddddd")
    Me.LblTasaCambio.Caption = TasaCambioR
End If

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


MDIPrimero.DtaConsulta.RecordSource = "SELECT  Nomina.NumNomina, SUM(DetalleNomina.INSSPatronal) AS INSSPatronal, SUM(DetalleNomina.IRPatronal) AS IRPatronal, SUM((Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30)) * 0.02) As INATEC FROM Nomina INNER JOIN Grupo INNER JOIN Cargo INNER JOIN TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN  DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON  TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado  GROUP BY Nomina.NumNomina  Having (Nomina.NumNomina = " & NumeroNomina & ")"
MDIPrimero.DtaConsulta.Refresh
If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
   Me.Field52.Text = Format(MDIPrimero.DtaConsulta.Recordset("INSSPatronal"), "##,##0.00")
   Me.Field151.Text = Format(MDIPrimero.DtaConsulta.Recordset("INATEC") * TasaCambioR, "##,##0.00")

End If


End Sub
