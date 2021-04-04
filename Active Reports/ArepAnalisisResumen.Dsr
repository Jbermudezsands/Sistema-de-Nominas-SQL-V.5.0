VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepAnalisisResumen 
   Caption         =   "ArepAnalisisResumen"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "ArepAnalisisResumen.dsx":0000
End
Attribute VB_Name = "ArepAnalisisResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TBasico As Double, TSH As Double, THB As Double
Private Sub ActiveReport_ReportStart()
     Me.LblTitulo3.Caption = "Reporte de Produccion de la Nomina No " & FrmReportes.TxtNNomina.Text
     Me.DataControl1.ConnectionString = ConexionReporte
     Me.LblTitulo.Caption = Titulo
     Me.LblSubtitulo.Caption = SubTitulo
     Me.ImgLogo.Picture = LoadPicture(RutaLogo)
End Sub

Private Sub Detail_Format()

 Dim NumNomina As Double, CodEmpleado As Double, SalarioHora As Double, HorasTrab As Double
 Dim TotalBasico As Double, Produccion As Double, HorasExtras As Double, Total As Double
 
 NumNomina = Me.FldNomina.Text
 CodEmpleado = Me.FldCodigoEmpleado.Text

 FrmReportes.AdoConsulta.RecordSource = "SELECT  * From DetalleHorasProduccion Where (NumNomina = " & NumNomina & ") And (CodEmpleado = " & CodEmpleado & ")"
 FrmReportes.AdoConsulta.Refresh
 If Not FrmReportes.AdoConsulta.Recordset.EOF Then
''   Select Case Me.FldDia.Text
'     Case "Lunes"
'           Me.LblHB.Caption = Format(FrmReportes.AdoConsulta.Recordset("Lunes"), "##,##0.00")
'           HorasTrab = FrmReportes.AdoConsulta.Recordset("Lunes")
'     Case "Martes"
'           Me.LblHB.Caption = Format(FrmReportes.AdoConsulta.Recordset("Martes"), "##,##0.00")
'           HorasTrab = FrmReportes.AdoConsulta.Recordset("Martes")
'     Case "Miercoles"
'           Me.LblHB.Caption = Format(FrmReportes.AdoConsulta.Recordset("Miercoles"), "##,##0.00")
'           HorasTrab = FrmReportes.AdoConsulta.Recordset("Miercoles")
'     Case "Jueves"
'           Me.LblHB.Caption = Format(FrmReportes.AdoConsulta.Recordset("Jueves"), "##,##0.00")
'           HorasTrab = FrmReportes.AdoConsulta.Recordset("Jueves")
'     Case "Viernes"
'           Me.LblHB.Caption = Format(FrmReportes.AdoConsulta.Recordset("Viernes"), "##,##0.00")
'           HorasTrab = FrmReportes.AdoConsulta.Recordset("Viernes")
'     Case "Sabado"
'           Me.LblHB.Caption = Format(FrmReportes.AdoConsulta.Recordset("Sabado"), "##,##0.00")
'           HorasTrab = FrmReportes.AdoConsulta.Recordset("Sabado")
'     Case "Domingo"
'           Me.LblHB.Caption = Format(FrmReportes.AdoConsulta.Recordset("Domingo"), "##,##0.00")
'           HorasTrab = FrmReportes.AdoConsulta.Recordset("Domingo")
'   End Select
   HorasExtras = Me.FldExtras.Text
   HorasTrab = FrmReportes.AdoConsulta.Recordset("Lunes") + FrmReportes.AdoConsulta.Recordset("Martes") + FrmReportes.AdoConsulta.Recordset("Miercoles") + FrmReportes.AdoConsulta.Recordset("Jueves") + FrmReportes.AdoConsulta.Recordset("Viernes") + FrmReportes.AdoConsulta.Recordset("Sabado") + FrmReportes.AdoConsulta.Recordset("Domingo")
   Me.LblHB.Caption = FrmReportes.AdoConsulta.Recordset("Lunes") + FrmReportes.AdoConsulta.Recordset("Martes") + FrmReportes.AdoConsulta.Recordset("Miercoles") + FrmReportes.AdoConsulta.Recordset("Jueves") + FrmReportes.AdoConsulta.Recordset("Viernes") + FrmReportes.AdoConsulta.Recordset("Sabado") + FrmReportes.AdoConsulta.Recordset("Domingo")
   Me.LblSH.Caption = Format(FrmReportes.AdoConsulta.Recordset("SalarioHora"), "##,##0.00")
   SalarioHora = Format(FrmReportes.AdoConsulta.Recordset("SalarioHora"), "##,##0.0000")
   TotalBasico = HorasTrab * SalarioHora
   Total = TotalBasico + HorasExtras
   Me.LblTotalB.Caption = Format(TotalBasico, "##,##0.00")
   Me.LblTotal.Caption = Format(Total, "##,##0.00")
   If Me.FldProduccion.Text <> "" Then
    Produccion = Me.FldProduccion.Text
   Else
    Produccion = 0
   End If
   Me.LblVariable.Caption = Format(Produccion - Total, "##,##0.00")
   If TotalBasico <> 0 Then
     Me.LblPorciento.Caption = ((Produccion - Total) / Total) * 100
     Me.LblPorciento.Caption = Format(Me.LblPorciento.Caption, "##,##0.00") & "%"
   Else
     Me.LblPorciento.Caption = "0%"
   End If
   
 Else
   Me.LblHB.Caption = "0.00"
   Me.LblSH.Caption = "0.00"
   
 End If
End Sub

