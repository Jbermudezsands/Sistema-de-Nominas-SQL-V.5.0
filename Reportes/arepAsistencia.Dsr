VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arepAsistenciaReal 
   Caption         =   "Asistencia"
   ClientHeight    =   10860
   ClientLeft      =   165
   ClientTop       =   705
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19156
   SectionData     =   "arepAsistencia.dsx":0000
End
Attribute VB_Name = "arepAsistenciaReal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sngHorasTrabajadas As Single


Private Sub ActiveReport_ReportStart()
  Me.LblTitulo.Caption = Titulo
End Sub

Private Sub Detail_BeforePrint()

Dim lDiaAnterior As Double
Dim lDiaSiguiente As Double


Me.TxtNombre.Text = Me.TxtNombre1.Text & " " & Me.TxtNombre2.Text & " " & Me.TxtApellido1.Text & " " & Me.TxtApellido2.Text

If Me.txtHLaboradas.Text <> "0" Then
   Me.txtHLaboradas.Text = Format(Me.txtHLaboradas.Text, "##.####")
End If

If Me.txtHExtras.Text <> "0" Then
   Me.txtHExtras.Text = Format(Me.txtHExtras.Text, "##.##")
End If


'If Me.txtFechaEntrada.Text = Me.txtFechaSalida.Text Then
'
'     Me.lblTotalEmpl.Caption = Format((DateDiff("n", Me.txtHEntrada.Text, Me.txtHSalida.Text) / 60) - (Me.txtTComida.Text / 60), "##.##")
'Else
'     lDiaAnterior = DateDiff("n", Me.txtHEntrada.Text, "23:59:59") / 60
'     lDiaSiguiente = (lDiaAnterior + DateDiff("n", "00:00:00", Me.txtHSalida.Text) / 60) - (Me.txtTComida.Text / 60)
'     Me.lblTotalEmpl.Caption = Format(lDiaSiguiente, "##.##")
'
'End If
'
'Me.sngHorasTrabajadas = CSng(Me.lblTotalEmpl.Caption)

     
End Sub

Private Sub Detail_Format()

If Me.txtDia.Text = "Lun" Then
   Me.lblDia.Caption = "Lunes"
ElseIf Me.txtDia.Text = "Mart" Then
   Me.lblDia.Caption = "Martes"
ElseIf Me.txtDia.Text = "Mierc" Then
   Me.lblDia.Caption = "Miercoles"
ElseIf Me.txtDia.Text = "Juev" Then
   Me.lblDia.Caption = "Jueves"
ElseIf Me.txtDia.Text = "Viern" Then
   Me.lblDia.Caption = "Viernes"
ElseIf Me.txtDia.Text = "Sab" Then
   Me.lblDia.Caption = "Sábado"
ElseIf Me.txtDia.Text = "Dom" Then
   Me.lblDia.Caption = "Domingo"
End If


End Sub

Private Sub GroupFooter1_AfterPrint()
'Me.lblTotalHorasEmpl.Caption = sngHorasTrabajadas
'Me.sngHorasTrabajadas = 0



End Sub

Private Sub GroupFooter1_BeforePrint()

Dim alto As String

If Me.TxtCodEmpleado.Text = "000095" Then
   alto = "0"
End If


Me.txtTHLaboradas.Text = Format(Me.txtTHLaboradas.Text, "##.##")
Me.txtTHExtras.Text = Format(Me.txtTHExtras.Text, "##.##")
If Me.txtNomina.Text = "Administracion" And IsNumeric(Me.txtTHLaboradas.Text) Then
   Me.lblFaltas.Caption = Abs(CSng(Me.txtTHLaboradas.Text) - 96)
ElseIf IsNumeric(Me.txtTHLaboradas.Text) Then
   If 48 - CSng(Me.txtTHLaboradas.Text) < 1 Then
        Me.lblFaltas.Caption = Format(CStr(48 - CSng(Me.txtTHLaboradas.Text)), "#.####")
   Else
        Me.lblFaltas.Caption = CStr(48 - CSng(Me.txtTHLaboradas.Text))
   End If
   
End If



End Sub

