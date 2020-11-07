VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arepHLaboradas 
   Caption         =   "Asistencia"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "arepHLaboradas.dsx":0000
End
Attribute VB_Name = "arepHLaboradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sngHorasTrabajadas As Single


Private Sub Detail_BeforePrint()

Dim lDiaAnterior As Double
Dim lDiaSiguiente As Double


Me.txtNombre.Text = Me.txtNombre1.Text & " " & Me.txtNombre2.Text & " " & Me.txtApellido1.Text & " " & Me.txtApellido2.Text

If IsNumeric(Me.txtHExtras.Text) Then
   Me.lblValor.Caption = Format(2 * CSng(Me.txtHExtras.Text) * CSng(Me.txtSalHora.Text), "##,###.##")
Else
   Me.lblValor.Caption = "0"
End If

If IsNumeric(Me.txtTotalSalHora.Text) And Me.txtTotalSalHora.Text <> "0" Then
   Me.txtTotalSalHora.Text = Format(Me.txtTotalSalHora.Text, "##, ###.##")
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

'If Me.txtDia.Text = "Lun" Then
'   Me.lblDia.Caption = "Lunes"
'ElseIf Me.txtDia.Text = "Mart" Then
'   Me.lblDia.Caption = "Martes"
'ElseIf Me.txtDia.Text = "Mierc" Then
'   Me.lblDia.Caption = "Miercoles"
'ElseIf Me.txtDia.Text = "Juev" Then
'   Me.lblDia.Caption = "Jueves"
'ElseIf Me.txtDia.Text = "Viern" Then
'   Me.lblDia.Caption = "Viernes"
'ElseIf Me.txtDia.Text = "Sab" Then
'   Me.lblDia.Caption = "Sábado"
'ElseIf Me.txtDia.Text = "Dom" Then
'   Me.lblDia.Caption = "Domingo"
'End If


End Sub

Private Sub GroupFooter1_AfterPrint()
'Me.lblTotalHorasEmpl.Caption = sngHorasTrabajadas
'Me.sngHorasTrabajadas = 0

End Sub

Private Sub GroupFooter1_BeforePrint()
'Me.txtTHLaboradas.Text = Format(Me.txtTHLaboradas.Text, "##.##")
'Me.txtTHExtras.Text = Format(Me.txtTHExtras.Text, "##.##")
End Sub

