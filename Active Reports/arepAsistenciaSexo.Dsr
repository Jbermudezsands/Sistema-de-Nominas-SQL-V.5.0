VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arepAsistenciaSexo 
   Caption         =   "Asistencia x Sexo"
   ClientHeight    =   11115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arepAsistenciaSexo.dsx":0000
End
Attribute VB_Name = "arepAsistenciaSexo"
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

Private Sub GroupFooter1_AfterPrint()
'Me.lblTotalHorasEmpl.Caption = sngHorasTrabajadas
'Me.sngHorasTrabajadas = 0
End Sub

Private Sub GroupHeader1_Format()
If Me.txtSexo.Text = "M" Or Me.txtSexo.Text = "Masculino" Then
   Me.lblSexo.Caption = "Masculino"
ElseIf Me.txtSexo.Text = "F" Or Me.txtSexo.Text = "Femenino" Then
   Me.lblSexo.Caption = "Femenino"
End If

End Sub
