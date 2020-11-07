VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arepAsistenciaDepto 
   Caption         =   "Asistencia x Depto"
   ClientHeight    =   10860
   ClientLeft      =   165
   ClientTop       =   705
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19156
   SectionData     =   "arepAsistenciaDepto.dsx":0000
End
Attribute VB_Name = "arepAsistenciaDepto"
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

