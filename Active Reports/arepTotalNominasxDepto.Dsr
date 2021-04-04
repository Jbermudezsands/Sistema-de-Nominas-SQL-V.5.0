VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arepTotalNominasxDepto 
   Caption         =   "Asistencia"
   ClientHeight    =   11490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "arepTotalNominasxDepto.dsx":0000
End
Attribute VB_Name = "arepTotalNominasxDepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sngHorasTrabajadas As Single
Public iGeneroMasculino As Integer
Public iGeneroFemenino As Integer


Private Sub Detail_BeforePrint()


Dim lDiaSiguiente As Double


Me.txtNombre.Text = Me.txtNombre1.Text & " " & Me.txtNombre2.Text & " " & Me.txtApellido1.Text & " " & Me.txtApellido2.Text

If Me.txtGenero.Text = "Masculino" Or Me.txtGenero.Text = "M" Then
   iGeneroMasculino = iGeneroMasculino + 1
Else
   iGeneroFemenino = iGeneroFemenino + 1

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

Me.txtNombre.Text = Me.txtNombre1.Text & " " & Me.txtNombre2.Text & " " & Me.txtApellido1.Text & " " & Me.txtApellido2.Text

If Me.txtGenero.Text = "Masculino" Or Me.txtGenero.Text = "M" Then
   iGeneroMasculino = iGeneroMasculino + 1
Else
   iGeneroFemenino = iGeneroFemenino + 1

End If




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

 
 Me.txtMasculino.Text = iGeneroMasculino
 Me.txtFemenino.Text = CInt(Me.txtConteoFinal.Text) - iGeneroMasculino
 Me.txtTGeneralMasculino.Text = CInt(Me.txtTGeneralMasculino.Text) + iGeneroMasculino
 Me.txtTGeneralFemenino.Text = CInt(Me.txtTGeneralFemenino.Text) + CInt(Me.txtFemenino.Text)
 
 iGeneroMasculino = 0
 iGeneroFemenino = 0
 




End Sub

Private Sub PageFooter_BeforePrint()

If bExportar = True Then

    Dim myExportObject As ActiveReportsExcelExport.ARExportExcel
    Dim Nombre As String

'    Nombre = InputBox("Digite el Nombre del Archivo", "Sistema de Nominas")
    Set myExportObject = CreateObject("ActiveReportsExcelExport.ARExportExcel")
    myExportObject.FileName = Directorio
    myExportObject.Export Me.Pages
    Set myExportObject = Nothing
    bExportar = False

End If



End Sub



