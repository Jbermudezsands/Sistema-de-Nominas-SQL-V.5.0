VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepHistorialSalarial13vo 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepHistorialSalarial13vo.dsx":0000
End
Attribute VB_Name = "ArepHistorialSalarial13vo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_ReportStart()
 Me.LblTitulo.Caption = Titulo
 Me.LblSubtitulo.Caption = SubTitulo
 Me.LblTipoNomina.Caption = "Desde " & FrmSalarioHistorial.TxtFINI13.Value & " Hasta " & FrmSalarioHistorial.TxtFFIN13.Value
 
 Me.LblFechaHoy.Caption = Now
 
 
 If Quien = "Vacaciones" Then
    Me.LblTipoNomina.Caption = "Desde " & FrmSalarioHistorial.TxtFINIVaca.Value & " Hasta " & FrmSalarioHistorial.TxtFFinVaca.Value
     If Month(FrmSalarioHistorial.TxtFINIVaca.Value) = 1 Then
'          Me.txtjunio.Visible = False
'          Me.txtjulio.Visible = False
'          Me.txtAgosto.Visible = False
'          Me.txtSeptiembre.Visible = False
'          Me.txtOctubre.Visible = False
'          Me.txtNoviembre.Visible = False

          Me.txtjunio.DataField = "Enero"
          Me.txtjulio.DataField = "Febrero"
          Me.txtAgosto.DataField = "Marzo"
          Me.txtSeptiembre.DataField = "Abril"
          Me.txtOctubre.DataField = "Mayo"
          Me.txtNoviembre.DataField = "Junio"
         
          Me.LblJunio.Caption = "Enero"
          Me.LblJulio.Caption = "Febrero"
          Me.LblAgosto.Caption = "Marzo"
          Me.LblSeptiembre.Caption = "Abril"
          Me.LblOctubre.Caption = "Mayo"
          Me.LblNoviembre.Caption = "Junio"
          
          Me.LblJunio.Caption = "Enero"
          Me.LblDiciembre.Visible = False
          Me.txtDiciembre.Visible = False
          
          
    End If
    
 End If
    
    
End Sub

