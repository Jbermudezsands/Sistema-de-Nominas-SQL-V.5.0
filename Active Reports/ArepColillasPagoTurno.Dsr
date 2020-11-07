VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepColillasPagoTurno 
   Caption         =   "Colillas de Pago"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "ArepColillasPagoTurno.dsx":0000
End
Attribute VB_Name = "ArepColillasPagoTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
 Dim Destino As String

        If Dir(RutaFoto & FrmBajas.txtCodEmpleado1.Text & ".jpg") <> "" Then
           Destino = RutaFoto & FrmBajas.txtCodEmpleado1.Text & ".jpg"
        ElseIf Dir(RutaFoto & FrmBajas.txtCodEmpleado1.Text & ".gif") <> "" Then
           Destino = RutaFoto & FrmBajas.txtCodEmpleado1.Text & ".gif"
        ElseIf Dir(RutaFoto & FrmBajas.txtCodEmpleado1.Text & ".bmp") <> "" Then
           Destino = RutaFoto & FrmBajas.txtCodEmpleado1.Text & ".bmp"
        End If
        
        If (Dir(Destino) <> "") Then
         Me.ImgLogo.Picture = LoadPicture(Destino)
        Else
          Destino = App.Path + "\Zw.bmp"
'          Destino = RutaLogo
         Me.ImgLogo.Picture = LoadPicture(Destino)
        End If
        
        RutaLogo = MDIPrimero.DtaEmpresa.Recordset("RutaLogo")
        If (Dir(RutaLogo) <> "") Then
         Me.ImgLogo.Picture = LoadPicture(RutaLogo)
        Else
          RutaLogo = App.Path + "\Zw.bmp"
          Me.ImgLogo.Picture = LoadPicture(RutaLogo)
        End If
End Sub

Private Sub Detail_Format()

'Me.LblTotalDeduccion.Caption = Format(CDbl(Me.FldAdelanto.Text) + CDbl(Me.FldDeducciones.Text) + CDbl(Me.FldDescuento.Text), "#,##0.00")
Me.LblTotalDeduccion.Caption = Format(CDbl(Me.FldDeducciones.Text) - CDbl(Me.FldDescuento.Text) - CDbl(Me.FldAdelanto.Text), "#,##0.00")

End Sub

