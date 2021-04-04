VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepColillasPago4 
   Caption         =   "Colillas de Pago"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   16642
   SectionData     =   "ArepColillasPago4.dsx":0000
End
Attribute VB_Name = "ArepColillasPago4"
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
          If Dir(RutaLogo) <> "" Then
          Me.ImgLogo.Picture = LoadPicture(RutaLogo)
          End If
        End If
                
End Sub

Private Sub Detail_Format()
 Dim CodigoEmpleado As String, NumeroNomina As Double, FechaNomina As Date
 
 CodigoEmpleado = Me.Field14.Text
 NumeroNomina = Me.Field17.Text
 Me.LblCanasta.Caption = Format(MontoDeduccion(NumeroNomina, "03", CodigoEmpleado), "##,##0.00")
 Me.LblAlimentos.Caption = Format(MontoDeduccion(NumeroNomina, "04", CodigoEmpleado), "##,##0.00")
 Me.LbllSindicatos.Caption = Format(MontoDeduccion(NumeroNomina, "05", CodigoEmpleado), "##,##0.00")
 
'Me.LblTotalDeduccion.Caption = Format(CDbl(Me.FldAdelanto.Text) + CDbl(Me.FldDeducciones.Text) + CDbl(Me.FldDescuento.Text), "#,##0.00")
Me.LblTotalDeduccion.Caption = Format(CDbl(Me.FldDeducciones.Text) - CDbl(Me.FldDescuento.Text) - CDbl(Me.FldAdelanto.Text), "#,##0.00")
Me.Field20.Text = Format(CDbl(Me.FldDeducciones.Text) - CDbl(Me.FldDescuento.Text) - CDbl(Me.FldAdelanto.Text) - CDbl(Me.LblCanasta.Caption) - CDbl(Me.LblAlimentos.Caption) - CDbl(Me.LbllSindicatos.Caption), "#,##0.00")

Dim sql As String
               
       sql = "SELECT DetalleDeduccion.Valor, DetalleDeduccion.NumNomina, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, TipoDeduccion.Deduccion FROM DetalleDeduccion INNER JOIN Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion INNER JOIN Empleado ON Deduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion WHERE        (DetalleDeduccion.NumNomina = " & Me.Field17.Text & ") AND (Empleado.CodEmpleado1 ='" & Field14.Text & "')"
       Set Me.SubReport1.object = New SubColillaDeduccion
       Me.SubReport1.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReport1.object.DataControl1.Source = sql

End Sub

