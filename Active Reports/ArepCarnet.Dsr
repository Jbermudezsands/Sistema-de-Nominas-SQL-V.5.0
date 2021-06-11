VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepCarnet 
   Caption         =   "Carnet del Empleado"
   ClientHeight    =   8730
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   20108
   _ExtentY        =   15399
   SectionData     =   "ArepCarnet.dsx":0000
End
Attribute VB_Name = "ArepCarnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
Dim CodigoEmpleado As Double

If Me.FldCodEmpleado.Text <> "" Then
  CodigoEmpleado = Me.FldCodEmpleado.Text
End If
Me.LblCodigo.Caption = Me.FldCodigoEmpleado.Text

 Dim Destino As String, RutaFirma As String
 
    If Dir(RutaFoto & Me.LblCodigo.Caption & ".jpg") <> "" Then
        Destino = RutaFoto & Me.LblCodigo.Caption & ".jpg"
    ElseIf Dir(RutaFoto & Me.LblCodigo.Caption & ".gif") <> "" Then
        Destino = RutaFoto & Me.LblCodigo.Caption & ".gif"
    ElseIf Dir(RutaFoto & Me.LblCodigo.Caption & ".bmp") <> "" Then
        Destino = RutaFoto & Me.LblCodigo.Caption & ".bmp"
    End If
    
    
   If Dir(RutaFoto & "firma.jpg") <> "" Then
        RutaFirma = RutaFoto & "firma.jpg"
    ElseIf Dir(RutaFoto & "firma.gif") <> "" Then
        RutaFirma = RutaFoto & "firma.gif"
    ElseIf Dir(RutaFoto & "firma.bmp") <> "" Then
        RutaFirma = RutaFoto & "firma.bmp"
    End If
    
    
    
    
           
    If (Dir(Destino) <> "") Then
        Me.ImgFoto.Picture = LoadPicture(Destino)
    Else
        Destino = App.Path + "\Zw.bmp"
        Me.ImgFoto.Picture = LoadPicture(Destino)
    End If
    
    If (Dir(RutaFirma) <> "") Then
        Me.ImgFirma.Picture = LoadPicture(RutaFirma)
    Else
        RutaFirma = App.Path + "\Zw.bmp"
        Me.ImgFirma.Picture = LoadPicture(RutaFirma)
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
    
    Me.LblTitulo.Caption = Titulo
    
    '////////////////CONSULTO LA FECHA DE INGRESO /////////////////////////////
    MDIPrimero.AdoConsulta.ConnectionString = Conexion
    MDIPrimero.AdoConsulta.RecordSource = "SELECT Empleado.CodEmpleado, Historico.FechaContrato, Empleado.CodEmpleado1 FROM  Empleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE (Empleado.CodEmpleado = " & CodigoEmpleado & ")"
    MDIPrimero.AdoConsulta.Refresh
    If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
      Me.LblFechaContrato.Caption = "Fecha Contrato: " & Format(MDIPrimero.AdoConsulta.Recordset("FechaContrato"), "dd/mm/yyyy")
    End If
    
    
    

End Sub
