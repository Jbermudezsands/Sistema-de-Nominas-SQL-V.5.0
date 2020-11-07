VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepCarnet 
   Caption         =   "Carnet del Empleado"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepCarnet.dsx":0000
End
Attribute VB_Name = "ArepCarnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
Me.lblCodigo.Caption = Me.FldCodigoEmpleado.Text

 Dim Destino As String, RutaFirma As String
 
    If Dir(RutaFoto & Me.lblCodigo.Caption & ".jpg") <> "" Then
        Destino = RutaFoto & Me.lblCodigo.Caption & ".jpg"
    ElseIf Dir(RutaFoto & Me.lblCodigo.Caption & ".gif") <> "" Then
        Destino = RutaFoto & Me.lblCodigo.Caption & ".gif"
    ElseIf Dir(RutaFoto & Me.lblCodigo.Caption & ".bmp") <> "" Then
        Destino = RutaFoto & Me.lblCodigo.Caption & ".bmp"
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
    
    Me.lblTitulo.Caption = Titulo
    
    

End Sub
