VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARDenominaciones 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19368
   SectionData     =   "ARDenominaciones.dsx":0000
End
Attribute VB_Name = "ARDenominaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_ReportStart()
LblTitulo.Caption = Titulo
LblSubtitulo.Caption = SubTitulo
If Dir(RutaLogo) <> "" Then
ImgLogo.Picture = LoadPicture(RutaLogo)
End If

'coloco los datos en las labels
'Txt500.Caption = FrmMonedas.Txt500.Text
Select Case Quien

Case "Nomina13vo"
    Txt200.Caption = FrmMonedas13vo.Txt200.Text
    Txt500.Caption = FrmMonedas13vo.Txt500.Text
    Txt1000.Caption = FrmMonedas13vo.Txt1000.Text
    Txt100.Caption = FrmMonedas13vo.Txt100.Text
    Txt50.Caption = FrmMonedas13vo.Txt50.Text
    Txt20.Caption = FrmMonedas13vo.Txt20.Text
    Txt10.Caption = FrmMonedas13vo.Txt10.Text
    Txt5.Caption = FrmMonedas13vo.Txt5.Text
    Txt1.Caption = FrmMonedas13vo.Txt1.Text

    TxtD50.Caption = FrmMonedas13vo.TxtD50.Text
    TxtD25.Caption = FrmMonedas13vo.TxtD25.Text
    TxtD10.Caption = FrmMonedas13vo.TxtD10.Text
    TxtD05.Caption = FrmMonedas13vo.TxtD05.Text
    TxtD01.Caption = FrmMonedas13vo.TxtD01.Text

    TxtTot1000.Caption = FrmMonedas13vo.TxtTot1000.Text
    TxtTot200.Caption = FrmMonedas13vo.TxtTot200.Text
    TxtTot500.Caption = FrmMonedas13vo.TxtTot500.Text
    TxtTot100.Caption = FrmMonedas13vo.TxtTot100.Text
    TxtTot50.Caption = FrmMonedas13vo.TxtTot50.Text
    TxtTot20.Caption = FrmMonedas13vo.TxtTot20.Text
    TxtTot10.Caption = FrmMonedas13vo.TxtTot10.Text
    TxtTot5.Caption = FrmMonedas13vo.TxtTot5.Text
    TxtTot1.Caption = FrmMonedas13vo.TxtTot1.Text

    TxtTotD50.Caption = FrmMonedas13vo.TxtTotD50.Text
    TxtTotD25.Caption = FrmMonedas13vo.TxtTotD25.Text
    TxtTotD10.Caption = FrmMonedas13vo.TxtTotD10.Text
    TxtTotD05.Caption = FrmMonedas13vo.TxtTotD05.Text
    TxtTotD01.Caption = FrmMonedas13vo.TxtTotD01.Text

    TxtGranTotal.Caption = FrmMonedas13vo.TxtGranTotal.Text
    LblSubtitulo2.Caption = FrmMonedas13vo.Label1.Caption



Case Else
    Me.Txt200.Caption = FrmMonedas.Txt200.Text
    Me.Txt500.Caption = FrmMonedas.Txt500.Text
    Me.Txt1000.Caption = FrmMonedas.Txt1000.Text
    Txt100.Caption = FrmMonedas.Txt100.Text
    Txt50.Caption = FrmMonedas.Txt50.Text
    Txt20.Caption = FrmMonedas.Txt20.Text
    Txt10.Caption = FrmMonedas.Txt10.Text
    Txt5.Caption = FrmMonedas.Txt5.Text
    Txt1.Caption = FrmMonedas.Txt1.Text

    TxtD50.Caption = FrmMonedas.TxtD50.Text
    TxtD25.Caption = FrmMonedas.TxtD25.Text
    TxtD10.Caption = FrmMonedas.TxtD10.Text
    TxtD05.Caption = FrmMonedas.TxtD05.Text
    TxtD01.Caption = FrmMonedas.TxtD01.Text
    
    TxtTot200.Caption = FrmMonedas.TxtTot200.Text
    TxtTot500.Caption = FrmMonedas.TxtTot500.Text
    TxtTot1000.Caption = FrmMonedas.TxtTot1000.Text
    TxtTot100.Caption = FrmMonedas.TxtTot100.Text
    TxtTot50.Caption = FrmMonedas.TxtTot50.Text
    TxtTot20.Caption = FrmMonedas.TxtTot20.Text
    TxtTot10.Caption = FrmMonedas.TxtTot10.Text
    TxtTot5.Caption = FrmMonedas.TxtTot5.Text
    TxtTot1.Caption = FrmMonedas.TxtTot1.Text

    TxtTotD50.Caption = FrmMonedas.TxtTotD50.Text
    TxtTotD25.Caption = FrmMonedas.TxtTotD25.Text
    TxtTotD10.Caption = FrmMonedas.TxtTotD10.Text
    TxtTotD05.Caption = FrmMonedas.TxtTotD05.Text
    TxtTotD01.Caption = FrmMonedas.TxtTotD01.Text

    TxtGranTotal.Caption = FrmMonedas.TxtGranTotal.Text
    LblSubtitulo2.Caption = FrmMonedas.Label1.Caption
End Select
End Sub

