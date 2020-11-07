VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} Nom13vo 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "Nom13vo.dsx":0000
End
Attribute VB_Name = "Nom13vo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dTotalPagar As Double


Private Sub ActiveReport_ReportEnd()
If Exportar = True Then
    Dim myExportObject As ActiveReportsExcelExport.ARExportExcel
    Dim Nombre As String
    
    Set myExportObject = CreateObject("ActiveReportsExcelExport.ARExportExcel")
    myExportObject.FileName = Directorio
    myExportObject.Export Me.Pages
    Set myExportObject = Nothing
 End If
End Sub

Private Sub Detail_BeforePrint()

If CDbl(Me.txtDiasAcum.Text) >= 30 Then

   Me.txtTotalPagar.Text = Me.txtMontoPagar.Text
   
   Me.dTotalPagar = Me.dTotalPagar + CDbl(Me.txtMontoPagar.Text)
   
Else

'   Me.txtTotalPagar.Text = Format((CDbl(Me.txtMontoPagar.Text) / 30.4167) * CDbl(Me.txtDiasAcum.Text), "##,###.#0")
'   Me.dTotalPagar = Me.dTotalPagar + Me.txtTotalPagar.Text
   
End If



End Sub

Private Sub GroupFooter1_BeforePrint()


Me.txtTGeneral.Text = Me.dTotalPagar



End Sub

