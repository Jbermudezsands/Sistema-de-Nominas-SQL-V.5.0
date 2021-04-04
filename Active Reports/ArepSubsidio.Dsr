VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepNomSubsidio 
   Caption         =   "Nomina de Subsidios"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepSubsidio.dsx":0000
End
Attribute VB_Name = "ArepNomSubsidio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GroupFooter1_Format()

Monto = Me.Field12.Text
Total = Monto + Total
Me.LblTotal.Caption = Format(Total, "##,##0.00")
End Sub

Private Sub GroupHeader1_Format()
 Dim Nombre As String
Nombre = Me.Field2.Text + " " + Me.Field3 + " " + Me.Field4 + " " + Me.Field5
Me.LblNombre = Nombre
End Sub
