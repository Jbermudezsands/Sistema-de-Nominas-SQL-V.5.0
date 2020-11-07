VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepBalance 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ArepBalance.dsx":0000
End
Attribute VB_Name = "ArepBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()

 If Me.FLdKeyGrupo.Text = "A" Or Me.FLdKeyGrupo.Text = "PC" Then
  If Val(Me.Field3.Text) <> 0 Then
   Me.Line1.Visible = True
   Me.Line2.Visible = True
  Else
   Me.Line1.Visible = False
   Me.Line2.Visible = False
  End If
 Else
  Me.Line1.Visible = False
  Me.Line2.Visible = False
 End If
 
 If Me.FldNivel.Text = "2" Then
  If Val(Me.Field3.Text) <> 0 Then
    Me.Line5.Visible = True
  Else
    Me.Line5.Visible = False
  End If
 Else
  Me.Line5.Visible = False
 End If

End Sub
