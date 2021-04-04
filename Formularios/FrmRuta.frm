VERSION 5.00
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Begin VB.Form FrmRuta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ruta"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   HelpContextID   =   40
   Icon            =   "FrmRuta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3105
   Begin vbskfree.Skinner Skinner1 
      Left            =   240
      Top             =   3360
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.DirListBox Dir1 
      DragIcon        =   "FrmRuta.frx":030A
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      DragIcon        =   "FrmRuta.frx":0614
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton CmdPegar 
      DownPicture     =   "FrmRuta.frx":091E
      Height          =   375
      Left            =   120
      MouseIcon       =   "FrmRuta.frx":2400
      MousePointer    =   99  'Custom
      Picture         =   "FrmRuta.frx":2842
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      DownPicture     =   "FrmRuta.frx":4324
      Height          =   375
      Left            =   1560
      MouseIcon       =   "FrmRuta.frx":5E06
      MousePointer    =   99  'Custom
      Picture         =   "FrmRuta.frx":6248
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
   End
End
Attribute VB_Name = "FrmRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdPegar_Click()

If QuienLlama = "Exporta Nomina" Then
    If Mid$(Dir1.Path, Len(Dir1.Path)) = "\" Then
      If FrmExporta.OptCuentas.Value = True Then
        Origen = FrmRuta.Dir1.Path & "Zeus.Txt"
      Else
       Origen = FrmRuta.Dir1.Path & "Zeus.Cns"
     End If
    Else
      If FrmExporta.OptCuentas.Value = True Then
       Origen = FrmRuta.Dir1.Path & "\" & "Zeus.Txt"
      Else
       Origen = FrmRuta.Dir1.Path & "\" & "Zeus.Cns"
      End If
    End If
    
     FrmExporta.TxtRuta.Text = Origen
     Unload Me
End If

If QuienLlama = "Exporta Préstamo" Then
    If Mid$(Dir1.Path, Len(Dir1.Path)) = "\" Then
       Origen = FrmRuta.Dir1.Path & "Prestamo.Cns"
    Else
       Origen = FrmRuta.Dir1.Path & "\" & "Prestamo.Cns"
    End If
     frmEmpleado.TxtRuta.Text = Origen
     Unload Me
End If
 End Sub

Private Sub CmdSalir_Click()
 Unload Me
End Sub

Private Sub Dir1_Change()
'File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo DriveErrs
    Dir1.Path = Drive1.Drive
Exit Sub
DriveErrs:
    Select Case Err
        Case 68
            MsgBox Prompt:="La unidad no está preparada. Inserte un disco en la unidad.", Buttons:=vbExclamation, Title:="Sistema de Nominas"
            ' Restablece la ruta a la unidad anterior.
            Drive1.Drive = Dir1.Path
            Exit Sub
        Case Else
            MsgBox Prompt:="Error en la aplicación.", Buttons:=vbExclamation
    End Select

End Sub

