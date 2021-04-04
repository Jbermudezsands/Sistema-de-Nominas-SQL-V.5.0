VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBackup 
   BorderStyle     =   0  'None
   Caption         =   "Crear Respaldo"
   ClientHeight    =   4725
   ClientLeft      =   -30
   ClientTop       =   -30
   ClientWidth     =   5040
   HelpContextID   =   39
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   315
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ShowInTaskbar   =   0   'False
   Begin Project1.xp_canvas xp_canvas1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8281
      Caption         =   "Crear Respaldo"
      Fixed_Single    =   -1  'True
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1815
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   2160
         Pattern         =   "*.Zn"
         TabIndex        =   9
         Top             =   1680
         Width           =   2775
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   2400
         TabIndex        =   8
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton CmdProcesar 
         DownPicture     =   "frmBackup.frx":030A
         Height          =   375
         Left            =   3480
         MouseIcon       =   "frmBackup.frx":1DEC
         MousePointer    =   99  'Custom
         Picture         =   "frmBackup.frx":20F6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton CmdCerrar 
         DownPicture     =   "frmBackup.frx":3BD8
         Height          =   375
         Left            =   3480
         MouseIcon       =   "frmBackup.frx":56BA
         MousePointer    =   99  'Custom
         Picture         =   "frmBackup.frx":59C4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2175
         Begin VB.OptionButton OptAbrir 
            Caption         =   "Abrir Respaldo"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton OptGurdar 
            Caption         =   "Crear  Respaldo"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   3600
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Crear Respaldo / Guardar Respaldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   0
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCerrar_Click()
Unload Me
End Sub

Private Sub CmdProcesar_Click()

 On Error GoTo TipoErrs
 If Mid$(frmBackup.File1.Path, Len(frmBackup.File1.Path)) = "\" Then
   Ruta = File1.Path & "Zeus.Zn"
 Else
   Ruta = File1.Path & "\" & "Zeus.Zn"
 End If
If frmBackup.OptGurdar.Value = True Then
 If (Dir(Ruta) <> "") Then
   Cadena = Cadena & "      Reescribir el Archivo?" & vbLf
   R% = MsgBox(Cadena, vbYesNo, "Sistema de Nominas")
    Cadena = ""
    If R% = 6 Then
      Kill (Ruta)
      Backad
      Unload Me
    Else
      Exit Sub
    End If
 Else
    Backad
    Unload Me
 End If
 Else
    If Mid$(File1.Path, Len(File1.Path)) = "\" Then
     Origen = File1.Path & File1.FileName
    Else
      Origen = File1.Path & "\" & File1.FileName
    End If
        
        Destino = Ruta
        If (Dir(Destino) <> "") And (Dir(Origen) <> "") Then
          Cadena = "                 Reescribir el Archivo?" & vbLf
          Cadena = Cadena & "Si reescribe Perdera la base de Datos Actual" & vbLf
          Cadena = Cadena & "        Respalde antes de Abrir un Respaldo" & vbLf
          R% = MsgBox(Cadena, vbYesNo, "Sistema de Nominas")
           Cadena = ""
          
           If R% = 6 Then
              frmBackup.CmdProcesar.Enabled = False
               frmBackup.CmdCerrar.Enabled = False
               FileCopy Origen, Destino
               Cadena = "El respaldo se ha Insertado" & vbLf
               Cadena = Cadena & "Correctamente.."
               R% = MsgBox(Cadena, vbExclamation, "Sistema de Nominas")
               Unload Me
           End If
        Else
         
        End If

End If
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
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
Private Sub Form_Load()

frmBackup.CmdCerrar.MousePointer = 99
frmBackup.CmdProcesar.MousePointer = 99
Label2.Caption = "Unidad a Guardar"
End Sub

Private Sub OptAbrir_Click()
Label2.Caption = "Unidad a Abrir"
End Sub

Private Sub OptGurdar_Click()
Label2.Caption = "Unidad a Guardar"
End Sub

