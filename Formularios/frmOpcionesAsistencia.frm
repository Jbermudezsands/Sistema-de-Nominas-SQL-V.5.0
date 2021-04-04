VERSION 5.00
Begin VB.Form frmOpciones 
   Caption         =   "Opciones de Asistencia"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
      Begin VB.OptionButton optExtraTurno 
         Caption         =   "Personal Laborando Turno Extra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   3000
         Width           =   3855
      End
      Begin VB.OptionButton optHorasLaboradas 
         Caption         =   "Calculo de Horas Laboradas y Reportes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   3
         Top             =   2160
         Width           =   4695
      End
      Begin VB.OptionButton optPermisos 
         Caption         =   "Ingreso/Modificación de Permisos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   2
         Top             =   1200
         Width           =   4215
      End
      Begin VB.OptionButton optAsistenciaManual 
         Caption         =   "Asistencia Manual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()

  If Me.optAsistenciaManual.Value Then
'     frmAsistManual.Show 1
  ElseIf Me.optPermisos.Value Then
'     frmPermisos.Show 1
  ElseIf Me.optExtraTurno.Value Then
'     frmExtraTurno.Show
  Else
'     frmRepAsistencia.Show
  End If
  


End Sub

Private Sub CmdSalir_Click()

Unload Me

End Sub
