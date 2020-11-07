VERSION 5.00
Begin VB.Form FrmActivaPeriodos 
   Caption         =   "Activacion de Periodos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      Begin VB.Label LblPeriodo 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label LblFechaFinal 
         Height          =   255
         Left            =   5400
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
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
         Left            =   3960
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label LblFechaIni 
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmActivaPeriodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
