VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form FrmPersecciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Persecciones/Deducciones  Fijas de Nominas"
   ClientHeight    =   4485
   ClientLeft      =   15
   ClientTop       =   300
   ClientWidth     =   7005
   Icon            =   "FrmPesecciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7005
   Begin VB.Frame Frame2 
      Caption         =   "Persepciones de los Empleados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   6495
      Begin VB.TextBox Text5 
         Height          =   735
         Left            =   4680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Text            =   "FrmPesecciones.frx":0442
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4680
         TabIndex        =   20
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4680
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FrmPesecciones.frx":045D
         Left            =   4680
         List            =   "FrmPesecciones.frx":0467
         TabIndex        =   18
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1080
         TabIndex        =   17
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1080
         TabIndex        =   16
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Comentario"
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Cuenta de Pasivo"
         Height          =   255
         Left            =   3240
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Cuenta Costo"
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo de Moneda"
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Categoria"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Importe"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cód. Persep."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Deducciones de los Empleados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   22
      Top             =   480
      Width           =   6495
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   1080
         TabIndex        =   30
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text10 
         Height          =   735
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         Height          =   315
         Left            =   1080
         TabIndex        =   28
         Top             =   1920
         Width           =   1695
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1080
         TabIndex        =   27
         Top             =   2400
         Width           =   1695
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   4680
         TabIndex        =   26
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   4680
         TabIndex        =   25
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   4680
         TabIndex        =   24
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   735
         Left            =   4680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Text            =   "FrmPesecciones.frx":047E
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Cód. Tabla"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Importe"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Categoria"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo de Moneda"
         Height          =   255
         Left            =   3240
         TabIndex        =   34
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Cuenta Costo"
         Height          =   255
         Left            =   3240
         TabIndex        =   33
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Cuenta de Pasivo"
         Height          =   255
         Left            =   3240
         TabIndex        =   32
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Comentario"
         Height          =   255
         Left            =   3240
         TabIndex        =   31
         Top             =   1920
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Botones de Comando"
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   3615
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "Borrar"
         Height          =   375
         Left            =   1320
         MouseIcon       =   "FrmPesecciones.frx":049D
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   2400
         MouseIcon       =   "FrmPesecciones.frx":07A7
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   240
         MouseIcon       =   "FrmPesecciones.frx":0AB1
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7858
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Persepciones"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Deducciones"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmPersecciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub TabStrip1_Click()
 Select Case TabStrip1.SelectedItem.Index
    Case Is = 1           ' Abre archivo.
        Frame2.Visible = True
        Frame3.Visible = False
   
    Case Is = 2
       Frame2.Visible = False
       Frame3.Visible = True
      
        
    End Select
End Sub
