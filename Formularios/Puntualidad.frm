VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " "
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.Data dtaDevengado 
      Caption         =   "Devengado Hora"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data dtaIngreso 
      Caption         =   "Ingreso"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton cmdReestablecer 
      Caption         =   "Ingresar Puntualidad"
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Una vez que se le de click en el boton, esperar que termine de procesar..... luego recalcular nomina y verificar en reportes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   3000
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReestablecer_Click()


Do While Not Me.dtaDevengado.Recordset.EOF

    Me.dtaIngreso.RecordSource = "SELECT * FROM Ingreso_Empl WHERE Cod_Ing = '09' AND Cod_Empl =" & Me.dtaDevengado.Recordset.Fields("Cod_Empl") & " AND Periodo =33 AND mes ='Agosto' AND año =2005"
    Me.dtaIngreso.Refresh
    
    If Not Me.dtaIngreso.Recordset.EOF Then
       Me.dtaDevengado.Recordset.Edit
       Me.dtaDevengado.Recordset.Fields("IncPunt") = Me.dtaIngreso.Recordset.Fields("Ingreso")
       Me.dtaDevengado.Recordset.Update
    End If
    
    Me.dtaDevengado.Recordset.MoveNext

Loop

MsgBox "Proceso terminado"



End Sub

Private Sub Form_Load()

Me.dtaDevengado.DatabaseName = App.Path & "\PlanMetro.mdb"
Me.dtaDevengado.RecordSource = "SELECT * FROM Devengado_Hora WHERE Periodo =33 AND mes ='Agosto' AND año =2005 AND IncPunt =0 ORDER BY Cod_Empl ASC"
Me.dtaDevengado.Refresh

Me.dtaIngreso.DatabaseName = App.Path & "\PlanMetro.mdb"
Me.dtaIngreso.RecordSource = "Ingreso_Empl"
Me.dtaIngreso.Refresh



End Sub
