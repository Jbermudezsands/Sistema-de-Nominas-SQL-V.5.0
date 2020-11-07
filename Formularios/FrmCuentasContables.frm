VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmCuentasContables 
   Caption         =   "Cuentas Contables"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   7800
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc AdoHistorico 
      Height          =   495
      Left            =   720
      Top             =   7560
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoHistorico"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox TxtCodEmpleado 
      Height          =   375
      Left            =   2160
      TabIndex        =   48
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   7815
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Empleado: 0001 Juan Bermudez"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   5160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   8880
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image2 
         Height          =   1020
         Left            =   0
         Picture         =   "FrmCuentasContables.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1290
      End
   End
   Begin XtremeSuiteControls.PushButton CmdGrabar 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5880
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Grabar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmCuentasContables.frx":0AFE
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   5880
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmCuentasContables.frx":2E62
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3615
      _Version        =   786432
      _ExtentX        =   6376
      _ExtentY        =   7858
      _StockProps     =   79
      Caption         =   "Cuentas Debito"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox TxtCuentaSubsidio 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   55
         Text            =   "11111"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtCtaHorasExtras 
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   11
         Text            =   "11111"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox TxtCtaINATEC 
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   10
         Text            =   "11111"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox TxtCtaINSSPatronal 
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   9
         Text            =   "11111"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox TxtCtaSueldos 
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   8
         Text            =   "11111"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtCtaPrevAquinaldo 
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   7
         Text            =   "11111"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxtCtaPrevVacaciones 
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   6
         Text            =   "11111"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TxtCtaOtrosIngresos 
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   5
         Text            =   "11111"
         Top             =   3360
         Width           =   1335
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   390
         Left            =   3000
         TabIndex        =   12
         Top             =   480
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":3366
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdProAguinaldo 
         Height          =   390
         Left            =   3000
         TabIndex        =   13
         Top             =   930
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":3868
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdPrevVaca 
         Height          =   390
         Left            =   3000
         TabIndex        =   14
         Top             =   1400
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":3D6A
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdInssPatronal 
         Height          =   390
         Left            =   3000
         TabIndex        =   15
         Top             =   1880
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":426C
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdInatec 
         Height          =   390
         Left            =   3000
         TabIndex        =   16
         Top             =   2360
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":476E
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdHorasExtras 
         Height          =   390
         Left            =   3000
         TabIndex        =   17
         Top             =   2840
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":4C70
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdOtrosIngresos 
         Height          =   390
         Left            =   3000
         TabIndex        =   18
         Top             =   3320
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":5172
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   390
         Left            =   3000
         TabIndex        =   56
         Top             =   3840
         Visible         =   0   'False
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":5674
         ImageAlignment  =   0
      End
      Begin VB.Label Label9 
         Caption         =   "Cuenta Subsidio"
         Height          =   255
         Left            =   0
         TabIndex        =   57
         Top             =   3840
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Horas Extras"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "INATEC"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "INSS Patronal"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Cuenta Sueldos:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label75 
         Caption         =   "Prov Aguinaldo:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label77 
         Caption         =   "Prov Vacaciones:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label70 
         Caption         =   "Otros Ingresos:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3360
         Width           =   1335
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   4575
      Left            =   3840
      TabIndex        =   26
      Top             =   1200
      Width           =   3855
      _Version        =   786432
      _ExtentX        =   6800
      _ExtentY        =   8070
      _StockProps     =   79
      Caption         =   "Cuentas Credito"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox TxtBanco 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   52
         Text            =   "11111"
         Top             =   4125
         Width           =   1455
      End
      Begin VB.TextBox TxtCtaInssPatronalPagar 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   49
         Text            =   "11111"
         Top             =   3165
         Width           =   1455
      End
      Begin VB.TextBox TxtCtaNominaxPagar 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   45
         Text            =   "11111"
         Top             =   3645
         Width           =   1455
      End
      Begin VB.TextBox TxtCtaInatecxPagar 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   42
         Text            =   "11111"
         Top             =   2295
         Width           =   1455
      End
      Begin VB.TextBox TxtCtaPasVacaciones 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   39
         Text            =   "11111"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox TxtCtaPasAguinaldo 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   36
         Text            =   "11111"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxtCtaInssxPagar 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   29
         Text            =   "11111"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TxtCtaIrxPagar 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   28
         Text            =   "11111"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox TxtCtaPrestamo 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   27
         Text            =   "11111"
         Top             =   2760
         Width           =   1455
      End
      Begin XtremeSuiteControls.PushButton CmdInssPagar 
         Height          =   390
         Left            =   3240
         TabIndex        =   30
         Top             =   1320
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":5B76
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdIRpagar 
         Height          =   390
         Left            =   3240
         TabIndex        =   31
         Top             =   1785
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":6078
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdPrestamo 
         Height          =   390
         Left            =   3240
         TabIndex        =   32
         Top             =   2715
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":657A
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdPasivoAguinaldo 
         Height          =   390
         Left            =   3240
         TabIndex        =   37
         Top             =   360
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":6A7C
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdPasivoVaca 
         Height          =   390
         Left            =   3240
         TabIndex        =   40
         Top             =   840
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":6F7E
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdInstecPagar 
         Height          =   390
         Left            =   3240
         TabIndex        =   43
         Top             =   2280
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":7480
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdNominaPagar 
         Height          =   390
         Left            =   3240
         TabIndex        =   46
         Top             =   3600
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":7982
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   390
         Left            =   3240
         TabIndex        =   50
         Top             =   3120
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":7E84
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   390
         Left            =   3240
         TabIndex        =   53
         Top             =   4080
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmCuentasContables.frx":8386
         ImageAlignment  =   0
      End
      Begin VB.Label Label8 
         Caption         =   "Cuenta Banco"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   4125
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "INSS Patronal Pag"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   3165
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Nomina x Pagar"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   3645
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "INATEC x Pagar"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   2295
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Pasivo Vacac"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Pasivo Aguinaldo"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label71 
         Caption         =   "INSS x Pagar"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label73 
         Caption         =   "IR x Pagar"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label69 
         Caption         =   "Cuenta Prestamo:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   2760
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmCuentasContables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGrabar_Click()
 Dim CodEmpleado As String
 
 CodEmpleado = Me.TxtCodEmpleado.Text
 
 Me.AdoHistorico.RecordSource = "SELECT  * From Historico Where (CodEmpleado = " & CodEmpleado & ")"
 Me.AdoHistorico.Refresh
 If Not Me.AdoHistorico.Recordset.EOF Then
 
   '/////////////////////////////////////////////////////////////////////////////////////////////
   '//////////////////////////CUENTAS DE DEBITO ////////////////////////////////////////////////
   '////////////////////////////////////////////////////////////////////////////////////////////
     
 
         If Me.TxtCtaSueldos.Text <> "" Then
           Me.AdoHistorico.Recordset("CuentaSueldos") = Me.TxtCtaSueldos.Text
         End If
        
         If Me.TxtCtaPrevAquinaldo.Text <> "" Then
           Me.AdoHistorico.Recordset("ProvAguinaldo") = Me.TxtCtaPrevAquinaldo.Text
         End If
        
         If Me.TxtCtaPrevVacaciones.Text <> "" Then
             Me.AdoHistorico.Recordset("ProvVacaciones") = Me.TxtCtaPrevVacaciones.Text
         End If
         
         If Me.TxtCtaINSSPatronal.Text <> "" Then
           Me.AdoHistorico.Recordset("INSSPatronal") = Me.TxtCtaINSSPatronal.Text
         End If
         
         If Me.TxtCtaINATEC.Text <> "" Then
            Me.AdoHistorico.Recordset("INATEC") = Me.TxtCtaINATEC.Text
         End If
         
         If Me.TxtCtaHorasExtras.Text <> "" Then
           Me.AdoHistorico.Recordset("CuentaHorasExtra") = Me.TxtCtaHorasExtras.Text
         End If
         
         If Me.TxtCtaOtrosIngresos.Text <> "" Then
           Me.AdoHistorico.Recordset("CuentaOtrosIngresos") = Me.TxtCtaOtrosIngresos.Text
         End If
   
         '/////////////////////////////////////////////////////////////////////////////////////////////
         '//////////////////////////CUENTAS DE CREDITO////////////////////////////////////////////////
         '////////////////////////////////////////////////////////////////////////////////////////////
         If Me.TxtCtaPasAguinaldo.Text <> "" Then
           Me.AdoHistorico.Recordset("AguinaldoxPagar") = Me.TxtCtaPasAguinaldo.Text
         End If
        
         If Me.TxtCtaPasVacaciones.Text <> "" Then
           Me.AdoHistorico.Recordset("VacacionesxPagar") = Me.TxtCtaPasVacaciones.Text
         End If
        
         If Me.TxtCtaInssxPagar.Text <> "" Then
           Me.AdoHistorico.Recordset("INSSxPagar") = Me.TxtCtaInssxPagar.Text
         End If
         
         If Me.TxtCtaInatecxPagar.Text <> "" Then
           Me.AdoHistorico.Recordset("INATECxPagar") = Me.TxtCtaInatecxPagar.Text
         End If
         
         If Me.TxtCtaIrxPagar.Text <> "" Then
           Me.AdoHistorico.Recordset("IRxPagar") = Me.TxtCtaIrxPagar.Text
         End If
         
         If Me.TxtCtaPrestamo.Text <> "" Then
           Me.AdoHistorico.Recordset("PrestamoxPagar") = Me.TxtCtaPrestamo.Text
         End If
         
         If Me.TxtCtaNominaxPagar.Text <> "" Then
           Me.AdoHistorico.Recordset("NominaxPagar") = Me.TxtCtaNominaxPagar.Text
         End If
         
         If Me.TxtCtaNominaxPagar.Text <> "" Then
           Me.AdoHistorico.Recordset("INSSPatronalPagar") = Me.TxtCtaInssPatronalPagar.Text
         End If
         
         If Me.TxtBanco.Text <> "" Then
           Me.AdoHistorico.Recordset("CuentaBanco") = Me.TxtBanco.Text
         End If
         
         If Me.TxtBanco.Text <> "" Then
           Me.AdoHistorico.Recordset("CuentaSubsidio") = Me.TxtCuentaSubsidio.Text
         End If
         
    Me.AdoHistorico.Recordset.Update
 End If
End Sub

Private Sub CmdHorasExtras_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaHorasExtras.Text = FrmConsulta.CuentaContable
End Sub

Private Sub CmdInatec_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaINATEC.Text = FrmConsulta.CuentaContable
End Sub

Private Sub CmdInssPagar_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaInssxPagar.Text = FrmConsulta.CuentaContable
End Sub

Private Sub CmdInssPatronal_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaINSSPatronal.Text = FrmConsulta.CuentaContable
End Sub

Private Sub CmdInstecPagar_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaInatecxPagar.Text = FrmConsulta.CuentaContable
End Sub

Private Sub CmdIRpagar_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaIrxPagar.Text = FrmConsulta.CuentaContable
End Sub

Private Sub CmdNominaPagar_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaNominaxPagar.Text = FrmConsulta.CuentaContable
End Sub

Private Sub CmdOtrosIngresos_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaOtrosIngresos.Text = FrmConsulta.CuentaContable
End Sub

Private Sub CmdPasivoAguinaldo_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaPasAguinaldo.Text = FrmConsulta.CuentaContable
End Sub

Private Sub CmdPasivoVaca_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaPasVacaciones.Text = FrmConsulta.CuentaContable
End Sub

Private Sub CmdPrestamo_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaPrestamo.Text = FrmConsulta.CuentaContable
End Sub

Private Sub CmdPrevVaca_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaPrevVacaciones.Text = FrmConsulta.CuentaContable
End Sub

Private Sub CmdProAguinaldo_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaPrevAquinaldo.Text = FrmConsulta.CuentaContable
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaSueldos.Text = FrmConsulta.CuentaContable
End Sub

Private Sub Form_Load()
With Me.AdoHistorico
   .ConnectionString = Conexion
End With
End Sub

Private Sub PushButton2_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtBanco.Text = FrmConsulta.CuentaContable
End Sub

Private Sub PushButton1_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCtaInssPatronalPagar.Text = FrmConsulta.CuentaContable
End Sub

Private Sub PushButton3_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtCuentaSubsidio.Text = FrmConsulta.CuentaContable
End Sub

Private Sub txtCodEmpleado_Change()
 Dim CodEmpleado As String
 
 CodEmpleado = Me.TxtCodEmpleado.Text
 
 Me.AdoHistorico.RecordSource = "SELECT  * From Historico Where (CodEmpleado = " & CodEmpleado & ")"
 Me.AdoHistorico.Refresh
 If Not Me.AdoHistorico.Recordset.EOF Then
 
   '/////////////////////////////////////////////////////////////////////////////////////////////
   '//////////////////////////CUENTAS DE DEBITO ////////////////////////////////////////////////
   '////////////////////////////////////////////////////////////////////////////////////////////
 
   If Not IsNull(Me.AdoHistorico.Recordset("CuentaSueldos")) Then
     Me.TxtCtaSueldos.Text = Me.AdoHistorico.Recordset("CuentaSueldos")
   End If
  
   If Not IsNull(Me.AdoHistorico.Recordset("ProvAguinaldo")) Then
     Me.TxtCtaPrevAquinaldo.Text = Me.AdoHistorico.Recordset("ProvAguinaldo")
   End If
  
   If Not IsNull(Me.AdoHistorico.Recordset("ProvVacaciones")) Then
     Me.TxtCtaPrevVacaciones.Text = Me.AdoHistorico.Recordset("ProvVacaciones")
   End If
   
   If Not IsNull(Me.AdoHistorico.Recordset("INSSPatronal")) Then
     Me.TxtCtaINSSPatronal.Text = Me.AdoHistorico.Recordset("INSSPatronal")
   End If
   
   If Not IsNull(Me.AdoHistorico.Recordset("INATEC")) Then
     Me.TxtCtaINATEC.Text = Me.AdoHistorico.Recordset("INATEC")
   End If
   
   If Not IsNull(Me.AdoHistorico.Recordset("CuentaHorasExtra")) Then
     Me.TxtCtaHorasExtras.Text = Me.AdoHistorico.Recordset("CuentaHorasExtra")
   End If
   
   If Not IsNull(Me.AdoHistorico.Recordset("CuentaOtrosIngresos")) Then
     Me.TxtCtaOtrosIngresos.Text = Me.AdoHistorico.Recordset("CuentaOtrosIngresos")
   End If
   
   '/////////////////////////////////////////////////////////////////////////////////////////////
   '//////////////////////////CUENTAS DE CREDITO////////////////////////////////////////////////
   '////////////////////////////////////////////////////////////////////////////////////////////
    If Not IsNull(Me.AdoHistorico.Recordset("AguinaldoxPagar")) Then
     Me.TxtCtaPasAguinaldo.Text = Me.AdoHistorico.Recordset("AguinaldoxPagar")
   End If
  
   If Not IsNull(Me.AdoHistorico.Recordset("VacacionesxPagar")) Then
     Me.TxtCtaPasVacaciones.Text = Me.AdoHistorico.Recordset("VacacionesxPagar")
   End If
  
   If Not IsNull(Me.AdoHistorico.Recordset("INSSxPagar")) Then
     Me.TxtCtaInssxPagar.Text = Me.AdoHistorico.Recordset("INSSxPagar")
   End If
   
   If Not IsNull(Me.AdoHistorico.Recordset("INATECxPagar")) Then
     Me.TxtCtaInatecxPagar.Text = Me.AdoHistorico.Recordset("INATECxPagar")
   End If
   
   If Not IsNull(Me.AdoHistorico.Recordset("IRxPagar")) Then
     Me.TxtCtaIrxPagar.Text = Me.AdoHistorico.Recordset("IRxPagar")
   End If
   
   If Not IsNull(Me.AdoHistorico.Recordset("PrestamoxPagar")) Then
     Me.TxtCtaPrestamo.Text = Me.AdoHistorico.Recordset("PrestamoxPagar")
   End If
   
   If Not IsNull(Me.AdoHistorico.Recordset("NominaxPagar")) Then
     Me.TxtCtaNominaxPagar.Text = Me.AdoHistorico.Recordset("NominaxPagar")
   End If
   
   If Not IsNull(Me.AdoHistorico.Recordset("INSSPatronalPagar")) Then
     Me.TxtCtaInssPatronalPagar.Text = Me.AdoHistorico.Recordset("INSSPatronalPagar")
   End If
   
    If Not IsNull(Me.AdoHistorico.Recordset("CuentaBanco")) Then
     Me.TxtBanco.Text = Me.AdoHistorico.Recordset("CuentaBanco")
   End If
   
    If Not IsNull(Me.AdoHistorico.Recordset("CuentaSubsidio")) Then
     Me.TxtCuentaSubsidio.Text = Me.AdoHistorico.Recordset("CuentaSubsidio")
   End If
   

 End If
 

End Sub
