VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmIncentivosGrupo 
   Appearance      =   0  'Flat
   Caption         =   "Incentivo Por Grupos"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   375
      Left            =   2640
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      Caption         =   "AdoEmpleados"
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
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton CmdProcesar 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.PictureBox Picture6 
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   8235
      TabIndex        =   2
      Top             =   1320
      Width           =   8295
      Begin VB.CheckBox ChkProcesar 
         Caption         =   "Procesar todos los empleados del filtro."
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   3975
      End
      Begin MSMask.MaskEdBox TxtMonto 
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label LblNombreEmpleado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   7575
      End
      Begin VB.Label LblIncentivo 
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label27 
         Caption         =   "Monto"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin XtremeSuiteControls.ProgressBar PBCalcNomina 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   7575
      _Version        =   786432
      _ExtentX        =   13361
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin VB.Label lbltitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "CREACION DE  VIATICOS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   0
      X2              =   8400
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   0
      Picture         =   "FrmIncentivosGrupo.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmIncentivosGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdProcesar_Click()
 Dim CodigoEmpleado As String, Nombres As String, NumeroCedula As String, NumeroInss As String
 Dim CodEmpleado As Double

If Me.ChkProcesar.Value = 1 Then
          Me.PBCalcNomina.Min = 0
          Me.PBCalcNomina.Max = FrmListadoEmpleado.rs.RecordCount
          Me.PBCalcNomina.Value = 0
          
          Do While Not FrmListadoEmpleado.rs.EOF  'esto nos sirve pa leer los datos desde
               
               
                    CodigoEmpleado = FrmListadoEmpleado.rs("CodEmpleado1")
                    Nombres = FrmListadoEmpleado.rs("Nombres")
                    NumeroCedula = FrmListadoEmpleado.rs("NumCedula")
                    NumeroInss = FrmListadoEmpleado.rs("NumeroInss")
        
                    Me.LblNombreEmpleado.Caption = CodigoEmpleado & "-" & Nombres
                    Me.PBCalcNomina.Value = Me.PBCalcNomina.Value + 1
                    FrmListadoEmpleado.rs.MoveNext
                    DoEvents
         
          Loop
          
Else
                
                    CodEmpleado = FrmListadoEmpleado.DbgrProducto.Columns(0).Text
                    CodigoEmpleado = FrmListadoEmpleado.rs("CodEmpleado1")
                    Nombres = FrmListadoEmpleado.rs("Nombres")
                    NumeroCedula = FrmListadoEmpleado.rs("NumCedula")
                    NumeroInss = FrmListadoEmpleado.rs("NumeroInss")
        
                    Me.LblNombreEmpleado.Caption = CodigoEmpleado & "-" & Nombres
                              
                            
        Me.AdoEmpleados.ConnectionString = Conexion
        Me.AdoEmpleados.RecordSource = "SELECT  Empleado.* From Empleado Where (CodEmpleado1 = " & CodigoEmpleado & ") AND (Activo = 1)"
        Me.AdoEmpleados.Refresh
        If Not Me.AdoEmpleados.Recordset.EOF Then
        
             res = Bitacora(Now, NombreUsuario, "Empleados", "Editando Viaticos: " & Nombres)
        
                 If IsNumeric(Me.TxtMonto.Text) Then
                 
                    AdoEmpleados.Recordset("MontoViatico") = Me.TxtMonto.Text
                    AdoEmpleados.Recordset.Update

                    MsgBox "Registro Agregado con Existo!!", vbExclamation, "Zeus Nominas"
                    Unload Me
                 End If
                  
        End If
        
End If

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
 Dim CodigoEmpleado As String, Nombres As String, NumeroCedula As String, NumeroInss As String
 Dim CodEmpleado As Double

                    CodigoEmpleado = FrmListadoEmpleado.rs("CodEmpleado1")
                    Nombres = FrmListadoEmpleado.rs("Nombres")
                    NumeroCedula = FrmListadoEmpleado.rs("NumCedula")
                    NumeroInss = FrmListadoEmpleado.rs("NumeroInss")

                    Me.AdoEmpleados.ConnectionString = Conexion
                    Me.AdoEmpleados.RecordSource = "SELECT  Empleado.* From Empleado Where (CodEmpleado1 = " & CodigoEmpleado & ") AND (Activo = 1)"
                    Me.AdoEmpleados.Refresh
                    If Not Me.AdoEmpleados.Recordset.EOF Then
                      If Not IsNull(Me.AdoEmpleados.Recordset("MontoViatico")) Then
                       Me.TxtMonto.Text = Me.AdoEmpleados.Recordset("MontoViatico")
                      Else
                       Me.TxtMonto.Text = 0
                      End If
                    Else
                       Me.TxtMonto.Text = 0
                    End If
        
                    Me.LblNombreEmpleado.Caption = CodigoEmpleado & "-" & Nombres
                    DoEvents
End Sub

