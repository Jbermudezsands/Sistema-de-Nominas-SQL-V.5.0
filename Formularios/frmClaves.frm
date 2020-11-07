VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmClaves 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   3720
   ClientLeft      =   15
   ClientTop       =   405
   ClientWidth     =   4065
   HelpContextID   =   6
   Icon            =   "frmClaves.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmClaves.frx":030A
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   4095
      TabIndex        =   11
      Top             =   0
      Width           =   4095
      Begin VB.Image Image2 
         Height          =   1020
         Left            =   0
         Picture         =   "frmClaves.frx":074C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1290
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   6720
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Registro de Usuarios"
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
         Left            =   1440
         TabIndex        =   12
         Top             =   360
         Width           =   2400
      End
   End
   Begin VB.TextBox TxtClave 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox TxtNivel 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton CmdSalir 
      DownPicture     =   "frmClaves.frx":1152
      Height          =   375
      Left            =   2400
      MouseIcon       =   "frmClaves.frx":2C34
      MousePointer    =   99  'Custom
      Picture         =   "frmClaves.frx":3076
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton CmdBorrar 
      DownPicture     =   "frmClaves.frx":4B58
      Height          =   375
      Left            =   1920
      MouseIcon       =   "frmClaves.frx":663A
      MousePointer    =   99  'Custom
      Picture         =   "frmClaves.frx":6A7C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton CmdGrabar 
      DownPicture     =   "frmClaves.frx":855E
      Height          =   375
      Left            =   480
      MouseIcon       =   "frmClaves.frx":A040
      MousePointer    =   99  'Custom
      Picture         =   "frmClaves.frx":A482
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc DtaUsuario 
      Height          =   375
      Left            =   360
      Top             =   6240
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "DtaUsuario"
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
   Begin MSAdodcLib.Adodc DtaNacceso 
      Height          =   375
      Left            =   600
      Top             =   4440
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "DtaNacceso"
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
   Begin MSDataListLib.DataCombo DBNombreUsuario 
      Bindings        =   "frmClaves.frx":BF64
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "NombreUsuario"
      Text            =   ""
   End
   Begin VB.TextBox TxtCodpasword 
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Text            =   "TxtCodpasword"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label LblClave 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave de Acceso:"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel del Usuario:"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Usuario:"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Claves de Usuarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   5040
      Width           =   2415
   End
End
Attribute VB_Name = "frmClaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CmdBorrar_Click()
 On Error GoTo TipoErrs
 Dim Respuesta, Rsp
'Elimino el registro activo en la pantalla
  Set Rsp = DtaUsuario.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando al Empleado: " & DBNombreUsuario.Text)
   If Respuesta = 6 Then
     Rsp.Delete
      DBNombreUsuario.Text = ""
      Me.TxtClave.Text = ""
      Me.TxtNivel.Text = ""
      
'      DtaUsuario.Recordset.MoveLast
'      DtaUsuario.Recordset.MovePrevious
      Salida = False
   End If
 Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdGrabar_Click()
'On Error GoTo TipoErrs
Dim Id As Integer
 Me.DtaUsuario.Refresh
  Salida = False
  Dim Rsp
  'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
    If val(NivelAcceso) <> 0 Then
      If val(val(TxtNivel.Text)) > NivelAcceso Then
        MsgBox "No Puede Cambiar el Nivel", vbExclamation, "Sistema de Nominas"
        Exit Sub
      End If
    End If
      If Controlador = 1 Then
          If TxtClave.Text = "" Then
           MsgBox "Verifique la Clave", vbInformation, "Sistema de Nominas"
           TxtClave.SetFocus
           Controlador = 1
           Salida = False
           Exit Sub
          Else
            If TxtClave.Text = VarClave Then
               'DtaUsuario.Recordset.Edit
               DtaUsuario.Recordset.Fields("Clave") = TxtClave.Text
               Me.DtaUsuario.Recordset.Fields("NivelAcceso") = val(TxtNivel.Text)
               DtaUsuario.Recordset.Update
               Controlador = 0
               LblClave.Caption = "Clave de Acceso"
               Salida = False
               Exit Sub
            Else
               MsgBox "Error Repita nuevamente las Claves", vbCritical, "Sistema de Nominas"
               TxtClave.Text = ""
               TxtClave.SetFocus
               Controlador = 1
               Salida = False
               Exit Sub
            End If
          End If
      Else 'Si el controlador es distinto de 1
        DtaUsuario.Refresh
        Do While Not DtaUsuario.Recordset.EOF
          If DtaUsuario.Recordset("NombreUsuario") = DBNombreUsuario.Text Then
            If DtaUsuario.Recordset("Clave") <> TxtClave.Text Then
             'Asigno una variable controlador para saber si se modifico clave.
                Controlador = 1
                VarClave = TxtClave.Text
                LblClave.Caption = "Repita la Clave"
                TxtClave.Text = ""
                TxtClave.SetFocus
                Salida = False
                Exit Sub
            Else 'Si la clave Digitada es igual a existente ejecuta
               'DtaUsuario.Recordset.Edit
               DtaUsuario.Recordset.Fields("NivelAcceso") = val(TxtNivel.Text)
               DtaUsuario.Recordset.Update
               DtaUsuario.Recordset.MoveLast
               DtaUsuario.Recordset.MovePrevious
               Salida = False
               Exit Sub
            End If 'Fin del If Pasword
                     
          End If 'fin del If NombreEmpleado
          DtaUsuario.Recordset.MoveNext
        Loop
            'En esta Parte Si no Existe el Usuario lo Agrega
            If val(TxtNivel.Text) = 0 Or TxtClave.Text = "" Then
              MsgBox "Los Campos no pueden quedar Vacio"
              TxtNivel.SetFocus
              Salida = False
              Exit Sub
            Else
              If TieneDatos = 0 Then
                TieneDatos = 1
                VarClave = TxtClave.Text
                LblClave.Caption = "Repita la Clave"
                TxtClave.Text = ""
                TxtClave.SetFocus
                Salida = False
                Exit Sub
              Else
               If VarClave = TxtClave.Text Then
                VarClave = ""
                Me.DtaUsuario.Refresh
                If DtaUsuario.Recordset.EOF Then
                 Id = 1
                Else
                  Me.DtaUsuario.Recordset.MoveLast
                  Id = Me.DtaUsuario.Recordset("CodUsuario") + 1
                End If
                DtaUsuario.Recordset.AddNew
                DtaUsuario.Recordset("CodUsuario") = Id
                DtaUsuario.Recordset.Fields("NombreUsuario") = DBNombreUsuario.Text
                DtaUsuario.Recordset.Fields("Clave") = TxtClave.Text
                DtaUsuario.Recordset.Fields("NivelAcceso") = val(TxtNivel.Text)
                TxtCodpasword.Text = Id
                DtaUsuario.Recordset.Update
                LblClave.Caption = "Clave de Acceso"
                DBNombreUsuario.Text = ""
                DtaUsuario.Recordset.MoveLast
                DtaUsuario.Recordset.MovePrevious
                
                
                
               Me.DtaNacceso.Refresh
                'Agrego un Dato de Acceso para el Usuario
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Editar Niveles"
                DtaNacceso.Recordset.Update
                
                    
                'Agrego una dato de Acceso para el usuario
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Registro Empleados"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Registro Empleados"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Registro Empleados"
                DtaNacceso.Recordset.Update
                
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Tabla Anotaciones"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Tabla Anotaciones"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Tabla Anotaciones"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Tabla Departamentos"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Tabla Departamentos"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Tabla Departamentos"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Tabla Cargo"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Tabla Cargo"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Tabla Cargo"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Tabla Cargo"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Tabla Incapacidad"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Tabla Incapacidad"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Tabla Incapacidad"
                DtaNacceso.Recordset.Update
                
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Tabla Incentivos"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Tabla Incentivos"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Tabla Incentivos"
                DtaNacceso.Recordset.Update
                                
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Tabla Tipo Incapacidad"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Tabla Tipo Incapacidad"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Tabla Tipo Incapacidad"
                DtaNacceso.Recordset.Update
                
                'Agrego una dato de Acceso para el usuario
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Tabla Tipo Deducciones"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Tabla Tipo Deducciones"
                DtaNacceso.Recordset.Update
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Tabla Tipo Deducciones"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Tabla Tipo Deducciones"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Tabla Tipo Subsidio"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Tabla Tipo Subsidio"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Tabla Tipo Subsidio"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Tabla Tipo Comision"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Tabla Tipo Comision"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Tabla Tipo Comision"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Tabla Tipo Destajo"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Tabla Tipo Destajo"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Tabla Tipo Destajo"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Tabla Tipo Nomina"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Tabla Tipo Nomina"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Tabla Tipo Nomina"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Tabla Tipo INSS/IR"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Tabla Tipo INSS/IR"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Tabla Tipo INSS/IR"
                DtaNacceso.Recordset.Update
                
                'Agrego una dato de Acceso para el usuario
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Tabla Listado Nominas"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Tabla Listado Nominas"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Tabla Listado Nominas"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Calcular Nominas"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Activar Nominas"
                DtaNacceso.Recordset.Update
                               
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Movimientos de la Nomina"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Movimientos de la Nomina"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Calcular 13vo y Vacaciones"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Calcular 13vo y Vacaciones"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Nomina de Subsidio"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Nomina de Subsidio"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Despido y Renuncias"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Despido y Renuncias"
                DtaNacceso.Recordset.Update
                                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Entradas y Salidas"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Calcular Horas Extra o Falta"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Reportes"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Controles Personalizados"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Abrir Usuarios"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Grabar Usuarios"
                DtaNacceso.Recordset.Update
                
                DtaNacceso.Recordset.AddNew
                DtaNacceso.Recordset.Fields("CodUsuario") = TxtCodpasword.Text
                DtaNacceso.Recordset.Fields("AccesoModulo") = "Borrar Usuarios"
                DtaNacceso.Recordset.Update
                
                TieneDatos = 0
                Salida = False
                Exit Sub
                 Else
                  MsgBox "Las Claves no Coinciden", vbInformation, "Sistema de Nominas"
               End If 'Fin del If VarClave
              End If
            End If
  End If 'fin del if controlador
                  
Exit Sub
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub
Private Sub DBNombreUsuario_Change()
 Dim Usuario As Integer
 On Error GoTo TipoErrs

 
 'Al ejecutar algun cambio en el combo actualizo el nombre del Empleado
  
   
   DtaUsuario.Refresh
   
   Do While Not DtaUsuario.Recordset.EOF
     If DtaUsuario.Recordset("NombreUsuario") = DBNombreUsuario.Text Then
        If NivelAcceso < DtaUsuario.Recordset("NivelAcceso") Then
           DBNombreUsuario = ""
           MsgBox "No Tiene Permiso para esta Accion", vbInformation, "Sistema de Nominas"
           Salida = False
           Exit Sub
        End If
        TxtNivel.Text = DtaUsuario.Recordset("NivelAcceso")
        TxtClave.Text = DtaUsuario.Recordset("Clave")
        TxtCodpasword.Text = DtaUsuario.Recordset("CodUsuario")
        Exit Do
        
     Else
        TxtNivel.Text = ""
        TxtClave.Text = ""
     End If
       DtaUsuario.Recordset.MoveNext
   Loop
Salida = False
Exit Sub
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub DBNombreUsuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtNivel.SetFocus
 End If
End Sub



Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub Form_Load()

With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Accesos"
   .Refresh
End With

With Me.DtaUsuario
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Usuarios"
   .Refresh
End With

frmClaves.CmdSalir.MousePointer = 99
frmClaves.CmdGrabar.MousePointer = 99
frmClaves.MousePointer = 99
Me.Top = 3500
Me.Left = 3500

End Sub
Private Sub Salir_Click()
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ValidaSalida ("en la Tabla Clave de Usuario")
If Contesta Then
  CmdGrabar.Value = True
  
  Unload Me
Else
    Salida = False
    Unload Me
   
End If
End Sub

Private Sub TxtClave_Change()
Salida = True
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  CmdGrabar.Value = True
  Else
   Evaluar = False
  End If
End Sub



Private Sub TxtNivel_Change()
PreparaSalida
End Sub

Private Sub TxtNivel_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   TxtClave.SetFocus
  Else
   Evaluar = False
  End If
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
