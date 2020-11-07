VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPermisos 
   Caption         =   "Permisos"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data dtaServidor 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   5040
      TabIndex        =   30
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   29
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   6000
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc adoEmpleado 
      Height          =   375
      Left            =   3120
      Top             =   6120
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
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
      Connect         =   $"frmPermisos.frx":0000
      OLEDBString     =   $"frmPermisos.frx":0088
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Empleado"
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
   Begin MSAdodcLib.Adodc adoPermiso 
      Height          =   375
      Left            =   2040
      Top             =   6120
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"frmPermisos.frx":0110
      OLEDBString     =   $"frmPermisos.frx":0198
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Permisos"
      Caption         =   "Permiso"
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
   Begin MSAdodcLib.Adodc adoAsistencia 
      Height          =   330
      Left            =   2880
      Top             =   6120
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"frmPermisos.frx":0220
      OLEDBString     =   $"frmPermisos.frx":02A8
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Empleado"
      Caption         =   "Asistencia Diaria"
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
   Begin VB.Frame Frame3 
      Caption         =   "Permiso"
      Height          =   2535
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   7455
      Begin VB.CheckBox chkRegreso 
         Caption         =   "Pendiente regreso"
         Height          =   255
         Left            =   5040
         TabIndex        =   9
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox chkJustificado 
         Caption         =   "Justificado"
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtMotivo 
         Height          =   615
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   25
         Text            =   " "
         Top             =   1680
         Width           =   5295
      End
      Begin MSComCtl2.DTPicker dtpFInicio 
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   23068673
         CurrentDate     =   38570
      End
      Begin MSMask.MaskEdBox mskPermisoHoraInicio 
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm AM/PM"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPermisoHoraRegreso 
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm AM/PM"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         Caption         =   "Motivo:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         Height          =   195
         Left            =   3120
         TabIndex        =   24
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         Height          =   195
         Left            =   3120
         TabIndex        =   23
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   870
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Asistencia Diaria"
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   7455
      Begin VB.CheckBox chkSalida 
         Caption         =   "Esta Laborando este dia"
         Height          =   255
         Left            =   5160
         TabIndex        =   20
         Top             =   960
         Width           =   2175
      End
      Begin MSMask.MaskEdBox mskAsistHoraEntrada 
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm AM/PM"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpFechEntrada 
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   23068673
         CurrentDate     =   38570
      End
      Begin MSMask.MaskEdBox mskAsistHoraSalida 
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm AM/PM"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpFecSalida 
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   23068673
         CurrentDate     =   38570
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Salida"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hora Salida:"
         Height          =   195
         Left            =   2880
         TabIndex        =   18
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hora Entrada"
         Height          =   195
         Left            =   2880
         TabIndex        =   16
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Entrada"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Empleado"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   4440
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cboCodigo 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Text            =   " "
         Top             =   480
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   23068673
         CurrentDate     =   38570
      End
      Begin VB.Label lblNombre 
         Caption         =   " "
         Height          =   495
         Left            =   840
         TabIndex        =   12
         Top             =   960
         Width           =   6375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   3120
         TabIndex        =   11
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sCodEmpl As String
Public dFecha As Date
Public sCodTipoNomina As String





Private Sub cboCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
Me.dtpFecha.SetFocus
End If

End Sub

Private Sub cmdAgregar_Click()

Dim dFecha As Date
Dim sFechaEntrada As String

   If Me.cmdAgregar.Caption = "&Agregar" Then
        
      Me.cmdModificar.Enabled = False
      Me.dtpFInicio.SetFocus
      Me.cmdAgregar.Caption = "&Guardar"
   
     
   ElseIf Trim(Me.txtMotivo.Text) <> "" And Chequear_Hora(Me.mskPermisoHoraInicio.Text) Then
    
      
    
      Me.adoPermiso.Recordset.AddNew
      Me.adoPermiso.Recordset.Fields("CodEmpleado") = sCodEmpl
      Me.adoPermiso.Recordset.Fields("CodTipoNomina") = sCodTipoNomina
      Me.adoPermiso.Recordset.Fields("Fecha") = Me.dtpFInicio.Value
      Me.adoPermiso.Recordset.Fields("HoraInicio") = Me.mskPermisoHoraInicio.Text
      Me.adoPermiso.Recordset.Fields("Motivo") = Me.txtMotivo.Text
      Me.adoPermiso.Recordset.Fields("RegresoPendiente") = Me.chkRegreso.Value
      Me.adoPermiso.Recordset.Fields("Justificado") = Me.chkJustificado.Value
      
      If Me.mskPermisoHoraRegreso.Text <> "__:__:__" And Chequear_Hora(Me.mskPermisoHoraRegreso.Text) Then
         Me.adoPermiso.Recordset.Fields("HoraFinal") = Me.mskPermisoHoraRegreso.Text
      End If
      Me.adoPermiso.Recordset.Update
      
      dFecha = Me.dtpFecha.Value
      
      sFechaEntrada = Mid$(dFecha, 7, 4) & "-" & Mid$(dFecha, 4, 2) & "-" & Mid$(dFecha, 1, 2)
      
      
      Me.adoAsistencia.CommandType = adCmdText
      Me.adoAsistencia.RecordSource = "SELECT CodEmpleado, FechaEntrada, FechaSalida, HoraEntrada, HoraSalida, bActivo FROM AsistenciaEmpleado WHERE [FechaEntrada] =CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102) AND [CodEmpleado] ='" & sCodEmpl & "'"
      Me.adoAsistencia.Refresh
      
           
      If Not Me.adoAsistencia.Recordset.EOF Then
           
         Me.adoAsistencia.Recordset.Fields("FechaEntrada") = Me.dtpFechEntrada.Value
         Me.adoAsistencia.Recordset.Fields("HoraEntrada") = Me.mskAsistHoraEntrada.Text
         
         If Chequear_Hora(Me.mskAsistHoraSalida.Text) Then
            Me.adoAsistencia.Recordset.Fields("FechaSalida") = Me.dtpFecSalida.Value
            Me.adoAsistencia.Recordset.Fields("HoraSalida") = Me.mskAsistHoraSalida.Text
         End If
         
         If Me.chkSalida.Value Then
            Me.adoAsistencia.Recordset.Fields("bActivo") = 1
         Else
            Me.adoAsistencia.Recordset.Fields("bActivo") = 0
         End If
              
         Me.adoAsistencia.Recordset.Update
                      
      End If
           
      
     Me.cmdModificar.Enabled = True
     Me.cmdAgregar.Caption = "&Agregar"

   End If



End Sub

Private Sub cmdBuscar_Click()

Dim sFechaEntrada As String

If Trim(Me.cboCodigo.Text) <> "" Then
   
   Me.adoEmpleado.Recordset.Find "CodEmpleado ='" & Trim(Me.cboCodigo.Text) & "'"
    
   If Not Me.adoEmpleado.Recordset.EOF Then
      
      Me.lblNombre.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
      sCodEmpl = Me.adoEmpleado.Recordset.Fields("CodEmpleado")
      sCodTipoNomina = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
      dFecha = Me.dtpFecha.Value
      
      'sFechaEntrada = Mid$(dFecha, 7, 4) & "-" & Mid$(dFecha, 4, 2) & "-" & Mid$(dFecha, 1, 2)
      
     ' Me.adoAsistencia.Recordset.Find "[FechaEntrada] = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102)"
      
      Me.adoAsistencia.CommandType = adCmdText
      Me.adoAsistencia.RecordSource = "SELECT CodEmpleado, FechaEntrada, FechaSalida, HoraEntrada, HoraSalida, bActivo FROM AsistenciaEmpleado WHERE [FechaEntrada] ='" & dFecha & "' AND [CodEmpleado] ='" & sCodEmpl & "'"
      Me.adoAsistencia.Refresh
      
      'Me.adoAsistencia.Recordset.Find "[FechaEntrada] = '" & dFecha & "' AND [CodEmpleado] ='" & sCodEmpl & "'"
      
      
      If Not Me.adoAsistencia.Recordset.EOF Then
         Me.dtpFechEntrada.Value = Me.adoAsistencia.Recordset.Fields("FechaEntrada")
         Me.mskAsistHoraEntrada.Text = Me.adoAsistencia.Recordset.Fields("HoraEntrada")
         If Not IsNull(Me.adoAsistencia.Recordset.Fields("FechaSalida")) Then
             Me.dtpFecSalida.Value = Me.adoAsistencia.Recordset.Fields("FechaSalida")
             Me.mskAsistHoraSalida.Text = Me.adoAsistencia.Recordset.Fields("HoraSalida")
         Else
             Me.lblNombre.Caption = Me.lblNombre.Caption & ", No tiene fecha y hora de salida para este dia, corrija estos datos"
         End If
         
         If Not IsNull(Me.adoAsistencia.Recordset.Fields("bActivo")) Then
            If Me.adoAsistencia.Recordset.Fields("bActivo") Then
               Me.chkSalida.Value = 1
            Else
               Me.chkSalida.Value = 0
            End If
         End If
         
         Me.adoPermiso.CommandType = adCmdText
         Me.adoPermiso.RecordSource = "SELECT * FROM Permisos WHERE [Fecha] ='" & dFecha & "' AND [CodEmpleado] ='" & sCodEmpl & "'"
         Me.adoPermiso.Refresh

         
         'Me.adoPermiso.Recordset.Find "[Fecha] ='" & dFecha & "' AND [CodEmpleado] ='" & sCodEmpl & "'"
         
         If Not Me.adoPermiso.Recordset.EOF Then
            Me.dtpFInicio.Value = Me.adoPermiso.Recordset.Fields("Fecha")
            Me.mskPermisoHoraInicio.Text = Me.adoPermiso.Recordset.Fields("HoraInicio")
            Me.txtMotivo.Text = Me.adoPermiso.Recordset.Fields("Motivo")
            
            If Me.adoPermiso.Recordset.Fields("Justificado") Then
               Me.chkJustificado.Value = 1
            Else
               Me.chkJustificado.Value = 0
            End If
            
            Me.cmdModificar.Enabled = False
            Me.cmdAgregar.Enabled = False
            
            If Me.adoPermiso.Recordset.Fields("RegresoPendiente") Then
               Me.chkRegreso.Value = 1
            
            Else
               
               Me.chkRegreso.Value = 0
               
               
            End If
            
            Me.cmdAgregar.Enabled = True
            Me.cmdModificar.Enabled = True
            
            If Not IsNull(Me.adoPermiso.Recordset.Fields("HoraFin")) Then
               Me.mskPermisoHoraRegreso.Text = Me.adoPermiso.Recordset.Fields("HoraFin")
            
            Else
               Me.mskPermisoHoraRegreso.Text = "__:__:__"
            
            End If
            
            'Me.adoPermiso.Recordset.Update
            
         'Sino tiene un permiso para este dia, limpiar valores
         Else
            
           Me.cmdAgregar.Enabled = True
           Me.dtpFInicio.Value = Me.dtpFecha.Value
           Me.mskPermisoHoraInicio.Text = "__:__:__"
           Me.txtMotivo.Text = " "
           Me.chkJustificado.Value = False
           Me.chkRegreso.Value = 1
            
            
         End If
         
         
      'Sino tiene asistencia ese dia, limpiamos los valores de la asistencia.
      Else
        
        Me.mskPermisoHoraInicio.Text = "__:__:__"
        Me.mskPermisoHoraRegreso.Text = "__:__:__"
        Me.txtMotivo.Text = " "
        Me.chkJustificado.Value = 0
        Me.lblNombre.Caption = Me.lblNombre.Caption & ", No asistio a trabajar el " & Me.dtpFecha.Value
        Me.mskAsistHoraEntrada.Text = "__:__:__"
        Me.mskAsistHoraSalida.Text = "__:__:__"
        Me.chkSalida.Value = 0
      
         
         
      End If
      
      
      Me.adoEmpleado.Refresh
      
   Else
   
      MsgBox "El empleado No. " & Me.cboCodigo.Text & " no se encuentra registrado", vbInformation, "Permisos - Nomina"
      Me.lblNombre.Caption = "Empleado No Encontrado"
      Me.mskAsistHoraEntrada.Text = "__:__:__"
      Me.mskAsistHoraSalida.Text = "__:__:__"
      Me.adoEmpleado.Refresh
      
   End If
    

End If



End Sub

Private Sub cmdModificar_Click()



If Trim(Me.txtMotivo.Text) <> "" And Chequear_Hora(Me.mskPermisoHoraInicio.Text) And Me.cmdModificar.Caption = "&Modificar" And sCodEmpl <> "" Then
  
  Me.cmdAgregar.Enabled = False
  Me.mskPermisoHoraInicio.SetFocus
  Me.cmdModificar.Caption = "&Guardar Cambios"
  
  

ElseIf Chequear_Hora(Me.mskPermisoHoraInicio.Text) And sCodEmpl <> "" Then
  
  Me.adoPermiso.CommandType = adCmdText
  Me.adoPermiso.RecordSource = "SELECT * FROM Permisos WHERE [Fecha] ='" & dFecha & "' AND [CodEmpleado] ='" & sCodEmpl & "'"
  Me.adoPermiso.Refresh
  
  Me.adoPermiso.Recordset.Fields("Fecha") = Me.dtpFInicio.Value
  Me.adoPermiso.Recordset.Fields("HoraInicio") = Me.mskPermisoHoraInicio.Text
  Me.adoPermiso.Recordset.Fields("Justificado") = Me.chkJustificado.Value
  
  
  If Me.chkRegreso.Value Then
   Me.adoPermiso.Recordset.Fields("RegresoPendiente") = 1
  Else
   Me.adoPermiso.Recordset.Fields("RegresoPendiente") = 0
  End If
  
  If Chequear_Hora(Me.mskPermisoHoraRegreso.Text) Then
    Me.adoPermiso.Recordset.Fields("HoraFin") = Me.mskPermisoHoraRegreso.Text
  End If
  
  dFecha = Me.dtpFecha.Value
      
  sFechaEntrada = Mid$(dFecha, 7, 4) & "-" & Mid$(dFecha, 4, 2) & "-" & Mid$(dFecha, 1, 2)
      
      
  Me.adoAsistencia.CommandType = adCmdText
  Me.adoAsistencia.RecordSource = "SELECT CodEmpleado, FechaEntrada, FechaSalida, HoraEntrada, HoraSalida, bActivo FROM AsistenciaEmpleado WHERE [FechaEntrada] =CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102) AND [CodEmpleado] ='" & sCodEmpl & "'"
  Me.adoAsistencia.Refresh
      
           
  If Not Me.adoAsistencia.Recordset.EOF Then
           
  Me.adoAsistencia.Recordset.Fields("FechaEntrada") = Me.dtpFechEntrada.Value
  Me.adoAsistencia.Recordset.Fields("HoraEntrada") = Me.mskAsistHoraEntrada.Text
         
  If Chequear_Hora(Me.mskAsistHoraSalida.Text) Then
     Me.adoAsistencia.Recordset.Fields("FechaSalida") = Me.dtpFecSalida.Value
     Me.adoAsistencia.Recordset.Fields("HoraSalida") = Me.mskAsistHoraSalida.Text
  End If
         
  If Me.chkSalida.Value Then
     Me.adoAsistencia.Recordset.Fields("bActivo") = 1
  Else
     Me.adoAsistencia.Recordset.Fields("bActivo") = 0
  End If
              
  Me.adoAsistencia.Recordset.Update
                      
  End If
  
  
  
  Me.adoPermiso.Recordset.Update
  Me.cmdAgregar.Enabled = True
  Me.cmdModificar.Caption = "&Modificar"
  

End If


End Sub

Private Sub cmdSalir_Click()

Unload Me


End Sub



Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
  Me.cmdBuscar.SetFocus

End If

End Sub

Private Sub Form_Activate()

Me.dtpFecha.Value = Mid$(Now, 1, 10)
Me.dtpFInicio.Value = Mid$(Now, 1, 10)
Me.cboCodigo.SetFocus



End Sub


Private Sub Form_Load()

Dim RutaServer As String
 Dim Server As String
 Dim Conexion As String


 Dim ConexionSTR1 As String
 Dim TxtClaveEntrada As String
'abro el archivo para solo lectura de la cadena de conexion
 Dim NextLine As String
 Dim Autorizado As Boolean
   Autorizado = False

 Open App.Path + "\SysInfo.dll" For Input As #1
  Do Until EOF(1)
   Line Input #1, NextLine
        ConexionSTR1 = Trim(NextLine)
   Loop
 Close #1
  
 
  Conexion = ConexionSTR1


Me.adoEmpleado.ConnectionString = Conexion
Me.adoEmpleado.CommandType = adCmdTable
Me.adoEmpleado.RecordSource = "Empleado"
Me.adoEmpleado.Refresh


Me.adoAsistencia.ConnectionString = Conexion
Me.adoAsistencia.CommandType = adCmdTable
Me.adoAsistencia.RecordSource = "AsistenciaEmpleado"
Me.adoAsistencia.Refresh

Me.adoPermiso.ConnectionString = Conexion
Me.adoPermiso.CommandType = adCmdTable
Me.adoPermiso.RecordSource = "Permisos"
Me.adoPermiso.Refresh


Do While Not Me.adoEmpleado.Recordset.EOF

  Me.cboCodigo.AddItem Me.adoEmpleado.Recordset.Fields("CodEmpleado")
  Me.adoEmpleado.Recordset.MoveNext


Loop

Me.adoEmpleado.Refresh

End Sub


Public Function Chequear_Hora(Hora As String) As Boolean

Dim sMinutos As String
Dim sHora As String
Dim sSegundos As String

sHora = Mid$(Hora, 1, 2)
sMinutos = Mid$(Hora, 4, 2)
sSegundos = Mid$(Hora, 7, 2)


If IsNumeric(sHora) And IsNumeric(sMinutos) And IsNumeric(sSegundos) Then

   Chequear_Hora = True

Else
   
   Chequear_Hora = False

End If






End Function
