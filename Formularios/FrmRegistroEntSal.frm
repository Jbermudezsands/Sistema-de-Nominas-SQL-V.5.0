VERSION 5.00
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{EAD60554-CF37-11D1-A050-70D904C10000}#3.0#0"; "MacButtn.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FrmRegistroEntSal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                               Registro de Entradas y Salidas"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "FrmRegistroEntSal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc DtaBusca 
      Height          =   375
      Left            =   1080
      Top             =   6720
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "DtaBusca"
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
   Begin MSAdodcLib.Adodc DtaSalida 
      Height          =   375
      Left            =   1200
      Top             =   6120
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "DtaSalida"
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
   Begin MSAdodcLib.Adodc DtaEmpleado 
      Height          =   375
      Left            =   1200
      Top             =   5640
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
      Caption         =   "DtaEmpleado"
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
   Begin MacintoshButton.MacButton CmdAceptar 
      Height          =   300
      Left            =   5280
      TabIndex        =   21
      Top             =   5640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Aceptar"
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   2040
      Top             =   5640
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.CommandButton CmdAceptar1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos del Sistema"
      Height          =   1695
      Left            =   3480
      TabIndex        =   14
      Top             =   2160
      Width           =   3375
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   2760
         Top             =   120
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "hora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Hora:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label LblFecha 
         Alignment       =   2  'Center
         Caption         =   "fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Control de Entradas y Salidas"
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   3375
      Begin VB.TextBox TxtSalida 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox TxtEntra 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label LblMensaje 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   495
         Left            =   840
         TabIndex        =   19
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Hora Salida"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Hora Entrada"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Foto del Empleado"
      Height          =   2055
      Left            =   4080
      TabIndex        =   8
      Top             =   120
      Width           =   2775
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1575
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2565
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin MSDataListLib.DataCombo DBCodigo 
         Bindings        =   "FrmRegistroEntSal.frx":0442
         Height          =   315
         Left            =   1680
         TabIndex        =   22
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodEmpleado"
         Text            =   ""
      End
      Begin VB.TextBox TxtCargo 
         Height          =   305
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox TxtApellido1 
         Height          =   305
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox TxtNombre1 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Cargo Empleado"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Apellido Empleado"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre Empleado"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Numero Empleado"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSCommLib.MSComm mscReloj2 
      Left            =   120
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm mscReloj 
      Left            =   840
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.Label LblPuerto 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   3960
      Width           =   6735
   End
End
Attribute VB_Name = "FrmRegistroEntSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AlarmTime, AlarmTiempo
Const conMinimized = 1#

Private Sub cmdAceptar_Click()
Dim HoraIni As Integer
Dim Hora As Integer
Dim Minutos As Integer
Dim HorasExtras As Double
Dim FechaHoy As Date
Dim SqlSalida As String

FechaHoy = Format(Now, "DD/MM/YYYY")
If DBCodigo.Text = "" Then
    Unload Me
    Exit Sub
End If

If txtNombre1.Text = "" Then
   MsgBox "Empleado no existe"
   Unload Me
   Exit Sub
End If
CodEmpleado = DBCodigo.Text
SqlSalida = "SELECT Salida.CodEmpleado, Salida.Fecha, Salida.HoraEntra, Salida.HoraSale, Salida.HorasExtras FROM Salida WHERE SAlida.CodEmpleado='" & CodEmpleado & "' AND Salida.Fecha= " & FechaHoy & " "
'SqlSalida = "SELECT CodEmpleado, Fecha, HoraEntra, HoraSale, HorasExtras From Salida WHERE     (CodEmpleado = '" & CodEmpleado & "') AND (Fecha = " & FechaHoy & ")"
DtaSalida.RecordSource = SqlSalida
DtaSalida.Refresh

If Not DtaSalida.Recordset.EOF Then
         'DtaSalida.Recordset.Edit
Else
         DtaSalida.Recordset.AddNew
End If
         DtaSalida.Recordset("CodEmpleado") = DBCodigo.Text
         DtaSalida.Recordset("Fecha") = Format(Now, "dd/mm/yyyy")
         DtaSalida.Recordset("HoraEntra") = Format(TxtEntra.Text, "long time")
         If Not TxtSalida.Text = "" Then
            DtaSalida.Recordset("horasale") = Format(TxtSalida.Text, "long time")
            If (CDate(TxtSalida.Text) - CDate("17:00")) > 0 Then
                HoraIni = 17
                Hora = Hour(TxtSalida.Text)
                Minutos = Minute(TxtSalida.Text)
                If HoraIni < Hora Then
                    HorasExtras = Hora - HoraIni
                    HorasExtras = HorasExtras + (Minutos * 1) / 60
                    DtaSalida.Recordset("HorasExtras") = HorasExtras
                Else
                    DtaSalida.Recordset("HorasExtras") = 0
                End If
            End If
         End If
        DtaSalida.Recordset.Update
        lblMensaje.Caption = ""
        DBCodigo.Text = ""
        lblTime.ForeColor = &H800000

Unload Me


End Sub


Private Sub DBCodigo_Change()
Dim Fecha As Date, HoraEntra As Variant
Destino = ""
DtaBusca.RecordSource = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Cargo.Cargo, Departamento.Departamento FROM Departamento INNER JOIN (Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) ON Departamento.CodDepartamento = Empleado.CodDepartamento WHERE (((Empleado.CodEmpleado)='" & DBCodigo.Text & "'))"
DtaBusca.Refresh
If Not DtaBusca.Recordset.EOF Then
   txtNombre1.Text = DtaBusca.Recordset("Nombre1")
   txtApellido1.Text = DtaBusca.Recordset("Apellido1")
   TxtCargo.Text = DtaBusca.Recordset("Cargo")
   If Dir(App.Path + "\Fotos\" & DBCodigo.Text & ".jpg") <> "" Then
           Destino = "c:\Nominas\Fotos\" & DBCodigo.Text & ".jpg"
        ElseIf Dir(App.Path + "\Fotos\" & DBCodigo.Text & ".gif") <> "" Then
           Destino = App.Path + "\Fotos\" & DBCodigo.Text & ".gif"
        ElseIf Dir(App.Path + "\Fotos\" & DBCodigo.Text & ".bmp") <> "" Then
           Destino = App.Path + "\Fotos\" & DBCodigo.Text & ".bmp"
   End If
        
     If Destino <> "" Then
         Image1.Picture = LoadPicture(Destino)
        Else
         Destino = App.Path + "\Fotos\Zw.bmp"
         Image1.Picture = LoadPicture(Destino)
     End If
Fecha = Format(Now, "DD/MM/YYYY")
DtaBusca.RecordSource = "SELECT Empleado.CodEmpleado, Salida.Fecha, Salida.HoraEntra, Salida.HoraSale FROM Empleado INNER JOIN Salida ON Empleado.CodEmpleado = Salida.CodEmpleado Where (((Empleado.CodEmpleado) = '" & DBCodigo.Text & "') And ((Salida.Fecha) = " & Fecha & "))"
DtaBusca.Refresh

If DtaBusca.Recordset.EOF Then
            TxtEntra.Text = Format(Now, "long time")
            TxtEntra.ForeColor = &H80&
            TxtSalida.ForeColor = &H80000008
            TxtSalida.Text = ""
            Quien = "Entrada"
        ElseIf Not IsNull(DtaBusca.Recordset("HoraEntra")) And Not IsNull(DtaBusca.Recordset("horasale")) Then
            MsgBox "Este empleado ya tiene grabadas las horas de entrada y salida del día de hoy)"
            DBCodigo.Text = ""
        Else
            TxtEntra.Text = Format(DtaBusca.Recordset("HoraEntra"), "long time")
            TxtSalida.Text = Format(Now, "long time")
            TxtSalida.ForeColor = &H80&
            TxtEntra.ForeColor = &H80000008
            Quien = ""
End If
Else
   TxtCargo.Text = ""
   txtNombre1.Text = ""
   txtApellido1.Text = ""
   TxtCargo.Text = "salida"
End If
End Sub



Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub mscReloj_OnComm()
Dim Longitud As Integer
Dim Cadena As String


 Select Case Me.mscReloj.CommEvent

'   Case CommBreak


   Case 2

       Cadena = Me.mscReloj.Input
       Longitud = Len(Cadena)
       Me.DBCodigo.Text = (Mid(Cadena, 1, Longitud - 2))
       Me.LblPuerto.Caption = "Registro del Reloj #1"

 End Select



End Sub

Private Sub mscReloj2_OnComm()

Dim Longitud As Integer
Dim Cadena As String




 Select Case Me.mscReloj2.CommEvent
'
'   Case CommBreak


   Case 2

       Cadena = Me.mscReloj2.Input
       Longitud = Len(Cadena)
       Me.DBCodigo.Text = (Mid(Cadena, 1, Longitud - 2))
       Me.LblPuerto.Caption = "Registro del Reloj #2"

 End Select


End Sub

Private Sub Form_Activate()
lblFecha.Caption = Format(Now, "Long Date")
End Sub
Private Sub Form_Resize()
    If WindowState = conMinimized Then      ' Si el formulario se minimiza, presenta la hora en un título.
        SetCaptionTime
    End If
End Sub

Private Sub SetCaptionTime()
    
    Caption = Format(Time, "Medium Time")   ' Presenta la hora con el formato Medium Time.
End Sub

Private Sub Form_Load()
On Error GoTo TipoErrs
'Me.DtaBusca.DatabaseName = Ruta
'Me.DtaEmpleado.DatabaseName = Ruta
'Me.DtaSalida.DatabaseName = Ruta
With Me.DtaBusca
 .ConnectionString = Conexion
End With

With Me.DtaEmpleado
   .ConnectionString = Conexion
   .RecordSource = "Empleado"
   .Refresh
End With

With Me.DtaSalida
   .ConnectionString = Conexion
   
End With
Me.Top = 2500
Me.Left = 2500
  AlarmTime = ""


Exit Sub
TipoErrs:
If Not Err.Number = 8002 Then
 MsgBox Err.Description
End If

End Sub
Public Sub Mes(Mes As Variant)
Select Case Mes
  Case 1
      Mes = "de Enero"
  Case 2
      Mes = "de Febrero"
  Case 3
      Mes = "de Marzo"
  Case 4
      Mes = "de Abril"
  Case 5
     Mes = "de Mayo"
  Case 6
     Mes = "de Junio"
  Case 7
     Mes = "de Julio"
  Case 8
      Mes = "de Agosto"
  Case 9
      Mes = "de Septiembre"
  Case 10
      Mes = "de Octubre"
  Case 11
     Mes = "de Noviembre"
  Case 12
     Mes = "de Diciembre"
End Select

End Sub

Private Sub Timer1_Timer()
Static AlarmSounded As Integer
  
If Quien = "Entrada" Then
 If TxtSalida.Text = "" Then
  If TxtEntra.ForeColor = &H800000 Then 'Color Azul
     TxtEntra.ForeColor = &HFF& 'Color Rojo
  Else
     TxtEntra.ForeColor = &H800000
  End If
  
 End If
 lblTime.ForeColor = &H800000    'Color Azul
  AlarmTiempo = "8:00 am"
  AlarmTime = "8:10 am"
  AlarmTime = CDate(AlarmTime)
  AlarmTiempo = CDate(AlarmTiempo)
  
    If lblTime.Caption <> CStr(Time) Then
        ' Ahora es un segundo diferente del presentado.
        If Time < AlarmTiempo Then
            
            
            If lblMensaje.ForeColor = &H8000& Then
             lblMensaje.Caption = "!!Bien!!"
             lblMensaje.ForeColor = &H80000000 'Color del Formulario
         Else
             lblMensaje.Caption = "!!Bien!!"
             lblMensaje.ForeColor = &H8000&        'Color Verde
         End If
         
        End If
        If Time >= AlarmTiempo And Time <= AlarmTime Then 'And Not AlarmSounded Then
            
            
            If lblMensaje.ForeColor = &H800000 Then
             lblMensaje.Caption = "!!Alarma!!"
             lblMensaje.ForeColor = &H80000000 'Color del Formulario
            Else
             lblMensaje.Caption = "!!Alarma!!"
             lblMensaje.ForeColor = &H800000    'Color Azul
            End If
            AlarmSounded = True
        ElseIf Time > AlarmTime Then
            lblTime.ForeColor = &HFF& 'Color Rojo
           If lblMensaje.ForeColor = &HFF& Then
             lblMensaje.Caption = "!!Tarde!!"
             lblMensaje.ForeColor = &H80000000
            Else
             lblMensaje.Caption = "!!Tarde!!"
             lblMensaje.ForeColor = &HFF&  'Color Rojo
            End If
            AlarmSounded = False
        End If
        If WindowState = conMinimized Then
            ' Si está minimizado, actualiza el título del formulario cada minuto.
            If Minute(CDate(Caption)) <> Minute(Time) Then SetCaptionTime
        Else
            ' Si no, actualiza la etiqueta del formulario cada segundo.
            lblTime.Caption = Time
        End If
    End If
Else
If TxtSalida.ForeColor = &H800000 Then 'Color Azul
     TxtSalida.ForeColor = &HFF& 'Color Rojo
  Else
     TxtSalida.ForeColor = &H800000
  End If
 If lblMensaje.ForeColor = &H800000 Then
             lblMensaje.Caption = "Salida"
             lblMensaje.ForeColor = &H80000000 'Color del Formulario
            Else
             lblMensaje.Caption = "Salida"
             lblMensaje.ForeColor = &H800000    'Color Azul
            End If
lblMensaje.Caption = ""
lblMensaje.ForeColor = &H800000    'Color Azul
lblTime.ForeColor = &H800000    'Color Azul
If WindowState = conMinimized Then
            ' Si está minimizado, actualiza el título del formulario cada minuto.
            If Minute(CDate(Caption)) <> Minute(Time) Then SetCaptionTime
        Else
            ' Si no, actualiza la etiqueta del formulario cada segundo.
            lblTime.Caption = Time
        End If
    
End If
End Sub
