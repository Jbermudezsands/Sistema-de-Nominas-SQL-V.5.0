VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmRespaldar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Respaldar Base de Datos"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   HelpContextID   =   120000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoPassword 
      Height          =   330
      Left            =   2280
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
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
      Caption         =   "AdoPassword"
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   970
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6735
      TabIndex        =   7
      Top             =   0
      Width           =   6735
      Begin VB.Image Image2 
         Height          =   645
         Left            =   480
         Picture         =   "FrmRes.frx":0000
         Top             =   120
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   6720
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Top             =   120
         Width           =   645
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Respaldar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   2520
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Respaldo"
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   2040
      Width           =   4335
      Begin VB.OptionButton OptCompleto 
         Caption         =   "Completo"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptDiferencial 
         Caption         =   "Diferencial"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informacion"
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4935
      Begin VB.CommandButton cmdcerrar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdBackup 
         Caption         =   "Respaldar"
         Height          =   375
         Left            =   2160
         TabIndex        =   0
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtruta 
         Height          =   288
         Left            =   360
         TabIndex        =   3
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox txtbd 
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "SistemaVentas"
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Examinar"
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   1800
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4320
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Data File Name:"
         Filter          =   "*.bkp"
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   0
         OleObjectBlob   =   "FrmRes.frx":166E
         Top             =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmRes.frx":25CE9B
         TabIndex        =   9
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmRes.frx":25CF1B
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmRespaldar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConexionBackup As New ADODB.Connection
Dim rstBD As ADODB.Recordset
Dim Contraseña As String

Sub Seleccionar(TextBox As TextBox)
    TextBox.SetFocus
    TextBox.SelStart = 0
    TextBox.SelLength = Len(TextBox)
End Sub

Private Sub cmdBackup_Click()
On Error GoTo error
Dim Longitud As Double
Dim Directorio As String

If txtbd.Text = "" Then
    MsgBox "Debe ingresar la base de datos a respaldar", vbExclamation
    Exit Sub
End If



 Directorio = ""
 Longitud = Len(Me.CommonDialog1.FileName)
 Directorio = Mid(Me.CommonDialog1.FileName, 1, Longitud - 4)
 
 
If Me.OptCompleto.Value Then
 Me.TxtRuta.Text = Directorio & " Full " & Format(Now, "dd-mm-yyyy") & ".bkp"
'    Me.txtruta.Text = App.Path & "\Respaldos\" & Me.txtbd.Text & "Full" & Format(Now, "ddmmyyyy-hh-mm-ss") & ".bkp"
Else
'    Me.txtruta.Text = App.Path & "\Respaldos\" & Me.txtbd.Text & "Dif" & Format(Now, "ddmmmyyyy-hh-mm-ss") & ".bkp"
 Me.TxtRuta.Text = Directorio & "Dif" & Format(Now, "dd-mm-yyyy") & ".bkp"
End If

If TxtRuta.Text = "" Then
    MsgBox "Debe indicar la ruta donde guardara el respaldo", vbExclamation
    TxtRuta.SetFocus
Else

    If OptCompleto.Value Then
        ConexionBackup.Execute "Backup DATABASE [" + txtbd.Text + "] TO DISK='" & TxtRuta & "'"
    ElseIf OptDiferencial.Value Then
        ConexionBackup.Execute "Backup DATABASE [" + txtbd.Text + "] TO DISK='" & TxtRuta & "' with DIFFERENTIAL"
    End If
    Debug.Print Me.TxtRuta.Text
    MsgBox "Base de datos Respaldada con exito", vbInformation
    Unload Me
End If

error:
If Err.Number <> 0 Then
    MsgBox Err.Description '"Ha ocurrido un error al momento de intentar realizar el respaldo", vbInformation
End If
'Exit Sub

End Sub

Private Sub cmdBrowse_Click()
 On Error GoTo errHandler:
     Me.CommonDialog1.FileName = Me.txtbd.Text
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "All Files (*.*)|*.*|Backup Files (*.bak)|*.bak"
    CommonDialog1.DefaultExt = "bak"
    CommonDialog1.DialogTitle = "Nombre del Respaldo"
    Me.CommonDialog1.ShowSave
    TxtRuta.Text = CommonDialog1.FileName
'    CommonDialog1.Action = 0
 

    
    Exit Sub
    
errHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub CmdCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.ApplySkin Me.hWnd
Dim NextLine As String, Cadena As Variant
Dim posicion As Long
Dim Servidor As String
Dim ConexionBackupSTR1 As String
Dim StrCn As String
Dim BaseDatos As String

'Open App.Path + "\SysInfo.dll" For Input As #1
'i = 1
' Do Until EOF(1)
'    Line Input #1, NextLine
'    If i = 1 Then
'        ConexionBackupSTR1 = Trim(NextLine)
'    Else
'        ConexionBackupSTR2 = Trim(NextLine)
'    End If
'    i = i + 1
' Loop
'Close #1
ConexionBackupSTR1 = Conexion
posicion = InStr(1, ConexionBackupSTR1, "Source")
Servidor = Mid(ConexionBackupSTR1, posicion + 7, Len(ConexionBackupSTR1) - posicion + 6)

Cadena = UCase(Conexion)

'Debug.Print Conexion
Dim inicioBD As Integer, FinBD As Integer
inicioBD = InStr(UCase(Conexion), UCase("Initial Catalog="))
inicioBD = inicioBD + 16

FinBD = InStr(UCase(Conexion), UCase(";Data Source="))

If FinBD <> 0 Then
  Me.txtbd.Text = Mid(Conexion, inicioBD, FinBD - inicioBD)
Else
 FinBD = Len(UCase(Conexion)) + 1
 Me.txtbd.Text = Mid(Conexion, inicioBD, FinBD - inicioBD)
End If

StrCn = "Driver=SQL SERVER;UID=ADMINISTRADOR;SERVER=" & Servidor & ";DATABASE=MASTER;TRUSTED_CONNECTION=YES; APP=VIRUS;WID=SISTEMAS"
StrCn = Conexion 'temporal
ConexionBackup.ConnectionString = StrCn
Debug.Print StrCn
On Error Resume Next
ConexionBackup.Open
Set rstBD = New ADODB.Recordset
rstBD.Open "select name from sysdatabases", ConexionBackup, adOpenDynamic, adLockOptimistic

'Me.AdoPassword.ConnectionString = Conexion
'Me.AdoPassword.RecordSource = "select * from configuracion"
'Me.AdoPassword.Refresh
'Contraseña = AdoPassword.Recordset!PasswordBackup

Me.TxtRuta.Text = App.Path & "\Respaldos\" & Me.txtbd.Text & Format(Date, "ddmmyyyy") & ".bkp"
Me.TxtRuta.Locked = True
Me.OptDiferencial.Value = 1
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtcon.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ConexionBackup = Nothing
    Set rstBD = Nothing
End Sub

