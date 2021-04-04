VERSION 5.00
Object = "{EAD61168-CF37-11D1-A050-70D904C10000}#2.0#0"; "MacWin.ocx"
Begin VB.Form FrmAcerca 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Acerca de Zeus Nóminas"
   ClientHeight    =   5460
   ClientLeft      =   2295
   ClientTop       =   1620
   ClientWidth     =   7200
   ClipControls    =   0   'False
   DrawStyle       =   6  'Inside Solid
   ForeColor       =   &H00FFEFEF&
   Icon            =   "FrmAcerca.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3768.589
   ScaleMode       =   0  'User
   ScaleWidth      =   6761.172
   ShowInTaskbar   =   0   'False
   Begin MacWindow.MacWin MacWin1 
      Height          =   300
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   529
      Caption         =   "Informacion del Sistema "
      Blue            =   20
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   3840
      Top             =   2160
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      DownPicture     =   "FrmAcerca.frx":030A
      Height          =   345
      Left            =   5040
      Picture         =   "FrmAcerca.frx":1DEC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      DownPicture     =   "FrmAcerca.frx":38CE
      Height          =   345
      Left            =   5040
      Picture         =   "FrmAcerca.frx":53B0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nominas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   6
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   6
      Left            =   120
      Picture         =   "FrmAcerca.frx":6CB2
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   5
      Left            =   240
      Picture         =   "FrmAcerca.frx":70F4
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   4
      Left            =   360
      Picture         =   "FrmAcerca.frx":7536
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   3
      Left            =   120
      Picture         =   "FrmAcerca.frx":7978
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   2
      Left            =   120
      Picture         =   "FrmAcerca.frx":7DBA
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   1
      Left            =   120
      Picture         =   "FrmAcerca.frx":81FC
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   197.201
      X2              =   6197.741
      Y1              =   2650.436
      Y2              =   2650.436
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmAcerca.frx":863E
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   6165
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Nominas"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   211.287
      X2              =   6085.055
      Y1              =   2733.262
      Y2              =   2733.262
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Versión 1.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4440
      TabIndex        =   5
      Top             =   2160
      Width           =   1485
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmAcerca.frx":8765
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   945
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   0
      Left            =   120
      Picture         =   "FrmAcerca.frx":881B
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "FrmAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Imagen As Integer
Option Explicit

' Opciones de seguridad de clave del Registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos ROOT de clave del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cadena Unicode terminada en valor nulo
Const REG_DWORD = 4                      ' Número de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
'    Me.Caption = "Acerca de " & App.Title
 '   lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
 '   lblTitle.Caption = App.Title
 '   Imagen = 0
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Intentar obtener ruta de acceso y nombre del programa de Info. del sistema a partir del Registro...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Intentar obtener sólo ruta del programa de Info. del sistema a partir del Registro...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validar la existencia de versión conocida de 32 bits del archivo
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error: no se puede encontrar el archivo...
        Else
            GoTo SysInfoErr
        End If
    ' Error: no se puede encontrar la entrada del Registro...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "La información del sistema no está disponible en este momento", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contador de bucle
    Dim rc As Long                                          ' Código de retorno
    Dim hKey As Long                                        ' Controlador de una clave de Registro abierta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo de datos de una clave de Registro
    Dim tmpVal As String                                    ' Almacenamiento temporal para un valor de clave de Registro
    Dim KeyValSize As Long                                  ' Tamaño de variable de clave de Registro
    '------------------------------------------------------------
    ' Abrir clave de registro bajo KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abrir clave de Registro
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Error de controlador...
    
    tmpVal = String$(1024, 0)                             ' Asignar espacio de variable
    KeyValSize = 1024                                       ' Marcar tamaño de variable
    
    '------------------------------------------------------------
    ' Obtener valor de clave de Registro...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtener o crear valor de clave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar errores
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agregar cadena terminada en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Encontrado valor nulo, se va a quitar de la cadena
    Else                                                    ' En WinNT las cadenas no terminan en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize)                   ' No se ha encontrado valor nulo, sólo se va a extraer la cadena
    End If
    '------------------------------------------------------------
    ' Determinar tipo de valor de clave para conversión...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Buscar tipos de datos...
    Case REG_SZ                                             ' Tipo de datos String de clave de Registro
        KeyVal = tmpVal                                     ' Copiar valor de cadena
    Case REG_DWORD                                          ' Tipo de datos Double Word de clave del Registro
        For i = Len(tmpVal) To 1 Step 1                    ' Convertir cada bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Generar valor carácter a carácter
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word a cadena
    End Select
    
    GetKeyValue = True                                      ' Se ha devuelto correctamente
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
    Exit Function                                           ' Salir
    
GetKeyError:      ' Borrar después de que se produzca un error...
    KeyVal = ""                                             ' Establecer valor a cadena vacía
    GetKeyValue = False                                     ' Fallo de retorno
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
End Function

Private Sub Timer1_Timer()
If Imagen = 6 Then
   Image1(Imagen).Visible = True
   Image1(Imagen - 1).Visible = False
   Imagen = 0
ElseIf Imagen = 0 Then
   Image1(Imagen).Visible = True
   Image1(6).Visible = False
   Imagen = Imagen + 1
Else
   Image1(Imagen).Visible = True
   Image1(Imagen - 1).Visible = False
   Imagen = Imagen + 1
End If

End Sub
