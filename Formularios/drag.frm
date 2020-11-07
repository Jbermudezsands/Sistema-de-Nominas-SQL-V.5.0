VERSION 5.00
Begin VB.Form frmDrag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Imagenes"
   ClientHeight    =   3435
   ClientLeft      =   2085
   ClientTop       =   3840
   ClientWidth     =   7965
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "drag.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   531
   Begin VB.CommandButton CmdSalir 
      DownPicture     =   "drag.frx":0442
      Height          =   375
      Left            =   6240
      MouseIcon       =   "drag.frx":1F24
      MousePointer    =   99  'Custom
      Picture         =   "drag.frx":2366
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton CmdPegar 
      DownPicture     =   "drag.frx":3E48
      Height          =   375
      Left            =   4680
      MouseIcon       =   "drag.frx":592A
      MousePointer    =   99  'Custom
      Picture         =   "drag.frx":5D6C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      DragIcon        =   "drag.frx":784E
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   5520
      Pattern         =   "*.bmp; *.gif;*.jpg"
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      DragIcon        =   "drag.frx":7B58
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   3480
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Image PictFotoempleado 
      BorderStyle     =   1  'Fixed Single
      Height          =   2895
      Left            =   120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdPegar_Click()
If frmEmpleado.DBCodigoEmpleado.Text = "" Then
    MsgBox "Para Agregar Foto se Necesita el Codigo del Empleado", vbInformation, "Error:Sistema de Nominas"
    frmEmpleado.DBCodigoEmpleado.SetFocus
    Exit Sub
  End If
CodEmpleado = frmEmpleado.DBCodigoEmpleado.Text
 ' Obtiene las tres últimas letras del nombre del archivo arrastrado.
    temp = Right$(File1.FileName, 3)

    ' Si el archivo arrastrado se encuentra en la raíz, agrega el nombre del archivo.
    If Mid$(File1.Path, Len(File1.Path)) = "\" Then
      dropfile = File1.Path & File1.FileName
    ' Si el archivo arrastrado no se encuentra en la raíz, agrega "\" al nombre del archivo.
    Else
      dropfile = File1.Path & "\" & File1.FileName
    End If
    Guarda = frmEmpleado.DBCodigoEmpleado.Text + "." + temp
    'InputBox("Digite el nombre")
    Origen = frmDrag.File1.Path & "\" & frmDrag.File1.FileName
       
    Destino = RutaFoto & Guarda
    frmEmpleado.Image1.Picture = LoadPicture("")
    Select Case UCase$(Trim$(temp))
         Case "BMP"
            frmEmpleado.Image1.Picture = LoadPicture(dropfile)
            'MsgBox Origen
            'MsgBox Destino
            If Dir(Destino) <> "" Then
                Kill RutaFoto + CodEmpleado + "*"
            End If
            FileCopy Origen, Destino
         Case "JPG"
            frmEmpleado.Image1.Picture = LoadPicture(dropfile)
            'MsgBox Origen
            'MsgBox Destino
            If Dir(Destino) <> "" Then
                Kill RutaFoto + CodEmpleado + "*"
            End If
            FileCopy Origen, Destino
         Case "GIF"
            frmEmpleado.Image1.Picture = LoadPicture(dropfile)
            'MsgBox Origen
            'MsgBox Destino
            If Dir(Destino) <> "" Then
                Kill RutaFoto + CodEmpleado + "*"
            End If
            FileCopy Origen, Destino
        Case Else
            MsgBox "SOLO SE PUEDE ARRASTRAR UN ARCHIVO DE TIPO NO GRÁFICO"
    End Select
    CodEmpleado = frmEmpleado.DBCodigoEmpleado.Text
    MsgBox "La foto del Empleado será Cambiada"
Unload Me
End Sub

Private Sub CmdSalir_Click()
 Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo DriveErrs
    Dir1.Path = Drive1.Drive
Exit Sub
DriveErrs:
   ControlErrores
End Sub

Private Sub File1_Click()
On Error GoTo TipoErr
Dim RutaTemp As String
If Len(Dir1.Path) > 3 Then
    RutaTemp = Dir1.Path + "\" + File1.FileName
Else
    RutaTemp = Dir1.Path + File1.FileName
End If
PictFotoempleado.Picture = LoadPicture(RutaTemp)
Exit Sub
TipoErr:
ControlErrores
End Sub

Private Sub File1_DblClick()
Dim RutaTemp As String
'coloco la foto en la imagen
If Len(Dir1.Path) > 3 Then
    RutaTemp = Dir1.Path + "\" + File1.FileName
Else
    RutaTemp = Dir1.Path + File1.FileName
End If
frmEmpleado.Image1.Picture = LoadPicture(RutaTemp)
'copio el archivo a la ruta foto
 temp = Right$(File1.FileName, 3)
 Guarda = frmEmpleado.DBCodigoEmpleado.Text + "." + temp

Destino = RutaFoto & Guarda

FileCopy RutaTemp, Destino
 CodEmpleado = frmEmpleado.DBCodigoEmpleado.Text
Unload Me
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    File1.DragIcon = Drive1.DragIcon
    File1.Drag

End Sub

Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
If frmEmpleado.DBCodigoEmpleado.Text = "" Then
    MsgBox "Para Agregar Foto se Necesita el Codigo del Empleado", vbInformation, "Error:Sistema de Nominas"
    frmEmpleado.DBCodigoEmpleado.SetFocus
    Exit Sub
  End If
CodEmpleado = frmEmpleado.DBCodigoEmpleado.Text
 ' Obtiene las tres últimas letras del nombre del archivo arrastrado.
    temp = Right$(File1.FileName, 3)

    ' Si el archivo arrastrado se encuentra en la raíz, agrega el nombre del archivo.
    If Mid$(File1.Path, Len(File1.Path)) = "\" Then
      dropfile = File1.Path & File1.FileName
    ' Si el archivo arrastrado no se encuentra en la raíz, agrega "\" al nombre del archivo.
    Else
      dropfile = File1.Path & "\" & File1.FileName
    End If
    Guarda = frmEmpleado.DBCodigoEmpleado.Text + "." + temp
    'InputBox("Digite el nombre")
    Origen = frmDrag.File1.Path & "\" & frmDrag.File1.FileName
       
    Destino = RutaFoto & Guarda
    frmEmpleado.Image1.Picture = LoadPicture("")
    Select Case UCase$(Trim$(temp))
         Case "BMP"
            frmEmpleado.Image1.Picture = LoadPicture(dropfile)
            'MsgBox Origen
            'MsgBox Destino
            If Dir(Destino) <> "" Then
                Kill RutaFoto + CodEmpleado + "*"
            End If
            FileCopy Origen, Destino
         Case "JPG"
            frmEmpleado.Image1.Picture = LoadPicture(dropfile)
            'MsgBox Origen
            'MsgBox Destino
            If Dir(Destino) <> "" Then
                Kill RutaFoto + CodEmpleado + "*"
            End If
            FileCopy Origen, Destino
         Case "GIF"
            frmEmpleado.Image1.Picture = LoadPicture(dropfile)
            'MsgBox Origen
            'MsgBox Destino
            If Dir(Destino) <> "" Then
                Kill RutaFoto + CodEmpleado + "*"
            End If
            FileCopy Origen, Destino
        Case Else
            MsgBox "SOLO SE PUEDE ARRASTRAR UN ARCHIVO DE TIPO NO GRÁFICO"
    End Select
    CodEmpleado = frmEmpleado.DBCodigoEmpleado.Text
    MsgBox "La foto del Empleado será Cambiada"
Unload Me
End Sub

Private Sub Image1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Select Case State
    Case 0
        ' Presenta un icono nuevo cuando el origen entra en el área de colocar.
        File1.DragIcon = Dir1.DragIcon
    Case 1
        ' Presenta el DragIcon original cuando el origen sale del área de colocar.
        File1.DragIcon = Drive1.DragIcon
    End Select

' Observe que Dir1.DragIcon y Drive1.DragIcon han sido establecidos
' en tiempo de diseño. Esto le permite cargar los iconos "Enter"
' y "Leave" de File1 en tiempo de ejecución sin que sea necesario
' que el usuario los tenga en el disco.
End Sub

Private Sub OLE1_Updated(Code As Integer)

End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub

