VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmEditaNiveles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar Niveles"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6525
   DrawStyle       =   1  'Dash
   HelpContextID   =   5
   Icon            =   "FrmEditaNiveles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   435
   Begin VB.TextBox TxtCodpasword 
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Text            =   "TxtCodpasword"
      Top             =   5160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox TxtNivel 
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Text            =   "TxtNivel"
      Top             =   5160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox ListAcceso 
      Height          =   1620
      ItemData        =   "FrmEditaNiveles.frx":0442
      Left            =   3240
      List            =   "FrmEditaNiveles.frx":0444
      TabIndex        =   9
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton CmdAceptar 
      DownPicture     =   "FrmEditaNiveles.frx":0446
      Height          =   375
      Left            =   3360
      MouseIcon       =   "FrmEditaNiveles.frx":1F28
      MousePointer    =   99  'Custom
      Picture         =   "FrmEditaNiveles.frx":236A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancelar 
      DownPicture     =   "FrmEditaNiveles.frx":3E4C
      Height          =   375
      Left            =   4800
      MouseIcon       =   "FrmEditaNiveles.frx":592E
      MousePointer    =   99  'Custom
      Picture         =   "FrmEditaNiveles.frx":5D70
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   2040
      TabIndex        =   5
      Top             =   3480
      Width           =   3975
      Begin VB.Label LblNombre 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Permisos"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   5895
      Begin VB.CheckBox ChEliminar 
         Caption         =   "Eliminar Datos"
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox ChGrabar 
         Caption         =   "Grabar Datos"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox ChAbrir 
         Caption         =   "Abrir o Ejecutar"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc DtaNacceso 
      Height          =   375
      Left            =   120
      Top             =   7320
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
   Begin MSAdodcLib.Adodc DtaPasword2 
      Height          =   375
      Left            =   120
      Top             =   6960
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
      Caption         =   "DtaPasword2"
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
   Begin MSAdodcLib.Adodc DtaPasword 
      Height          =   375
      Left            =   120
      Top             =   6600
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "DtaPasword"
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
   Begin MSDataListLib.DataList DBLNEmpleado 
      Bindings        =   "FrmEditaNiveles.frx":7852
      Height          =   1620
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2858
      _Version        =   393216
      ListField       =   "NombreUsuario"
   End
   Begin VB.Label Label3 
      Caption         =   "Usuario Seleccionado:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Ventanas de Zeus Nóminas"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Usuarios o Grupos:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "FrmEditaNiveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CmdAceptar_Click()
On Error GoTo TipoErrs

'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
     If NivelAcceso < Val(TxtNivel.Text) Then
        MsgBox "Imposible Modificar su Nivel es Inferior", vbExclamation, "Sistema de Nominas"
        Exit Sub
      End If
      
  Select Case Me.ListAcceso.Text
        Case "Editar Niveles"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Editar Niveles'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Editar Niveles"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Editar Niveles'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
        Case "Registro Empleados"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Registro Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Registro Empleados"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Registro Empleados'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Registro Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Registro Empleados"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Registro Empleados'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Registro Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Registro Empleados"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Registro Empleados'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
        Case "Tabla Anotaciones"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Anotaciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tabla Anotaciones"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Anotaciones'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Anotaciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tabla Anotaciones"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Anotaciones'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Anotaciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tabla Anotaciones"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Anotaciones'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
        Case "Tabla Departamentos"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Departamentos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tabla Departamentos"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Departamentos'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Departamentos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tabla Departamentos"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Departamentos'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Departamentos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tabla Departamentos"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Departamentos'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
        Case "Tabla Cargo"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Cargo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tabla Cargo"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Cargo'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Cargo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tabla Cargo"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Cargo'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Cargo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tabla Cargo"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Cargo'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
        Case "Tabla Incapacidad"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tabla Incapacidad"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tabla Incapacidad"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tabla Incapacidad"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
 
         Case "Tabla Incentivos"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Incentivos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tabla Incentivos"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Incentivos'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Incentivos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tabla Incentivos"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Incentivos'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Incentivos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tabla Incentivos"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Incentivos'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
    Case "Tabla Tipo Incapacidad"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tabla Tipo Incapacidad"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tabla Tipo Incapacidad"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tabla Tipo Incapacidad"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
 
     Case "Tabla Tipo Deducciones"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Deducciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tabla Tipo Deducciones"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Deducciones'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Deducciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tabla Tipo Deducciones"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Deducciones'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Deducciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tabla Tipo Deducciones"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Deducciones'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
 
      Case "Tabla Tipo Subsidio"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Subsidio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tabla Tipo Subsidio"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Subsidio'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Subsidio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tabla Tipo Subsidio"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Subsidio'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Subsidio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tabla Tipo Subsidio"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Subsidio'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
 
       Case "Tabla Tipo Comision"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Comision'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tabla Tipo Comision"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Comision'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Comision'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tabla Tipo Comision"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Comision'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Comision'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tabla Tipo Comision"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Comision'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       Case "Tabla Tipo Destajo"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Destajo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tabla Tipo Destajo"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Destajo'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Destajo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tabla Tipo Destajo"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Destajo'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Destajo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tabla Tipo Destajo"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Destajo'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       Case "Tabla Tipo Nomina"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Nomina'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tabla Tipo Nomina"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Nomina'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Nomina'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tabla Tipo Nomina"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Nomina'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Nomina'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tabla Tipo Nomina"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Nomina'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       Case "Tabla Tipo INSS/IR"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo INSS/IR'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tabla Tipo INSS/IR"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo INSS/IR'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo INSS/IR'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Tabla Tipo INSS/IR"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo INSS/IR'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo INSS/IR'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Tabla Tipo INSS/IR"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo INSS/IR'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       Case "Tabla Listado Nominas"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Listado Nominas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Tabla Listado Nominas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Listado Nominas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       Case "Calcular Nominas"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Calcular Nominas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Calcular Nominas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Calcular Nominas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       Case "Activar Nominas"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Activar Nominas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Activar Nominas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Activar Nominas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       Case "Movimientos de la Nomina"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Movimientos de la Nomina'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Movimientos de la Nomina"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Movimientos de la Nomina'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Movimientos de la Nomina'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Movimientos de la Nomina"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Movimientos de la Nomina'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       Case "Calcular 13vo y Vacaciones"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Calcular 13vo y Vacaciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Calcular 13vo y Vacaciones"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Calcular 13vo y Vacaciones'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       Case "Nomina de Subsidio"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Nomina de Subsidio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Nomina de Subsidio"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Nomina de Subsidio'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       Case "Despido y Renuncias"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Despido y Renuncias'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Despido y Renuncias"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Despido y Renuncias'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       Case "Entradas y Salidas"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Entradas y Salidas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Entradas y Salidas"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Entradas y Salidas'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If

       Case "Calcular Horas Extra o Falta"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Calcular Horas Extra o Falta'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Calcular Horas Extra o Falta"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Calcular Horas Extra o Falta'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       Case "Reportes"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Reportes"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
 
        Case "Controles Personalizados"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Controles Personalizados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Controles Personalizados"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Controles Personalizados'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       Case "Usuarios"
       '//////////////Verifico el Chabrir//////////////////
        If Me.ChAbrir.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Abrir Usuarios"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Usuarios'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChGrabar//////////////////
        If Me.ChGrabar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Grabar Usuarios"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
        
       '//////////////Verifico el ChBorrar//////////////////
        If Me.ChEliminar.Value = 1 Then
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.AddNew
             DtaNacceso.Recordset("CodUsuario") = CodigoUsuario
             DtaNacceso.Recordset("AccesoModulo") = "Borrar Usuarios"
           Me.DtaNacceso.Recordset.Update
         End If
        Else
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Not Me.DtaNacceso.Recordset.EOF Then
           Me.DtaNacceso.Recordset.Delete
         End If
        End If
 
   End Select
      
Exit Sub
TipoErrs:
 ControlErrores
Unload Me
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub
Private Sub DBLNEmpleado_Click()
'On Error GoTo TipoErrs
LblNombre.Caption = DBLNEmpleado.Text
DtaPasword2.Refresh
      Do While Not DtaPasword2.Recordset.EOF
       If DtaPasword2.Recordset("NombreUsuario") = DBLNEmpleado.Text Then
         TxtCodpasword.Text = DtaPasword2.Recordset("CodUsuario")
         TxtNivel.Text = DtaPasword2.Recordset("NivelAcceso")
         Exit Do
       End If
        DtaPasword2.Recordset.MoveNext
      Loop
 
 
      
'Exit Sub
'TipoErrs:
' ControlErrores
' Unload Me
End Sub

Private Sub Form_Activate()
DtaPasword.Refresh
'DBGrid1.Columns.Add 3
'DBGrid1.Columns(3).Visible = True
'DBGrid1.Columns(0).Visible = False
'DBGrid1.Columns(2).Visible = False
'DBGrid1.Columns(1).Width = TextWidth("                              ")
'DBGrid1.Columns(3).Width = TextWidth("Respuesta  ")
'DBGrid1.Columns(3).Caption = "Respuesta"
'DBGrid1.Columns(1).Locked = True
'DBGrid1.Columns(3).Locked = False
End Sub

Private Sub Form_Load()
With Me.DtaNacceso
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaPasword
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Usuarios"
   .Refresh
End With

With Me.DtaPasword2
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Usuarios"
   .Refresh
End With

Me.ListAcceso.AddItem ("Editar Niveles")
Me.ListAcceso.AddItem ("Registro Empleados")
Me.ListAcceso.AddItem ("Tabla Anotaciones")
Me.ListAcceso.AddItem ("Tabla Departamentos")
Me.ListAcceso.AddItem ("Tabla Cargo")
Me.ListAcceso.AddItem ("Tabla Incapacidad")
Me.ListAcceso.AddItem ("Tabla Incentivos")
Me.ListAcceso.AddItem ("Tabla Tipo Incapacidad")
Me.ListAcceso.AddItem ("Tabla Tipo Deducciones")
Me.ListAcceso.AddItem ("Tabla Tipo Subsidio")
Me.ListAcceso.AddItem ("Tabla Tipo Comision")
Me.ListAcceso.AddItem ("Tabla Tipo Destajo")
Me.ListAcceso.AddItem ("Tabla Tipo Nomina")
Me.ListAcceso.AddItem ("Tabla Tipo INSS/IR")
Me.ListAcceso.AddItem ("Tabla Listado Nominas")
Me.ListAcceso.AddItem ("Calcular Nominas")
Me.ListAcceso.AddItem ("Activar Nominas")
Me.ListAcceso.AddItem ("Movimientos de la Nomina")
Me.ListAcceso.AddItem ("Calcular 13vo y Vacaciones")
Me.ListAcceso.AddItem ("Nomina de Subsidio")
Me.ListAcceso.AddItem ("Despido y Renuncias")
Me.ListAcceso.AddItem ("Entradas y Salidas")
Me.ListAcceso.AddItem ("Calcular Horas Extra o Falta")
Me.ListAcceso.AddItem ("Reportes")
Me.ListAcceso.AddItem ("Controles Personalizados")
Me.ListAcceso.AddItem ("Usuarios")


End Sub

Private Sub ListAcceso_Click()
On Error GoTo TipoErrs


    Select Case Me.ListAcceso.Text
      
        Case "Editar Niveles"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Editar Niveles'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
          Me.ChEliminar.Value = 0
          Me.ChGrabar.Value = 0
         Else
          Me.ChAbrir.Value = 1
          Me.ChEliminar.Value = 0
          Me.ChGrabar.Value = 0
         End If
         
        Case "Registro Empleados"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Registro Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Registro Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Registro Empleados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
         
        Case "Tabla Anotaciones"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Anotaciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Anotaciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Anotaciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If

        Case "Tabla Departamentos"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Departamentos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Departamentos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Departamentos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
         
        Case "Tabla Cargo"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Cargo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Cargo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Cargo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
         
        Case "Tabla Incapacidad"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
         
        Case "Tabla Incentivos"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Incentivos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Incentivos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Incentivos'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
         
        Case "Tabla Tipo Incapacidad"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Incapacidad'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
         
        Case "Tabla Tipo Deducciones"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Deducciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Deducciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Deducciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
         
        Case "Tabla Tipo Subsidio"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Subsidio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Subsidio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Subsidio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
         
        Case "Tabla Tipo Comision"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Comision'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Comision'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Comision'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
         
        Case "Tabla Tipo Destajo"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Destajo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Destajo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Destajo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
         
        Case "Tabla Tipo Nomina"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo Nomina'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo Destajo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo Destajo'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
         
        Case "Tabla Tipo INSS/IR"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Tipo INSS/IR'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Tabla Tipo INSS/IR'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Tabla Tipo INSS/IR'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
         
        Case "Tabla Listado Nominas"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Tabla Listado Nominas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
   
          Me.ChGrabar.Value = 0
          Me.ChEliminar.Value = 0
          
        Case "Calcular Nominas"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Calcular Nominas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
   
          Me.ChGrabar.Value = 0
          Me.ChEliminar.Value = 0

        Case "Activar Nominas"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Activar Nominas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
   
          Me.ChGrabar.Value = 0
          Me.ChEliminar.Value = 0
         
         Case "Movimientos de la Nomina"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = False

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Movimientos de la Nomina'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
          '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Movimientos de la Nomina'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If

          Me.ChEliminar.Value = 0
          
        Case "Calcular 13vo y Vacaciones"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Calcular 13vo y Vacaciones'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
          Me.ChGrabar.Value = 0
          Me.ChEliminar.Value = 0
          
        Case "Nomina de Subsidio"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Nomina de Subsidio'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
          Me.ChGrabar.Value = 0
          Me.ChEliminar.Value = 0
         
        Case "Despido y Renuncias"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Despido y Renuncias'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
          Me.ChGrabar.Value = 0
          Me.ChEliminar.Value = 0
         
        Case "Entradas y Salidas"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Entradas y Salidas'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
          Me.ChGrabar.Value = 0
          Me.ChEliminar.Value = 0
          
        Case "Calcular Horas Extra o Falta"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Calcular Horas Extra o Falta'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
          Me.ChGrabar.Value = 0
          Me.ChEliminar.Value = 0
         
        Case "Reportes"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Reportes'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
          Me.ChGrabar.Value = 0
          Me.ChEliminar.Value = 0
          
        Case "Controles Personalizados"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = False
          Me.ChEliminar.Enabled = False

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Controles Personalizados'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
          Me.ChGrabar.Value = 0
          Me.ChEliminar.Value = 0
          
        Case "Usuarios"
          Me.ChAbrir.Enabled = True
          Me.ChGrabar.Enabled = True
          Me.ChEliminar.Enabled = True

          
        '///////Chek abrir////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Abrir Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChAbrir.Value = 0
         Else
          Me.ChAbrir.Value = 1
         End If
         
        '///////Chek Grabar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Grabar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChGrabar.Value = 0
         Else
          Me.ChGrabar.Value = 1
         End If
         
        '///////Chek Borrar////
          Me.DtaNacceso.RecordSource = "SELECT Accesos.CodUsuario, Accesos.AccesoModulo From Accesos WHERE (((Accesos.CodUsuario)= " & CodigoUsuario & ") AND ((Accesos.AccesoModulo)='Borrar Usuarios'))"
          Me.DtaNacceso.Refresh
         If Me.DtaNacceso.Recordset.EOF Then
          Me.ChEliminar.Value = 0
         Else
          Me.ChEliminar.Value = 1
         End If
         
   End Select
Exit Sub
TipoErrs:
ControlErrores
Unload Me
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
