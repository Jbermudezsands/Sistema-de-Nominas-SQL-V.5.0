VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmTipoNomina 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Nominas"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9390
   HelpContextID   =   28
   Icon            =   "FrmTipoNomina.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmTipoNomina.frx":030A
   ScaleHeight     =   472
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   626
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   32
      Top             =   5760
      Width           =   3135
      Begin VB.CommandButton CmdAnterior 
         Caption         =   "Anterior"
         DownPicture     =   "FrmTipoNomina.frx":064C
         Height          =   375
         Left            =   120
         MouseIcon       =   "FrmTipoNomina.frx":212E
         MousePointer    =   99  'Custom
         Picture         =   "FrmTipoNomina.frx":2570
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdSiguiente 
         Caption         =   "Siguiente"
         DownPicture     =   "FrmTipoNomina.frx":4052
         Height          =   375
         Left            =   1560
         MouseIcon       =   "FrmTipoNomina.frx":5B34
         MousePointer    =   99  'Custom
         Picture         =   "FrmTipoNomina.frx":5F76
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdUltimo 
         Caption         =   "Ultimo"
         DownPicture     =   "FrmTipoNomina.frx":7A58
         Height          =   375
         Left            =   1560
         MouseIcon       =   "FrmTipoNomina.frx":953A
         MousePointer    =   99  'Custom
         Picture         =   "FrmTipoNomina.frx":997C
         TabIndex        =   34
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton CmdPrimero 
         Caption         =   "Primero"
         DownPicture     =   "FrmTipoNomina.frx":B45E
         Height          =   375
         Left            =   120
         MouseIcon       =   "FrmTipoNomina.frx":CF40
         MousePointer    =   99  'Custom
         Picture         =   "FrmTipoNomina.frx":D382
         TabIndex        =   33
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      DownPicture     =   "FrmTipoNomina.frx":EE64
      Height          =   375
      Left            =   4680
      MouseIcon       =   "FrmTipoNomina.frx":10946
      MousePointer    =   99  'Custom
      Picture         =   "FrmTipoNomina.frx":10D88
      TabIndex        =   31
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      DownPicture     =   "FrmTipoNomina.frx":1286A
      Height          =   375
      Left            =   4680
      MouseIcon       =   "FrmTipoNomina.frx":1434C
      MousePointer    =   99  'Custom
      Picture         =   "FrmTipoNomina.frx":1478E
      TabIndex        =   30
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      DownPicture     =   "FrmTipoNomina.frx":16270
      Height          =   375
      Left            =   7800
      MouseIcon       =   "FrmTipoNomina.frx":17D52
      MousePointer    =   99  'Custom
      Picture         =   "FrmTipoNomina.frx":18194
      TabIndex        =   29
      Top             =   6480
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.Frame Frame6 
         Height          =   975
         Left            =   120
         TabIndex        =   37
         Top             =   4320
         Width           =   6975
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "FrmTipoNomina.frx":19C76
            TabIndex        =   38
            Top             =   360
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtpSalida 
            Height          =   495
            Left            =   5160
            TabIndex        =   39
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   16711682
            CurrentDate     =   40750
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmTipoNomina.frx":19CF0
            TabIndex        =   40
            Top             =   360
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtpEntrada 
            Height          =   495
            Left            =   1680
            TabIndex        =   41
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   16711682
            CurrentDate     =   40750
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Moneda"
         Height          =   855
         Left            =   5640
         TabIndex        =   25
         Top             =   720
         Width           =   2895
         Begin VB.CheckBox ChkMantValor 
            Caption         =   "Mantenimiento de Valor"
            Height          =   375
            Left            =   1320
            TabIndex        =   28
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton OptUs 
            Caption         =   "Dólares"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton OptCS 
            Caption         =   "Córdobas"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Pago"
         Height          =   1335
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   5175
         Begin VB.OptionButton OptTodo 
            Caption         =   "Salario Fijo, Comision, Destajo"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   960
            Width           =   3375
         End
         Begin VB.OptionButton OptSalarioDestajoComision 
            Caption         =   "Salario al Destajo y Comisión"
            Height          =   375
            Left            =   2400
            TabIndex        =   23
            Top             =   600
            Width           =   2655
         End
         Begin VB.OptionButton OptSalarioFijoComision 
            Caption         =   "Salario Fijo y Comision"
            Height          =   375
            Left            =   2400
            TabIndex        =   22
            ToolTipText     =   "Se tiene un salario fijo y una comisión por ventas o sobreproducción. Aplican las horas extras"
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton OptSalarioDestajo 
            Caption         =   "Salario al Destajo"
            Height          =   375
            Left            =   240
            TabIndex        =   21
            ToolTipText     =   "Este es un sueldo por obra ya sea ventas, elaboración del algun producto o Horas trabajadas, puede incluir horas extras"
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton OptSalario 
            Caption         =   "Sueldo Fijo"
            Height          =   375
            Left            =   240
            TabIndex        =   20
            ToolTipText     =   "Este es un salario que se paga igual todos los meses, puede incluir horas extras"
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo de Calculo del INSS"
         Height          =   975
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   6975
         Begin VB.OptionButton Option1 
            Caption         =   "Calcular por Porcentaje"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Calcular por la Tabla"
            Height          =   255
            Left            =   3360
            TabIndex        =   15
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox TxtPorcentaje 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox TxtInssPatronal 
            Height          =   285
            Left            =   3240
            TabIndex        =   13
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "% Inss Laboral"
            Height          =   255
            Left            =   1800
            TabIndex        =   18
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "% Inss Patronal"
            Height          =   255
            Left            =   5160
            TabIndex        =   17
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.ComboBox CmbPeriodo 
         Height          =   315
         ItemData        =   "FrmTipoNomina.frx":19D6C
         Left            =   6720
         List            =   "FrmTipoNomina.frx":19D6E
         TabIndex        =   7
         Text            =   "Periodos de Nómina"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   4440
         MaxLength       =   25
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de Calculo del IR"
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   3240
         Width           =   5175
         Begin VB.TextBox TxtPorcientoIr 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Calcular por la Tabla"
            Height          =   255
            Left            =   2520
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Calcular por Porcentaje"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "Porcentaja a Calcular"
            Height          =   255
            Left            =   3120
            TabIndex        =   5
            Top             =   600
            Visible         =   0   'False
            Width           =   1935
         End
      End
      Begin MSDataListLib.DataCombo DBCCodigo 
         Bindings        =   "FrmTipoNomina.frx":19D70
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodTipoNomina"
         Text            =   ""
      End
      Begin VB.Image Image1 
         Height          =   1320
         Left            =   6960
         Picture         =   "FrmTipoNomina.frx":19D8C
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Período"
         Height          =   375
         Left            =   6000
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion:"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Código Tipo Nómina:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc DtaTipoNomina 
      Height          =   375
      Left            =   1080
      Top             =   7080
      Width           =   3135
      _ExtentX        =   5530
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
      Caption         =   "DtaTipoNomina"
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
End
Attribute VB_Name = "FrmTipoNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbPeriodo_LostFocus()
On Error GoTo TipoErrs
Select Case CmbPeriodo.Text
Case "Semanal Viernes"
Case "Semanal Sabado"
Case "Catorcenal los Viernes"
Case "Catorcenal los Sabados"
Case "Quincenal"
Case "Mensual"
Case "Trimestral"
Case "Semestral"
Case Else
   MsgBox "El período de nómina es erróneo", vbCritical
   CmbPeriodo.SetFocus
End Select

Exit Sub
TipoErrs:
ControlErrores
Unload Me
End Sub

Private Sub CmdAnterior_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Tipo de Nomina")
If Contesta Then
  cmdGrabar.Value = True
End If
 DtaTipoNomina.Recordset.MovePrevious

If DtaTipoNomina.Recordset.BOF Then
 DtaTipoNomina.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
DBCCodigo.Text = DtaTipoNomina.Recordset("CodTipoNomina")

End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdBorrar_Click()
On Error GoTo TipoErrs
Dim Respuesta, Rsp
'Elimino el registro activo en la pantalla
  Set Rsp = DtaTipoNomina.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando Tipo Nómina: " & TxtDescripcion.Text)
   If Respuesta = 6 Then
     Rsp.Delete
      DBCCodigo.Text = ""
      DtaTipoNomina.Recordset.MoveLast
      DtaTipoNomina.Recordset.MovePrevious
      Salida = False
   End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub cmdGrabar_Click()
'On Error GoTo TipoErrs
  Salida = False
  
  If CmbPeriodo.Text = "Periodos de Nómina" Then
     MsgBox "No ha seleccionado el período de la nómina"
     Exit Sub
  End If
  
  'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
      DtaTipoNomina.Refresh
      Do While Not DtaTipoNomina.Recordset.EOF
       If DtaTipoNomina.Recordset("CodTipoNomina") = DBCCodigo.Text Then
         'DtaTipoNomina.Recordset.Edit
         DtaTipoNomina.Recordset("nomina") = TxtDescripcion.Text
         DtaTipoNomina.Recordset("Periodo") = CmbPeriodo.Text
         
         If OptSalario Then
            DtaTipoNomina.Recordset("TipoPago") = "Salario Fijo"
         ElseIf OptSalarioDestajo Then
            DtaTipoNomina.Recordset("TipoPago") = "Salario Destajo"
         ElseIf OptSalarioFijoComision Then
            DtaTipoNomina.Recordset("TipoPago") = "Salario Fijo y Comision"
         ElseIf Me.OptSalarioDestajoComision Then
           DtaTipoNomina.Recordset("TipoPago") = "Salario Destajo y Comision"
         ElseIf OptTodo Then
           DtaTipoNomina.Recordset("TipoPago") = "Salario Fijo,Destajo y Comision"
           
         End If
         
         If Me.Option1.Value = True Then
           DtaTipoNomina.Recordset("TasaInss") = Me.txtPorcentaje.Text
           DtaTipoNomina.Recordset("TasaInssPatronal") = Me.TxtInssPatronal.Text
           DtaTipoNomina.Recordset("PorcientoInss") = 1
         Else
           DtaTipoNomina.Recordset("TasaInss") = 0#
           DtaTipoNomina.Recordset("PorcientoInss") = 0
         End If
         
         If Me.Option4.Value = True Then
           DtaTipoNomina.Recordset("PorcientoIr") = 1
           DtaTipoNomina.Recordset("TasaIr") = Me.TxtPorcientoIr.Text
         Else
           DtaTipoNomina.Recordset("PorcientoIr") = 0
           DtaTipoNomina.Recordset("TasaIr") = 0#
        End If
         
         If OptCS Then DtaTipoNomina.Recordset("Moneda") = "CS"
         If OptUs Then DtaTipoNomina.Recordset("Moneda") = "US"
         
         If ChkMantValor = 1 Then
            DtaTipoNomina.Recordset("MantValor") = 1
         Else
            DtaTipoNomina.Recordset("MantValor") = 0
         End If
         'DtaTipoNomina.Recordset("activa") = 0
         DtaTipoNomina.Recordset.Update
         
         DBCCodigo.Text = ""
         DtaTipoNomina.Recordset.MoveLast
         DtaTipoNomina.Recordset.MovePrevious
         Salida = False
         Exit Sub
             
      End If
      DtaTipoNomina.Recordset.MoveNext
      Loop
  'Si despues de Buscar no exite el codigo grabo todos los cambios
         DtaTipoNomina.Recordset.AddNew
         DtaTipoNomina.Recordset("CodTipoNomina") = DBCCodigo.Text
         DtaTipoNomina.Recordset("nomina") = TxtDescripcion.Text
         DtaTipoNomina.Recordset("Periodo") = CmbPeriodo.Text
         If OptSueldo Then
            DtaTipoNomina.Recordset("TipoPago") = "Salario Fijo"
         ElseIf OptSalarioDestajo Then
            DtaTipoNomina.Recordset("TipoPago") = "Salario Destajo"
         ElseIf OptSalarioFijoComision Then
            DtaTipoNomina.Recordset("TipoPago") = "Salario Fijo y Comision"
         ElseIf OptTodo Then
           DtaTipoNomina.Recordset("TipoPago") = "Salario Fijo,Destajo y Comision"
         Else
           DtaTipoNomina.Recordset("TipoPago") = "Salario Destajo y Comision"
         End If
         
         If OptCS Then DtaTipoNomina.Recordset("Moneda") = "CS"
         If OptUs Then DtaTipoNomina.Recordset("Moneda") = "US"
         
         If ChkMantValor = 1 Then
          DtaTipoNomina.Recordset("MantValor") = 1
         Else
             DtaTipoNomina.Recordset("MantValor") = 0
         End If
         DtaTipoNomina.Recordset("activa") = 0
         DtaTipoNomina.Recordset.Update
         DBCCodigo.Text = ""
         DtaTipoNomina.Recordset.MoveLast
         DtaTipoNomina.Recordset.MovePrevious
         Salida = False
         Exit Sub
         
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub CmdPrimero_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Tipo de Nomina")
If Contesta Then
  cmdGrabar.Value = True
End If
 DtaTipoNomina.Recordset.MoveFirst

If DtaTipoNomina.Recordset.BOF Then
 DtaTipoNomina.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
DBCCodigo.Text = DtaTipoNomina.Recordset("CodTipoNomina")

End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdSalir_Click()
 Unload Me
End Sub

Private Sub CmdSiguiente_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Tipo de Nomina")
If Contesta Then
  cmdGrabar.Value = True
End If
 DtaTipoNomina.Recordset.MoveNext

If DtaTipoNomina.Recordset.EOF Then
 DtaTipoNomina.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
DBCCodigo.Text = DtaTipoNomina.Recordset("CodTipoNomina")

End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdUltimo_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Tipo de Nomina")
If Contesta Then
  cmdGrabar.Value = True
End If
 DtaTipoNomina.Recordset.MoveLast

If DtaTipoNomina.Recordset.EOF Then
 DtaTipoNomina.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
DBCCodigo.Text = DtaTipoNomina.Recordset("CodTipoNomina")

End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub DBCCodigo_Change()
On Error GoTo TipoErrs

Evaluar = True
 'Al ejecutar algun cambio en el combo actualizo el nombre del departamento
   DtaTipoNomina.Refresh
   Do While Not DtaTipoNomina.Recordset.EOF
     If DtaTipoNomina.Recordset("CodTipoNomina") = DBCCodigo.Text Then
        TxtDescripcion.Text = DtaTipoNomina.Recordset("nomina")
        CmbPeriodo.Text = DtaTipoNomina.Recordset("Periodo")
        
        If DtaTipoNomina.Recordset("TipoPago") = "Salario Fijo" Then
              OptSalario.Value = True
        ElseIf DtaTipoNomina.Recordset("TipoPago") = "Salario Destajo" Then
              OptSalarioDestajo.Value = True
        ElseIf DtaTipoNomina.Recordset("TipoPago") = "Salario Fijo y Comision" Then
              OptSalarioFijoComision.Value = True
        ElseIf DtaTipoNomina.Recordset("TipoPago") = "Salario Destajo y Comision" Then
              OptSalarioDestajoComision.Value = True
        ElseIf DtaTipoNomina.Recordset("TipoPago") = "Salario Fijo,Destajo y Comision" Then
           OptTodo.Value = True
        End If
         
         If DtaTipoNomina.Recordset("Moneda") = "CS" Then OptCS.Value = True
         If DtaTipoNomina.Recordset("Moneda") = "US" Then OptUs.Value = True
         If DtaTipoNomina.Recordset("MantValor") = True Then
            ChkMantValor.Value = 1
        Else
            ChkMantValor.Value = 0
        End If
        
        If DtaTipoNomina.Recordset("PorcientoInss") = True Then
            Me.Option1.Value = 1
            Me.txtPorcentaje.Visible = True
            Me.Label5.Visible = True
            Me.txtPorcentaje.Text = DtaTipoNomina.Recordset("TasaInss")
            Me.TxtInssPatronal.Text = DtaTipoNomina.Recordset("TasaInssPatronal")
        Else
            Me.Option2.Value = True
            Me.txtPorcentaje.Visible = False
            Me.Label5.Visible = False
            Me.txtPorcentaje.Text = 0#
            Me.TxtInssPatronal.Text = 0#
        End If
        
        If DtaTipoNomina.Recordset("PorcientoIr") = True Then
            Me.Option4.Value = 1
            Me.TxtPorcientoIr.Visible = True
            Me.Label6.Visible = True
            Me.TxtPorcientoIr.Text = DtaTipoNomina.Recordset("TasaIr")
        Else
            Me.Option3.Value = True
            Me.TxtPorcientoIr.Visible = False
            Me.Label6.Visible = False
            Me.TxtPorcientoIr.Text = 0#
        End If
        
        
        
        Exit Do
     Else
        Me.TxtInssPatronal.Text = ""
        TxtDescripcion.Text = ""
        Me.txtPorcentaje.Text = ""
        Me.TxtPorcientoIr.Text = ""
        CmbPeriodo.Text = "Periodos de Nómina"
     End If
       DtaTipoNomina.Recordset.MoveNext
   Loop
  Salida = False
  Exit Sub
TipoErrs:
  ControlErrores
  Unload Me
End Sub



Private Sub DBCCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 TxtDescripcion.SetFocus
End If
End Sub

Private Sub Form_Activate()
 'If Not BTipoNomina = True Then
  ' CmdBorrar.Enabled = False
 'End If
 'If Not GTipoNomina = True Then
   'CmdGrabar.Enabled = False
 'End If
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
With Me.DtaTipoNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoNomina"
   .Refresh
End With

Me.dtpSalida.Value = Now

CmbPeriodo.AddItem "Semanal Viernes"
CmbPeriodo.AddItem "Semanal Sabado"
CmbPeriodo.AddItem "Catorcenal los Viernes"
CmbPeriodo.AddItem "Catorcenal los Sabados"
CmbPeriodo.AddItem "Quincenal"
CmbPeriodo.AddItem "Mensual"
CmbPeriodo.AddItem "Trimestral"
CmbPeriodo.AddItem "Semestral"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GTipoNomina = True Then
ValidaSalida ("en la Tabla Tipo de Nomina")
 If Contesta Then
  cmdGrabar.Value = True
  Salida = False
  Unload Me
 Else
  Salida = False
  Unload Me
 End If
End If
End Sub

Private Sub OptCS_Click()
ChkMantValor.Enabled = True
End Sub

Private Sub OptSueldo_Click()

End Sub

Private Sub Option1_Click()
 If Option1.Value = True Then
   Me.txtPorcentaje.Visible = True
   Me.Label5.Visible = True
   Me.TxtInssPatronal.Visible = True
   Me.Label7.Visible = True
 End If

End Sub

Private Sub Option2_Click()
 If Option2.Value = True Then
   Me.txtPorcentaje.Visible = False
   Me.Label5.Visible = False
   Me.TxtInssPatronal.Visible = False
   Me.Label7.Visible = False
 End If
End Sub

Private Sub Option3_Click()
 If Me.Option3.Value = True Then
   Me.TxtPorcientoIr.Visible = False
   Me.Label6.Visible = False
 End If
End Sub

Private Sub Option4_Click()
 If Me.Option4.Value = True Then
   Me.TxtPorcientoIr.Visible = True
   Me.Label6.Visible = True
 End If
End Sub

Private Sub OptUs_Click()
ChkMantValor.Value = 0
ChkMantValor.Enabled = False
End Sub

Private Sub TxtDescripcion_Change()
Salida = True
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 TxtDias.SetFocus
Else
   Evaluar = False
  End If
End Sub

Private Sub TxtDias_Change()
Salida = True
End Sub

Private Sub TxtDias_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 cmdGrabar.SetFocus
Else
   Evaluar = False
  End If
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
