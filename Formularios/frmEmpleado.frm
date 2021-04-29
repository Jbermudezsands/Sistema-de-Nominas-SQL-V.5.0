VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form frmEmpleado 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro Empleados"
   ClientHeight    =   6870
   ClientLeft      =   -60
   ClientTop       =   345
   ClientWidth     =   9720
   ClipControls    =   0   'False
   HelpContextID   =   12
   Icon            =   "frmEmpleado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   458
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   648
   Begin VB.TextBox TxtCodigoEmpleados 
      Height          =   375
      Left            =   6480
      TabIndex        =   208
      Text            =   "Text3"
      Top             =   6000
      Visible         =   0   'False
      Width           =   3135
   End
   Begin XtremeSuiteControls.CheckBox ChkDolarizado 
      Height          =   255
      Left            =   5160
      TabIndex        =   183
      Top             =   6120
      Width           =   2655
      _Version        =   786432
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Salario Basico Dolarizado"
      UseVisualStyle  =   -1  'True
   End
   Begin MSAdodcLib.Adodc AdoNumerosDisponibles 
      Height          =   375
      Left            =   600
      Top             =   7800
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "AdoNumerosDisponibles"
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
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5235
      ScaleWidth      =   9435
      TabIndex        =   23
      Top             =   480
      Width           =   9495
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5640
         TabIndex        =   207
         Text            =   "Text2"
         Top             =   5280
         Width           =   1695
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5175
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   9450
         _ExtentX        =   16669
         _ExtentY        =   9128
         _Version        =   393216
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   520
         BackColor       =   -2147483629
         TabCaption(0)   =   "Generales"
         TabPicture(0)   =   "frmEmpleado.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Histórico"
         TabPicture(1)   =   "frmEmpleado.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture3"
         Tab(1).Control(1)=   "Label74"
         Tab(1).Control(2)=   "Label72"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Información"
         TabPicture(2)   =   "frmEmpleado.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Picture4"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Préstamo"
         TabPicture(3)   =   "frmEmpleado.frx":035E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Picture5"
         Tab(3).Control(1)=   "DbgrLibreta"
         Tab(3).Control(2)=   "SSTab2"
         Tab(3).Control(3)=   "Label67"
         Tab(3).ControlCount=   4
         TabCaption(4)   =   "Incentivos"
         TabPicture(4)   =   "frmEmpleado.frx":037A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "CmdAnular"
         Tab(4).Control(1)=   "Picture6"
         Tab(4).Control(2)=   "CmdEliminarIncentivo"
         Tab(4).Control(3)=   "CmdHistoIncentivos"
         Tab(4).Control(4)=   "DbGIncentivos"
         Tab(4).ControlCount=   5
         TabCaption(5)   =   "Deducciones"
         TabPicture(5)   =   "frmEmpleado.frx":0396
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "DbgDeducciones"
         Tab(5).Control(1)=   "CmdEliminarDeduccion"
         Tab(5).Control(2)=   "Picture7"
         Tab(5).ControlCount=   3
         TabCaption(6)   =   "Subsidios"
         TabPicture(6)   =   "frmEmpleado.frx":03B2
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Picture8"
         Tab(6).Control(1)=   "CmdEliminarSubsidio"
         Tab(6).Control(2)=   "DbgrSubsidios"
         Tab(6).ControlCount=   3
         Begin VB.CommandButton CmdAnular 
            Caption         =   "Anular"
            DownPicture     =   "frmEmpleado.frx":03CE
            Enabled         =   0   'False
            Height          =   375
            Left            =   -69600
            Picture         =   "frmEmpleado.frx":1EB0
            TabIndex        =   185
            Top             =   3660
            Width           =   1455
         End
         Begin VB.PictureBox Picture8 
            Height          =   2655
            Left            =   -74880
            ScaleHeight     =   2595
            ScaleWidth      =   3435
            TabIndex        =   173
            Top             =   900
            Width           =   3495
            Begin VB.CommandButton CmdAgregarSubsidio 
               Caption         =   "Agregar"
               DownPicture     =   "frmEmpleado.frx":3992
               Height          =   375
               Left            =   1800
               Picture         =   "frmEmpleado.frx":5474
               TabIndex        =   177
               Top             =   2040
               Width           =   1455
            End
            Begin VB.TextBox TxtNumVecesSubsidio 
               Height          =   375
               Left            =   2160
               TabIndex        =   176
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox TxtMontoSubsidio 
               Height          =   285
               Left            =   1080
               TabIndex        =   175
               Top             =   840
               Width           =   2175
            End
            Begin VB.TextBox TxtDescripcion 
               Height          =   285
               Left            =   1080
               MaxLength       =   25
               TabIndex        =   174
               Top             =   1080
               Width           =   2175
            End
            Begin MSDataListLib.DataCombo DBCTipoSubsidio 
               Bindings        =   "frmEmpleado.frx":6F56
               Height          =   315
               Left            =   120
               TabIndex        =   178
               Top             =   360
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Subsidio"
               Text            =   "Tipo Subsidio"
            End
            Begin VB.Label Label66 
               Caption         =   "Subsidio"
               Height          =   255
               Left            =   120
               TabIndex        =   182
               Top             =   120
               Width           =   1815
            End
            Begin VB.Label Label42 
               Caption         =   "Número de Veces"
               Height          =   255
               Left            =   360
               TabIndex        =   181
               Top             =   1680
               Width           =   1575
            End
            Begin VB.Label Label43 
               Caption         =   "Monto"
               Height          =   255
               Left            =   240
               TabIndex        =   180
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
               Caption         =   "Descripción"
               Height          =   255
               Left            =   120
               TabIndex        =   179
               Top             =   1080
               Width           =   855
            End
         End
         Begin VB.PictureBox Picture7 
            Height          =   2535
            Left            =   -74880
            ScaleHeight     =   2475
            ScaleWidth      =   3315
            TabIndex        =   165
            Top             =   900
            Width           =   3375
            Begin VB.TextBox TxtVecesDeduccion 
               Height          =   375
               Left            =   1920
               TabIndex        =   168
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox TxtMontoDeduccion 
               Height          =   405
               Left            =   960
               TabIndex        =   167
               Top             =   960
               Width           =   2175
            End
            Begin VB.CommandButton CmdAgregarDeduccion 
               Caption         =   "Agregar"
               DownPicture     =   "frmEmpleado.frx":6F74
               Height          =   495
               Left            =   1680
               Picture         =   "frmEmpleado.frx":8A56
               TabIndex        =   166
               Top             =   1800
               Width           =   1455
            End
            Begin MSDataListLib.DataCombo DbcDeducciones 
               Bindings        =   "frmEmpleado.frx":A538
               Height          =   315
               Left            =   240
               TabIndex        =   169
               Top             =   360
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Deduccion"
               Text            =   "Tipo Deducciones"
            End
            Begin VB.Label Label52 
               Caption         =   "Deduccion"
               Height          =   255
               Left            =   120
               TabIndex        =   172
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label57 
               Caption         =   "Monto"
               Height          =   255
               Left            =   240
               TabIndex        =   171
               Top             =   960
               Width           =   615
            End
            Begin VB.Label Label58 
               Caption         =   "Número de Veces"
               Height          =   255
               Left            =   120
               TabIndex        =   170
               Top             =   1440
               Width           =   1935
            End
         End
         Begin VB.PictureBox Picture6 
            Height          =   2655
            Left            =   -74880
            ScaleHeight     =   2595
            ScaleWidth      =   3435
            TabIndex        =   157
            Top             =   900
            Width           =   3495
            Begin VB.TextBox TxtNumVeces 
               Height          =   375
               Left            =   2040
               TabIndex        =   159
               ToolTipText     =   "Digite ""n"" para indicar que este incentivo es de caracter infinito"
               Top             =   1440
               Width           =   1215
            End
            Begin VB.CommandButton CmdAgregarIncentivo 
               Caption         =   "Agregar  "
               DownPicture     =   "frmEmpleado.frx":A557
               Height          =   495
               Left            =   1800
               Picture         =   "frmEmpleado.frx":C039
               TabIndex        =   158
               Top             =   1920
               Width           =   1455
            End
            Begin MSDataListLib.DataCombo DbCTipoIncentivo 
               Bindings        =   "frmEmpleado.frx":DB1B
               Height          =   315
               Left            =   360
               TabIndex        =   160
               Top             =   480
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Incentivo"
               Text            =   "Tipo Incentivos"
            End
            Begin MSMask.MaskEdBox TxtMonto 
               Height          =   375
               Left            =   1080
               TabIndex        =   161
               Top             =   960
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   661
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label Label40 
               Caption         =   "Incentivo"
               Height          =   375
               Left            =   240
               TabIndex        =   164
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label27 
               Caption         =   "Monto"
               Height          =   255
               Left            =   360
               TabIndex        =   163
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label49 
               Caption         =   "Número de Veces"
               Height          =   255
               Left            =   240
               TabIndex        =   162
               Top             =   1440
               Width           =   1815
            End
         End
         Begin VB.PictureBox Picture5 
            Height          =   1935
            Left            =   -74760
            ScaleHeight     =   1875
            ScaleWidth      =   4875
            TabIndex        =   143
            Top             =   1140
            Width           =   4935
            Begin VB.TextBox TxtInteresprestamo 
               Height          =   285
               Left            =   720
               TabIndex        =   150
               Top             =   840
               Width           =   615
            End
            Begin VB.Frame Frame4 
               Caption         =   "Moneda"
               Height          =   975
               Left            =   2520
               TabIndex        =   144
               Top             =   120
               Width           =   1215
               Begin VB.OptionButton OptUS 
                  Caption         =   "US$"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   154
                  Top             =   600
                  Width           =   735
               End
               Begin VB.OptionButton OptC 
                  Caption         =   "C$"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   152
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   735
               End
            End
            Begin VB.CheckBox ChkTipoPago 
               Caption         =   "Cuotas Iguales"
               Height          =   255
               Left            =   2400
               TabIndex        =   156
               Top             =   1200
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.TextBox TxtCuotas 
               Height          =   285
               Left            =   1560
               TabIndex        =   148
               Top             =   480
               Width           =   495
            End
            Begin MSMask.MaskEdBox TxtSaldo 
               Height          =   285
               Left            =   600
               TabIndex        =   145
               Top             =   1200
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TxtMontoPrestamoUS 
               Height          =   285
               Left            =   600
               TabIndex        =   146
               Top             =   120
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label59 
               Caption         =   "%"
               Height          =   255
               Left            =   1440
               TabIndex        =   155
               Top             =   840
               Width           =   255
            End
            Begin VB.Label Label7 
               Caption         =   "Interes"
               Height          =   255
               Left            =   120
               TabIndex        =   153
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label64 
               Caption         =   "Monto"
               Height          =   375
               Left            =   120
               TabIndex        =   151
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label63 
               Caption         =   "Número de Cuotas"
               Height          =   255
               Left            =   120
               TabIndex        =   149
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label62 
               Caption         =   "Saldo"
               Height          =   255
               Left            =   120
               TabIndex        =   147
               Top             =   1200
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture4 
            Height          =   4455
            Left            =   -74880
            ScaleHeight     =   4395
            ScaleWidth      =   9075
            TabIndex        =   106
            Top             =   480
            Width           =   9135
            Begin VB.TextBox TxtReembolso 
               Height          =   285
               Left            =   8160
               TabIndex        =   212
               Top             =   3600
               Width           =   735
            End
            Begin VB.CheckBox ChkDeducirPorcentaje 
               Caption         =   "Deducir por Porcentaje"
               Height          =   255
               Left            =   6480
               TabIndex        =   211
               Top             =   2160
               Width           =   2535
            End
            Begin VB.TextBox TxtViatico 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   5760
               TabIndex        =   209
               Text            =   "0.00"
               Top             =   3960
               Width           =   735
            End
            Begin VB.TextBox txtAumentoBasico 
               Height          =   285
               Left            =   8160
               TabIndex        =   204
               Top             =   3240
               Width           =   735
            End
            Begin VB.TextBox TxtDiasBasico 
               Alignment       =   1  'Right Justify
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
               Left            =   5760
               TabIndex        =   202
               Text            =   "0"
               Top             =   3600
               Width           =   735
            End
            Begin VB.TextBox TxtDiasAdicionales 
               Alignment       =   1  'Right Justify
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
               Left            =   2640
               TabIndex        =   200
               Text            =   "0"
               Top             =   3240
               Width           =   735
            End
            Begin VB.TextBox TxtSalarioPorciento 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   5760
               TabIndex        =   198
               Text            =   "0.00"
               ToolTipText     =   "Comisión"
               Top             =   3240
               Width           =   975
            End
            Begin VB.CheckBox ChkHorasTurno 
               Caption         =   "Calcular Horas x Turnos"
               Height          =   255
               Left            =   3840
               TabIndex        =   197
               Top             =   2160
               Width           =   2775
            End
            Begin VB.TextBox TxtSueldoPeriodo 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   5520
               TabIndex        =   139
               Text            =   "0.00"
               ToolTipText     =   "Sueldo Fijo del Periodo"
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox TxtPorcientoHora 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   6840
               TabIndex        =   122
               Top             =   1920
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Calcular % Incentivo Horas Extras"
               Height          =   375
               Left            =   3840
               TabIndex        =   121
               Top             =   1800
               Width           =   2895
            End
            Begin VB.CheckBox ChkSalarioFijo 
               Caption         =   "Calcular Solo Salario Fijo"
               Height          =   255
               Left            =   120
               TabIndex        =   120
               Top             =   960
               Value           =   1  'Checked
               Width           =   3255
            End
            Begin VB.Frame Frame12 
               Caption         =   "Otros Ingresos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1455
               Left            =   6720
               TabIndex        =   117
               Top             =   360
               Width           =   2055
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
                  Height          =   255
                  Left            =   480
                  OleObjectBlob   =   "frmEmpleado.frx":DB3A
                  TabIndex        =   141
                  Top             =   840
                  Width           =   975
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
                  Height          =   255
                  Left            =   720
                  OleObjectBlob   =   "frmEmpleado.frx":DBAE
                  TabIndex        =   140
                  Top             =   240
                  Width           =   615
               End
               Begin VB.TextBox TxtDescripOtrIngre 
                  Height          =   285
                  Left            =   120
                  MaxLength       =   20
                  TabIndex        =   119
                  Top             =   1080
                  Width           =   1815
               End
               Begin VB.TextBox TxtOtrosIngresos 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   480
                  TabIndex        =   118
                  Text            =   "0.00"
                  Top             =   480
                  Width           =   975
               End
            End
            Begin VB.TextBox TxtCodGrupo 
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   116
               Top             =   3960
               Width           =   735
            End
            Begin VB.TextBox TxtDiasDescuento 
               Alignment       =   1  'Right Justify
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
               Left            =   2640
               TabIndex        =   115
               Text            =   "0"
               Top             =   2760
               Width           =   735
            End
            Begin VB.TextBox TxtCodTipoNomina 
               Height          =   285
               Left            =   7080
               Locked          =   -1  'True
               TabIndex        =   114
               Top             =   120
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.ComboBox CmbTipoPago 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmEmpleado.frx":DC16
               Left            =   1200
               List            =   "frmEmpleado.frx":DC26
               TabIndex        =   113
               Top             =   600
               Width           =   2295
            End
            Begin VB.TextBox TxtComision 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   5520
               TabIndex        =   112
               Text            =   "0.00"
               ToolTipText     =   "Comisión"
               Top             =   1440
               Width           =   975
            End
            Begin VB.TextBox TxtTarifaHoraria 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   5520
               TabIndex        =   111
               Text            =   "0.00"
               Top             =   960
               Width           =   975
            End
            Begin VB.ComboBox CmbSalarioMinimo 
               Height          =   315
               ItemData        =   "frmEmpleado.frx":DC66
               Left            =   1800
               List            =   "frmEmpleado.frx":DC70
               TabIndex        =   110
               Top             =   1320
               Width           =   1455
            End
            Begin VB.ComboBox CmbExentoInss 
               Height          =   315
               ItemData        =   "frmEmpleado.frx":DC86
               Left            =   1800
               List            =   "frmEmpleado.frx":DC90
               TabIndex        =   109
               Top             =   1680
               Width           =   1455
            End
            Begin VB.ComboBox CmbPagoInssPatronal 
               Height          =   315
               ItemData        =   "frmEmpleado.frx":DCA6
               Left            =   1800
               List            =   "frmEmpleado.frx":DCB0
               TabIndex        =   108
               Top             =   2400
               Width           =   1455
            End
            Begin VB.ComboBox CmbExentoIr 
               Height          =   315
               ItemData        =   "frmEmpleado.frx":DCC6
               Left            =   1800
               List            =   "frmEmpleado.frx":DCD0
               TabIndex        =   107
               Top             =   2040
               Width           =   1455
            End
            Begin MSDataListLib.DataCombo DBCGrupo 
               Bindings        =   "frmEmpleado.frx":DCE6
               Height          =   315
               Left            =   960
               TabIndex        =   123
               Top             =   3960
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Grupo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DBCTipoNomina 
               Bindings        =   "frmEmpleado.frx":DCFD
               Height          =   315
               Left            =   1200
               TabIndex        =   124
               Top             =   240
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Nomina"
               Text            =   ""
            End
            Begin VB.Label Label75 
               Caption         =   "Reembolso"
               Height          =   255
               Left            =   7320
               TabIndex        =   213
               Top             =   3600
               Width           =   855
            End
            Begin VB.Label Label73 
               Caption         =   "Viatico x Dia Asistecia"
               Height          =   255
               Left            =   3840
               TabIndex        =   210
               Top             =   3960
               Width           =   1695
            End
            Begin VB.Label Label71 
               Caption         =   "Aumento Basico"
               Height          =   255
               Left            =   6960
               TabIndex        =   205
               Top             =   3240
               Width           =   1215
            End
            Begin VB.Label Label70 
               Caption         =   "Rtar Dia Basico"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   3840
               TabIndex        =   203
               Top             =   3600
               Width           =   2535
            End
            Begin VB.Label Label69 
               Caption         =   "Dias Adicionales"
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
               Height          =   255
               Left            =   120
               TabIndex        =   201
               Top             =   3240
               Width           =   2415
            End
            Begin VB.Label Label17 
               Caption         =   "Salario %"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3840
               TabIndex        =   199
               Top             =   3240
               Width           =   975
            End
            Begin VB.Label LblPor 
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4080
               TabIndex        =   138
               Top             =   1440
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label Label53 
               Alignment       =   2  'Center
               Caption         =   "Grupos de Nómina"
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
               TabIndex        =   137
               Top             =   3600
               Width           =   3375
            End
            Begin VB.Label Label38 
               Caption         =   "Dias de Descuento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   120
               TabIndex        =   136
               Top             =   2760
               Width           =   2535
            End
            Begin VB.Label Label39 
               Caption         =   "Opciones de Calculo:"
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
               Left            =   3720
               TabIndex        =   135
               Top             =   2760
               Width           =   3375
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000002&
               BorderStyle     =   4  'Dash-Dot
               X1              =   3600
               X2              =   9000
               Y1              =   2640
               Y2              =   2640
            End
            Begin VB.Label Label19 
               Caption         =   "Codigo  Tipo Nomina"
               Height          =   255
               Left            =   5160
               TabIndex        =   134
               Top             =   120
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Label Label51 
               Caption         =   "Comision"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4440
               TabIndex        =   133
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label34 
               Caption         =   "Tipo Nóminas:"
               Height          =   255
               Left            =   120
               TabIndex        =   132
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label35 
               Caption         =   "Tipo de Pago:"
               Height          =   255
               Left            =   120
               TabIndex        =   131
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label36 
               Caption         =   "Sueldo Por Periodo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   130
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label Label37 
               Alignment       =   2  'Center
               Caption         =   "Tarifa / Cantidad"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   129
               Top             =   960
               Width           =   1815
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00800000&
               BorderStyle     =   5  'Dash-Dot-Dot
               X1              =   3600
               X2              =   3600
               Y1              =   120
               Y2              =   4200
            End
            Begin VB.Label Label46 
               Caption         =   "Salario Minimo"
               Height          =   375
               Left            =   120
               TabIndex        =   128
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label Label47 
               Caption         =   "Excento INSS"
               Height          =   255
               Left            =   120
               TabIndex        =   127
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label48 
               Caption         =   "Pago INSS Patronal"
               Height          =   375
               Left            =   120
               TabIndex        =   126
               Top             =   2400
               Width           =   1575
            End
            Begin VB.Label Label50 
               Caption         =   "Excento IR"
               Height          =   255
               Left            =   120
               TabIndex        =   125
               Top             =   2040
               Width           =   975
            End
         End
         Begin VB.PictureBox Picture3 
            Height          =   4335
            Left            =   -74880
            ScaleHeight     =   4275
            ScaleWidth      =   8955
            TabIndex        =   72
            Top             =   480
            Width           =   9015
            Begin VB.CommandButton CmdCuentas 
               Caption         =   "Cuentas Contables"
               Height          =   375
               Left            =   4080
               TabIndex        =   193
               Top             =   3000
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker MaskEdContrato 
               Height          =   315
               Left            =   1320
               TabIndex        =   190
               Top             =   720
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Format          =   16711681
               CurrentDate     =   40886
            End
            Begin MSComCtl2.DTPicker MaskEdNacimiento 
               Height          =   315
               Left            =   1320
               TabIndex        =   189
               Top             =   360
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Format          =   16711681
               CurrentDate     =   40886
            End
            Begin MSComCtl2.DTPicker DTPFechaContratoVaca 
               Height          =   315
               Left            =   1320
               TabIndex        =   188
               Top             =   1080
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Format          =   16711681
               CurrentDate     =   40213
            End
            Begin VB.CommandButton CmdIncapacidad 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Index           =   0
               Left            =   8520
               MouseIcon       =   "frmEmpleado.frx":DD19
               Picture         =   "frmEmpleado.frx":E15B
               Style           =   1  'Graphical
               TabIndex        =   105
               ToolTipText     =   "Consulta Tabla Incapacidades"
               Top             =   1680
               Width           =   255
            End
            Begin VB.TextBox TxtAumento 
               Height          =   285
               Left            =   4080
               MaxLength       =   2
               TabIndex        =   80
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox TxtMotivoBaja 
               Height          =   285
               Left            =   4080
               MaxLength       =   150
               TabIndex        =   79
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtMotivoAumento 
               Height          =   525
               Left            =   4080
               MaxLength       =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   78
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox TxtMotivoSuspencion 
               Height          =   525
               Left            =   6960
               MaxLength       =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   77
               Top             =   1080
               Width           =   1455
            End
            Begin VB.TextBox TxtSueldoInicial 
               Height          =   285
               Left            =   1320
               TabIndex        =   76
               Text            =   "0.00"
               Top             =   2520
               Width           =   1455
            End
            Begin VB.TextBox TxtSueldoActual 
               Height          =   285
               Left            =   1320
               TabIndex        =   75
               Text            =   "0.00"
               Top             =   3240
               Width           =   1455
            End
            Begin VB.TextBox TxtSueldoAnterior 
               Height          =   285
               Left            =   1320
               TabIndex        =   74
               Text            =   "0.00"
               Top             =   2880
               Width           =   1455
            End
            Begin VB.ComboBox CmbIncapacidad 
               Height          =   315
               ItemData        =   "frmEmpleado.frx":E59D
               Left            =   7080
               List            =   "frmEmpleado.frx":E5A7
               TabIndex        =   73
               Top             =   1680
               Width           =   1335
            End
            Begin MSDataListLib.DataCombo DBCargoInicial 
               Bindings        =   "frmEmpleado.frx":E5B3
               Height          =   315
               Left            =   1320
               TabIndex        =   81
               Top             =   1440
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Cargo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DBCargoAnterior 
               Bindings        =   "frmEmpleado.frx":E5CA
               Height          =   315
               Left            =   1320
               TabIndex        =   82
               Top             =   1800
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Cargo"
               Text            =   ""
            End
            Begin MSMask.MaskEdBox MaskEdAumento 
               Height          =   285
               Left            =   4080
               TabIndex        =   83
               Top             =   2040
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "99/99/9999"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBaja 
               Height          =   285
               Left            =   4080
               TabIndex        =   84
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "99/99/9999"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdFinalSusp 
               Height          =   285
               Left            =   6960
               TabIndex        =   85
               Top             =   720
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "99/99/9999"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdSuspencion 
               Height          =   285
               Left            =   6960
               TabIndex        =   86
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "99/99/9999"
               PromptChar      =   "_"
            End
            Begin MSDataListLib.DataCombo DBCargoActual 
               Bindings        =   "frmEmpleado.frx":E5E1
               Height          =   315
               Left            =   1320
               TabIndex        =   87
               Top             =   2160
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Cargo"
               Text            =   ""
            End
            Begin XtremeSuiteControls.CheckBox ChkSueldoActual 
               Height          =   255
               Left            =   1320
               TabIndex        =   196
               Top             =   3600
               Width           =   4335
               _Version        =   786432
               _ExtentX        =   7646
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Sueldo Actual --> Basico en Liquidacion y Vacaciones"
               UseVisualStyle  =   -1  'True
            End
            Begin VB.Label Label68 
               Caption         =   "Contrato Vac:"
               Height          =   255
               Left            =   120
               TabIndex        =   186
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label15 
               Caption         =   "Nacimiento:"
               Height          =   255
               Left            =   240
               TabIndex        =   104
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label16 
               Caption         =   "Contrato:"
               Height          =   255
               Left            =   480
               TabIndex        =   103
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label18 
               Caption         =   "Fecha  Aumento"
               Height          =   255
               Left            =   2880
               TabIndex        =   102
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Label Label20 
               Caption         =   "Inicio  Suspención"
               Height          =   255
               Left            =   5520
               TabIndex        =   101
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label21 
               Caption         =   "Fecha Baja"
               Height          =   255
               Left            =   3120
               TabIndex        =   100
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label23 
               Caption         =   "Cargo anterior"
               Height          =   255
               Left            =   120
               TabIndex        =   99
               Top             =   1800
               Width           =   1215
            End
            Begin VB.Label Label24 
               Caption         =   "Cargo Actual"
               Height          =   255
               Left            =   240
               TabIndex        =   98
               Top             =   2160
               Width           =   1095
            End
            Begin VB.Label Label25 
               Caption         =   "Motivo Aumento:"
               Height          =   495
               Left            =   3240
               TabIndex        =   97
               Top             =   1440
               Width           =   735
            End
            Begin VB.Label Label26 
               Caption         =   "Motivo Suspención"
               Height          =   375
               Left            =   5760
               TabIndex        =   96
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label Label28 
               Caption         =   "Motivo Baja"
               Height          =   255
               Left            =   3120
               TabIndex        =   95
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label29 
               Caption         =   "Sueldo Inicial"
               Height          =   255
               Left            =   120
               TabIndex        =   94
               Top             =   2520
               Width           =   1215
            End
            Begin VB.Label Label30 
               Caption         =   "Sueldo Anterior"
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   2880
               Width           =   1215
            End
            Begin VB.Label Label31 
               Caption         =   "Sueldo Actual"
               Height          =   255
               Left            =   120
               TabIndex        =   92
               Top             =   3240
               Width           =   1095
            End
            Begin VB.Label Label33 
               Caption         =   "Fin Suspención"
               Height          =   255
               Left            =   5760
               TabIndex        =   91
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label22 
               Caption         =   "Cargo Inicial"
               Height          =   255
               Left            =   240
               TabIndex        =   90
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label32 
               Caption         =   "Aumento"
               Height          =   255
               Left            =   3240
               TabIndex        =   89
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label11 
               Caption         =   "Incapcidades"
               Height          =   255
               Left            =   6000
               TabIndex        =   88
               Top             =   1680
               Width           =   1095
            End
         End
         Begin VB.PictureBox Picture2 
            Height          =   4575
            Left            =   120
            ScaleHeight     =   4515
            ScaleWidth      =   9075
            TabIndex        =   45
            Top             =   360
            Width           =   9135
            Begin VB.TextBox Text1 
               Height          =   315
               Left            =   3600
               TabIndex        =   218
               Top             =   0
               Visible         =   0   'False
               Width           =   3375
            End
            Begin VB.TextBox TxtNumHijos 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   6480
               TabIndex        =   216
               Top             =   3600
               Width           =   1935
            End
            Begin VB.TextBox TxtTelefono 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   1440
               TabIndex        =   214
               Top             =   4080
               Width           =   1935
            End
            Begin VB.TextBox TxtCuentaBanco 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6480
               MaxLength       =   30
               TabIndex        =   194
               Top             =   4080
               Width           =   2415
            End
            Begin VB.CommandButton CmdDisponibles 
               Caption         =   "No Disponible"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2040
               TabIndex        =   184
               Top             =   1440
               Width           =   1300
            End
            Begin VB.CheckBox ChkSuspendido 
               Caption         =   "Subsidio"
               Height          =   255
               Left            =   3360
               TabIndex        =   55
               Top             =   1560
               Width           =   1215
            End
            Begin VB.TextBox TxtNRuc 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6480
               MaxLength       =   20
               TabIndex        =   9
               Top             =   1320
               Width           =   2055
            End
            Begin VB.TextBox TxtCodCargo 
               Height          =   285
               Left            =   600
               TabIndex        =   54
               Text            =   "TxtCodCargo"
               Top             =   480
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox TxtNombre2 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1440
               MaxLength       =   20
               TabIndex        =   1
               Top             =   2640
               Width           =   2535
            End
            Begin VB.TextBox TxtApellido1 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1440
               MaxLength       =   20
               TabIndex        =   2
               Top             =   3000
               Width           =   2535
            End
            Begin VB.TextBox TxtApellido2 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1440
               MaxLength       =   20
               TabIndex        =   3
               Top             =   3360
               Width           =   2535
            End
            Begin VB.ComboBox CmbSindicalista 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmEmpleado.frx":E5F8
               Left            =   6480
               List            =   "frmEmpleado.frx":E602
               TabIndex        =   14
               Top             =   3240
               Width           =   1575
            End
            Begin VB.TextBox TxtDireccion 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1440
               MaxLength       =   200
               TabIndex        =   4
               Top             =   3720
               Width           =   3135
            End
            Begin VB.TextBox TxtCodDepartamento 
               Height          =   285
               Left            =   360
               TabIndex        =   53
               Text            =   "TxtCodDepartamento"
               Top             =   840
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.TextBox TxtNInss 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6480
               MaxLength       =   20
               TabIndex        =   10
               Top             =   1680
               Width           =   2055
            End
            Begin VB.ComboBox CmbSexo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "frmEmpleado.frx":E60E
               Left            =   7320
               List            =   "frmEmpleado.frx":E618
               TabIndex        =   8
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox TxtCodPostal 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   7440
               MaxLength       =   10
               TabIndex        =   7
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox TxtNacionalidad 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   7440
               MaxLength       =   12
               TabIndex        =   6
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox TxtNombre1 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1440
               MaxLength       =   20
               TabIndex        =   0
               Top             =   2280
               Width           =   2535
            End
            Begin VB.TextBox TxtNumCedula 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6480
               MaxLength       =   30
               TabIndex        =   11
               Top             =   2040
               Width           =   2415
            End
            Begin VB.CommandButton CmdBuscarEmpleado 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   5160
               Picture         =   "frmEmpleado.frx":E631
               Style           =   1  'Graphical
               TabIndex        =   52
               Top             =   1920
               Width           =   375
            End
            Begin VB.Frame Frame6 
               Caption         =   "Foto"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   2160
               TabIndex        =   49
               Top             =   120
               Width           =   1695
               Begin VB.CommandButton AgregaFoto 
                  Caption         =   "Agregar"
                  DownPicture     =   "frmEmpleado.frx":E77F
                  Height          =   375
                  Left            =   120
                  Picture         =   "frmEmpleado.frx":10261
                  TabIndex        =   51
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.CommandButton EliminaFoto 
                  Caption         =   "Borrar"
                  DownPicture     =   "frmEmpleado.frx":11D43
                  Height          =   375
                  Left            =   120
                  Picture         =   "frmEmpleado.frx":13825
                  TabIndex        =   50
                  Top             =   600
                  Width           =   1455
               End
            End
            Begin VB.TextBox TxtCodEmpleado 
               Height          =   285
               Left            =   360
               TabIndex        =   5
               Top             =   480
               Visible         =   0   'False
               Width           =   3135
            End
            Begin VB.Frame Frame11 
               Height          =   1095
               Left            =   3840
               TabIndex        =   46
               Top             =   120
               Width           =   1695
               Begin VB.CommandButton CmdAnotaciones 
                  Caption         =   "Anotar"
                  DownPicture     =   "frmEmpleado.frx":15307
                  Height          =   375
                  Left            =   120
                  Picture         =   "frmEmpleado.frx":16DE9
                  TabIndex        =   48
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.CommandButton CmdCarnet 
                  Caption         =   "Carnet"
                  DownPicture     =   "frmEmpleado.frx":188CB
                  Height          =   375
                  Left            =   120
                  Picture         =   "frmEmpleado.frx":1A3AD
                  TabIndex        =   47
                  Top             =   600
                  Width           =   1455
               End
            End
            Begin MSDataListLib.DataCombo DBCCargo 
               Bindings        =   "frmEmpleado.frx":1BE8F
               Height          =   315
               Left            =   6480
               TabIndex        =   13
               Top             =   2880
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Cargo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DBCDepartamento 
               Bindings        =   "frmEmpleado.frx":1BEA6
               Height          =   315
               Left            =   6480
               TabIndex        =   12
               Top             =   2520
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "Departamento"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DBCodigoEmpleado 
               Bindings        =   "frmEmpleado.frx":1BEC4
               Height          =   315
               Left            =   1440
               TabIndex        =   206
               Top             =   1920
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "CodEmpleado1"
               Text            =   ""
            End
            Begin VB.Label Label41 
               Caption         =   "No Hijos"
               Height          =   255
               Left            =   5520
               TabIndex        =   217
               Top             =   3600
               Width           =   855
            End
            Begin VB.Label Label77 
               Caption         =   "Telefono:"
               Height          =   255
               Left            =   600
               TabIndex        =   215
               Top             =   4080
               Width           =   855
            End
            Begin VB.Label Label13 
               Caption         =   "Cuenta Banco"
               Height          =   255
               Left            =   5160
               TabIndex        =   195
               Top             =   4080
               Width           =   1215
            End
            Begin VB.Label Label56 
               Caption         =   "Segundo Apellido:"
               Height          =   255
               Left            =   120
               TabIndex        =   71
               Top             =   3360
               Width           =   1335
            End
            Begin VB.Label Label55 
               Caption         =   "Primer Apellido:"
               Height          =   375
               Left            =   240
               TabIndex        =   70
               Top             =   3000
               Width           =   1095
            End
            Begin VB.Label Label54 
               Caption         =   "Segundo Nombre:"
               Height          =   255
               Left            =   120
               TabIndex        =   69
               Top             =   2640
               Width           =   1455
            End
            Begin VB.Label Label3 
               Caption         =   "Direccion:"
               Height          =   255
               Left            =   600
               TabIndex        =   68
               Top             =   3720
               Width           =   975
            End
            Begin VB.Label Label2 
               Caption         =   "Primer Nombre:"
               Height          =   255
               Left            =   240
               TabIndex        =   67
               Top             =   2280
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "Numero :"
               Height          =   255
               Left            =   600
               TabIndex        =   66
               Top             =   1920
               Width           =   1215
            End
            Begin VB.Image Image1 
               BorderStyle     =   1  'Fixed Single
               Height          =   1335
               Left            =   240
               Stretch         =   -1  'True
               Top             =   240
               Width           =   1725
            End
            Begin VB.Label Label14 
               Caption         =   "Sindicalista:"
               Height          =   255
               Left            =   5400
               TabIndex        =   65
               Top             =   3240
               Width           =   1095
            End
            Begin VB.Label Label12 
               Caption         =   "Cargo:"
               Height          =   255
               Left            =   5640
               TabIndex        =   64
               Top             =   2880
               Width           =   735
            End
            Begin VB.Label Label10 
               Caption         =   "Depto:"
               Height          =   255
               Left            =   5640
               TabIndex        =   63
               Top             =   2520
               Width           =   735
            End
            Begin VB.Label Label9 
               Caption         =   "INSS #"
               Height          =   255
               Left            =   5760
               TabIndex        =   62
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label Label8 
               Caption         =   "R.U.C. #"
               Height          =   255
               Left            =   5640
               TabIndex        =   61
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label Label6 
               Caption         =   "Sexo:"
               Height          =   255
               Left            =   6720
               TabIndex        =   60
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label5 
               Caption         =   "Cód.Postal:"
               Height          =   255
               Left            =   6480
               TabIndex        =   59
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label4 
               Caption         =   "Nacionalidad:"
               Height          =   255
               Left            =   6000
               TabIndex        =   58
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label45 
               Caption         =   "Cédula #"
               Height          =   375
               Left            =   5640
               TabIndex        =   57
               Top             =   2040
               Width           =   735
            End
            Begin VB.Label LblSuspendido 
               Caption         =   "SUBSIDIO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   3360
               TabIndex        =   56
               Top             =   1200
               Visible         =   0   'False
               Width           =   1455
            End
         End
         Begin VB.CommandButton CmdEliminarDeduccion 
            Caption         =   "Borrar"
            DownPicture     =   "frmEmpleado.frx":1BEDF
            Enabled         =   0   'False
            Height          =   375
            Left            =   -71400
            Picture         =   "frmEmpleado.frx":1D9C1
            TabIndex        =   31
            Top             =   3540
            Width           =   1455
         End
         Begin VB.CommandButton CmdEliminarIncentivo 
            Caption         =   "Borrar"
            DownPicture     =   "frmEmpleado.frx":1F4A3
            Enabled         =   0   'False
            Height          =   375
            Left            =   -71160
            Picture         =   "frmEmpleado.frx":20F85
            TabIndex        =   30
            Top             =   3660
            Width           =   1455
         End
         Begin VB.CommandButton CmdEliminarSubsidio 
            Caption         =   "Borrar"
            DownPicture     =   "frmEmpleado.frx":22A67
            Enabled         =   0   'False
            Height          =   375
            Left            =   -71160
            Picture         =   "frmEmpleado.frx":24549
            TabIndex        =   29
            Top             =   3660
            Width           =   1455
         End
         Begin VB.CommandButton CmdHistoIncentivos 
            Caption         =   "HISTORICO"
            DownPicture     =   "frmEmpleado.frx":2602B
            Height          =   375
            Left            =   -74880
            Picture         =   "frmEmpleado.frx":27B0D
            TabIndex        =   28
            Top             =   3660
            Visible         =   0   'False
            Width           =   1455
         End
         Begin TrueOleDBGrid70.TDBGrid DbgrSubsidios 
            Bindings        =   "frmEmpleado.frx":295EF
            Height          =   2655
            Left            =   -71280
            TabIndex        =   24
            Top             =   900
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   4683
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   14215660
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid DbgDeducciones 
            Bindings        =   "frmEmpleado.frx":29610
            Height          =   2535
            Left            =   -71400
            TabIndex        =   25
            Top             =   900
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   4471
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   14215660
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=13,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid DbGIncentivos 
            Bindings        =   "frmEmpleado.frx":29632
            Height          =   2655
            Left            =   -71280
            TabIndex        =   26
            Top             =   900
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   4683
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   14215660
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid DbgrLibreta 
            Bindings        =   "frmEmpleado.frx":29654
            Height          =   1575
            Left            =   -74640
            TabIndex        =   27
            Top             =   3300
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   2778
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   14215660
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=101,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   2295
            Left            =   -69600
            TabIndex        =   32
            Top             =   780
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   4048
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Préstamo"
            TabPicture(0)   =   "frmEmpleado.frx":29671
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "CmdAfectuar"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "CmdEditarPresLinea"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "CmdCancelarPrestamo"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "CmdEstadoPrestamo"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "Exportar"
            TabPicture(1)   =   "frmEmpleado.frx":2968D
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label60"
            Tab(1).Control(1)=   "Label65"
            Tab(1).Control(2)=   "Label61"
            Tab(1).Control(3)=   "CmdRuta"
            Tab(1).Control(4)=   "TxtRuta"
            Tab(1).Control(5)=   "CmdExportar"
            Tab(1).Control(6)=   "TxtDebitoPrestamo"
            Tab(1).Control(7)=   "TxtCreditoPrestamo"
            Tab(1).ControlCount=   8
            Begin VB.TextBox TxtCreditoPrestamo 
               Height          =   375
               Left            =   -73560
               MaxLength       =   25
               TabIndex        =   41
               Top             =   960
               Width           =   1935
            End
            Begin VB.TextBox TxtDebitoPrestamo 
               Height          =   375
               Left            =   -73560
               MaxLength       =   25
               TabIndex        =   40
               Top             =   480
               Width           =   1935
            End
            Begin VB.CommandButton CmdEstadoPrestamo 
               DownPicture     =   "frmEmpleado.frx":296A9
               Height          =   375
               Left            =   600
               Picture         =   "frmEmpleado.frx":2C4AB
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   1560
               Width           =   2415
            End
            Begin VB.CommandButton CmdCancelarPrestamo 
               DownPicture     =   "frmEmpleado.frx":2F2AD
               Height          =   375
               Left            =   600
               Picture         =   "frmEmpleado.frx":320AF
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   1200
               Width           =   2415
            End
            Begin VB.CommandButton CmdEditarPresLinea 
               DownPicture     =   "frmEmpleado.frx":34EB1
               Height          =   375
               Left            =   600
               Picture         =   "frmEmpleado.frx":37CB3
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   840
               Width           =   2415
            End
            Begin VB.CommandButton CmdAfectuar 
               DownPicture     =   "frmEmpleado.frx":3AAB5
               Height          =   375
               Left            =   600
               Picture         =   "frmEmpleado.frx":3D8B7
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   480
               Width           =   2415
            End
            Begin VB.CommandButton CmdExportar 
               DownPicture     =   "frmEmpleado.frx":40291
               Height          =   375
               Left            =   -74280
               Picture         =   "frmEmpleado.frx":43093
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   1440
               Width           =   2415
            End
            Begin VB.TextBox TxtRuta 
               Height          =   285
               Left            =   -74400
               TabIndex        =   34
               Top             =   1920
               Width           =   2415
            End
            Begin VB.CommandButton CmdRuta 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   -72000
               Picture         =   "frmEmpleado.frx":45DD5
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   1920
               Width           =   375
            End
            Begin VB.Label Label61 
               Caption         =   "Cuenta Debito:"
               Height          =   255
               Left            =   -74880
               TabIndex        =   44
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label65 
               Caption         =   "Cuenta Credito:"
               Height          =   255
               Left            =   -74880
               TabIndex        =   43
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label60 
               Caption         =   "Ruta"
               Height          =   255
               Left            =   -74880
               TabIndex        =   42
               Top             =   1920
               Width           =   495
            End
         End
         Begin VB.Label Label74 
            Caption         =   "Cuenta Sueldos:"
            Height          =   255
            Left            =   -72000
            TabIndex        =   191
            Top             =   3660
            Width           =   1335
         End
         Begin VB.Label Label72 
            Caption         =   "Cuenta INSS"
            Height          =   255
            Left            =   -68760
            TabIndex        =   187
            Top             =   3660
            Width           =   975
         End
         Begin VB.Label Label67 
            Alignment       =   2  'Center
            Caption         =   "Nuevos Creditos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   255
            Left            =   -74040
            TabIndex        =   142
            Top             =   900
            Width           =   3975
         End
      End
   End
   Begin MSAdodcLib.Adodc DtaEmpleados 
      Height          =   375
      Left            =   6720
      Top             =   8520
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaEmpleados"
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
   Begin MSAdodcLib.Adodc DtaHorarioEmpleado 
      Height          =   375
      Left            =   3720
      Top             =   8400
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
      Caption         =   "DtaHorarioEmpleado"
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
   Begin MSAdodcLib.Adodc DtaTurnos 
      Height          =   375
      Left            =   240
      Top             =   9360
      Width           =   2895
      _ExtentX        =   5106
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaTurnos"
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
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   330
      Left            =   3600
      Top             =   11160
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaConsulta"
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
   Begin MSAdodcLib.Adodc DtaDetalleIncentivo 
      Height          =   330
      Left            =   360
      Top             =   9840
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaDetalleIncentivo"
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
   Begin MSAdodcLib.Adodc DtaSuspenciones 
      Height          =   330
      Left            =   480
      Top             =   8640
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaSuspenciones"
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
   Begin MSAdodcLib.Adodc DtaMovPrestamo 
      Height          =   375
      Left            =   360
      Top             =   11280
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaMovPrestamo"
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
   Begin MSAdodcLib.Adodc DtaPrestamo 
      Height          =   375
      Left            =   360
      Top             =   10800
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaPrestamo"
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
   Begin MSAdodcLib.Adodc DtadetalleSubsidio2 
      Height          =   375
      Left            =   3600
      Top             =   10680
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
      Caption         =   "DtadetalleSubsidio2"
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
   Begin MSAdodcLib.Adodc DtaDetalleSubsidio 
      Height          =   375
      Left            =   6600
      Top             =   10680
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaDetalleSubsidio"
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
   Begin MSAdodcLib.Adodc DtaSubsidio 
      Height          =   375
      Left            =   3600
      Top             =   10320
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
      Caption         =   "DtaSubsidio"
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
   Begin MSAdodcLib.Adodc DtaTipoSubsidio 
      Height          =   375
      Left            =   360
      Top             =   10440
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaTipoSubsidio"
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
   Begin MSAdodcLib.Adodc DtaDetalleDeduccion 
      Height          =   375
      Left            =   6600
      Top             =   10320
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaDetalleDeduccion"
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
   Begin MSAdodcLib.Adodc DtaDetalleDeduccion2 
      Height          =   375
      Left            =   3600
      Top             =   9960
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
      Caption         =   "DtaDetalleDeduccion2"
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
   Begin MSAdodcLib.Adodc DtaDeduccion 
      Height          =   375
      Left            =   6600
      Top             =   9960
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaDeduccion"
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
   Begin MSAdodcLib.Adodc DtaTipoDeduccion 
      Height          =   375
      Left            =   360
      Top             =   10080
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaTipoDeduccion"
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
   Begin MSAdodcLib.Adodc DtaDetalleIncentivo2 
      Height          =   375
      Left            =   6600
      Top             =   11160
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaDetalleIncentivo"
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
   Begin MSAdodcLib.Adodc DtaIncentivo 
      Height          =   375
      Left            =   6720
      Top             =   8880
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaIncentivo"
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
   Begin MSAdodcLib.Adodc DtaTipoIncentivo 
      Height          =   375
      Left            =   3840
      Top             =   8760
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
      Caption         =   "DtaTipoIncentivo"
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
   Begin MSAdodcLib.Adodc DtaCargo 
      Height          =   375
      Left            =   3600
      Top             =   7560
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaCargo"
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
   Begin MSAdodcLib.Adodc DtaGrupo 
      Height          =   375
      Left            =   6720
      Top             =   8040
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaGrupo"
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
   Begin MSAdodcLib.Adodc DtaHistorico 
      Height          =   375
      Left            =   480
      Top             =   7440
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
      Caption         =   "DtaHistorico"
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
   Begin MSAdodcLib.Adodc DtaIncapacidades 
      Height          =   375
      Left            =   240
      Top             =   9000
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaIncapacidades"
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
      Left            =   6720
      Top             =   7320
      Width           =   2895
      _ExtentX        =   5106
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
   Begin MSAdodcLib.Adodc DtaTipoNomina 
      Height          =   375
      Left            =   6720
      Top             =   7680
      Width           =   2895
      _ExtentX        =   5106
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
   Begin MSAdodcLib.Adodc DtaInfNomina 
      Height          =   375
      Left            =   3720
      Top             =   9120
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
      Caption         =   "DtaInfNomina"
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
   Begin MSAdodcLib.Adodc DtaConsecutivos 
      Height          =   375
      Left            =   3600
      Top             =   8040
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
      Caption         =   "DtaConsecutivos"
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
   Begin MSAdodcLib.Adodc DtaDepartamento 
      Height          =   375
      Left            =   480
      Top             =   8280
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaDepartamento"
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
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "Ultimo"
      DownPicture     =   "frmEmpleado.frx":45F23
      Height          =   375
      Left            =   1680
      MouseIcon       =   "frmEmpleado.frx":47A05
      MousePointer    =   99  'Custom
      Picture         =   "frmEmpleado.frx":47E47
      TabIndex        =   20
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton CmdPrimero 
      Caption         =   "Primero"
      DownPicture     =   "frmEmpleado.frx":49929
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmEmpleado.frx":4B40B
      MousePointer    =   99  'Custom
      Picture         =   "frmEmpleado.frx":4B84D
      TabIndex        =   19
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      DownPicture     =   "frmEmpleado.frx":4D32F
      Height          =   375
      Left            =   1680
      MouseIcon       =   "frmEmpleado.frx":4EE11
      MousePointer    =   99  'Custom
      Picture         =   "frmEmpleado.frx":4F253
      TabIndex        =   18
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "Anterior"
      DownPicture     =   "frmEmpleado.frx":50D35
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmEmpleado.frx":52817
      MousePointer    =   99  'Custom
      Picture         =   "frmEmpleado.frx":52C59
      TabIndex        =   17
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      DownPicture     =   "frmEmpleado.frx":5473B
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      MouseIcon       =   "frmEmpleado.frx":5621D
      MousePointer    =   99  'Custom
      Picture         =   "frmEmpleado.frx":5665F
      TabIndex        =   21
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Salir"
      DownPicture     =   "frmEmpleado.frx":58141
      Height          =   375
      Left            =   8160
      MouseIcon       =   "frmEmpleado.frx":59C23
      MousePointer    =   99  'Custom
      Picture         =   "frmEmpleado.frx":5A065
      TabIndex        =   22
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      DownPicture     =   "frmEmpleado.frx":5BB47
      Height          =   375
      Left            =   3480
      MouseIcon       =   "frmEmpleado.frx":5D629
      MousePointer    =   99  'Custom
      Picture         =   "frmEmpleado.frx":5DA6B
      TabIndex        =   16
      Top             =   6000
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc AdoUser 
      Height          =   375
      Left            =   240
      Top             =   7200
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "AdoUser"
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
   Begin VB.Label Label76 
      Caption         =   "Prov Aguinaldo:"
      Height          =   255
      Left            =   3120
      TabIndex        =   192
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "frmEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodTipoNomina As String
Public Bandera As Boolean, CodEmpleado As Double, SueldoActual As Double

Private Sub ChkSueldoActual_Click()
res = Bitacora(Now, NombreUsuario, "Empleados", "Se cambio el Estatus Sueldo Actual: " & DBCodigoEmpleado.Text & " " & Me.TxtNombre1.Text)
End Sub

Private Sub CmdAnular_Click()
On Error GoTo TipoErrs
Dim NumIncentivo As Integer

NumIncentivo = Me.DbGIncentivos.Columns(0).Text

Me.DtaConsulta.RecordSource = "SELECT  * From DetalleIncentivo Where (NumIncentivo =" & NumIncentivo & " )"
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
  Me.DtaConsulta.Recordset("Valor") = 0
  Me.DtaConsulta.Recordset.Update
End If
Me.DtaDetalleIncentivo.Refresh

Me.CmdAnular.Enabled = False

DbGIncentivos.Columns(0).Visible = False
DbGIncentivos.Columns(2).Visible = False
DbGIncentivos.Columns(5).Visible = False
Exit Sub
TipoErrs:
  ControlErrores
  Unload Me

End Sub

Private Sub Check1_Click()
 If Me.Check1.Value = 0 Then
    Me.TxtPorcientoHora.Visible = False
    Me.LblPor.Visible = False
 Else
    Me.TxtPorcientoHora.Visible = True
    Me.LblPor.Visible = True
 End If
End Sub

Private Sub ChkSuspendido_Click()
On Error GoTo TipoErr
Dim NumeroEmpleado As Integer
Dim rs As New ADODB.Recordset
Dim Respuesta As Integer

If Bandera Then
        If DBCodigoEmpleado = "" Then
           MsgBox "No has seleccionado al empleado"
           ChkSuspendido.Value = 0
           Exit Sub
        End If
        If ChkSuspendido.Value = 1 Then
           k% = MsgBox("Desea poner en Subsidio al Empleado", vbYesNo)
           If k <> 6 Then
             ChkSuspendido.Value = 0
             Exit Sub
          
           End If
           
             NumeroEmpleado = Me.TxtCodEmpleado.Text
             sql = "SELECT DetalleNomina.NumNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HE,DetalleNomina.DD, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre,DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR," & vbLf
             sql = sql & "DetalleNomina.Vacaciones, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.INATEC, DetalleNomina.Mes13,DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.TotalSubsidio, DetalleNomina.VacacionesPagadas," & vbLf
             sql = sql & "DetalleNomina.DiasVacaciones, DetalleNomina.AdelantosVacaciones, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia," & vbLf
             sql = sql & "DetalleNomina.IncetivoProduccion , DetalleNomina.TarifaHoraria, Nomina.Activa" & vbLf
             sql = sql & "FROM  DetalleNomina INNER JOIN  Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina Where (DetalleNomina.CodEmpleado = " & NumeroEmpleado & " ) And (Nomina.Activa = 1)"
             Me.DtaConsulta.RecordSource = sql
             Me.DtaConsulta.Refresh
             Do While Not Me.DtaConsulta.Recordset.EOF
'               Me.DtaConsulta.Recordset.Delete
               Me.DtaConsulta.Recordset.MoveNext
             Loop
           LblSuspendido.Visible = True
           FrmSuspencion.TxtCodEmpleado.Text = DBCodigoEmpleado.Text
           FrmSuspencion.txtNombre = TxtNombre1.Text + " " + TxtNombre2.Text + " " + TxtApellido1 + " " + TxtApellido2
           FrmSuspencion.Show 1
           
        Else
           LblSuspendido.Visible = False
        End If
        
        
End If

 If Me.ChkSuspendido.Value = 0 Then
 
   NumeroEmpleado = Me.TxtCodEmpleado.Text
   Respuesta = MsgBox("¿Desea Eliminar el subsidio?", vbYesNo, "Zeus Nominas")
   If Respuesta = 6 Then
    rs.Open "DELETE FROM Subsidios Where (CodEmpleado = " & NumeroEmpleado & " ) And (Activo = 1)", Conexion
    rs.Open "UPDATE Empleado Set [Ausente] = 'False'  Where (CodEmpleado = " & NumeroEmpleado & " )"
   
   End If
   
 End If

Exit Sub
TipoErr:
  'ControlErrores
End Sub

Private Sub CmbExentoInss_Click()
Salida = True
End Sub

Private Sub CmbExentoInss_Change()
PreparaSalida
End Sub

Private Sub CmbExentoInss_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   CmbExentoIr.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub CmbExentoIr_Click()
Salida = True
End Sub

Private Sub CmbExentoIr_Change()
PreparaSalida
End Sub

Private Sub CmbExentoIr_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   CmdGrabar.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub CmbMonedaComision_Click()
Salida = True
End Sub

Private Sub CmbMonedaComision_Change()
PreparaSalida
End Sub

Private Sub CmbMonedaComision_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   CmbPagoInssPatronal.SetFocus
Else
   Evaluar = False
  End If
End Sub

Private Sub CmbMonedaSueldo_Click()
Salida = True
End Sub

Private Sub CmbMonedaSueldo_Change()
PreparaSalida
End Sub

Private Sub CmbMonedaSueldo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   CmbMonedaComision.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub CmbPagoInssPatronal_Click()
Salida = True
End Sub

Private Sub CmbPagoInssPatronal_Change()
PreparaSalida
End Sub

Private Sub CmbPagoInssPatronal_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  CmbSalarioMinimo.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub CmbSalarioMinimo_Click()
Salida = True
End Sub

Private Sub CmbSalarioMinimo_Change()
PreparaSalida
End Sub

Private Sub CmbSalarioMinimo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  CmbExentoInss.SetFocus
Else
   Evaluar = False
  End If
End Sub



Private Sub CmbSexo_Click()
Salida = True
End Sub

Private Sub CmbSexo_Change()
Salida = True
End Sub

Private Sub CmbSexo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  TxtNRuc.SetFocus
 Else
   Evaluar = False
  End If
End Sub




Private Sub CmbSindicalista_Click()
Salida = True
End Sub

Private Sub CmbSindicalista_Change()
PreparaSalida
End Sub

Private Sub CmbSindicalista_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   CmdGrabar.SetFocus
  Else
   Evaluar = False
  End If
End Sub

Private Sub CmbTipoPago_Click()
Salida = True
  On Error GoTo TipoErrs
   If frmEmpleado.CmbTipoPago.Text = "Salario Fijo" Then
     frmEmpleado.TxtTarifaHoraria.Enabled = False
     frmEmpleado.TxtSueldoPeriodo.Enabled = True
     frmEmpleado.TxtTarifaHoraria.Text = "0.00"
     frmEmpleado.TxtComision.Text = "0.00"
    End If
If frmEmpleado.CmbTipoPago.Text = "Destajo" Then
  frmEmpleado.TxtTarifaHoraria.Enabled = True
  frmEmpleado.TxtSueldoPeriodo.Enabled = False
  frmEmpleado.TxtSueldoPeriodo.Text = "0.00"
  frmEmpleado.TxtComision.Text = "0.00"
End If
If frmEmpleado.CmbTipoPago.Text = "Comisiones" Then
  frmEmpleado.TxtTarifaHoraria.Enabled = False
  frmEmpleado.TxtSueldoPeriodo.Enabled = False
  frmEmpleado.TxtTarifaHoraria.Text = "0.00"
  frmEmpleado.TxtSueldoPeriodo.Text = "0.00"
End If
 If frmEmpleado.CmbTipoPago.Text = "Salario Fijo/Comisiones" Then
  frmEmpleado.TxtSueldoPeriodo.Enabled = True
  frmEmpleado.TxtTarifaHoraria.Enabled = False
  frmEmpleado.TxtTarifaHoraria.Text = "0.00"
End If

Exit Sub

TipoErrs:
    ControlErrores
End Sub

Private Sub CmbTipoPago_Change()
PreparaSalida
End Sub

Private Sub CmbTipoPago_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtSueldoPeriodo.Text = Format((TxtSueldoPeriodo.Text), "##,##0.00")
  TxtSueldoPeriodo.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub CmdAgregarIncentico_Click()

End Sub

Private Sub CmdAfectuar_Click()
On Error GoTo TipoErrs
Dim NumPrestamo As Double
Dim CuotaIgual As Double
Dim Saldo As Double
Dim Interes As Double
Dim CantCuotas As Long
Dim CuotaPrincipal As Double
Dim MontoTotalInteres As Double
Dim PagoInteres(100) As Double


DtaConsecutivos.Refresh
'DtaConsecutivos.Recordset.Edit
DtaConsecutivos.Recordset("prestamos") = DtaConsecutivos.Recordset("prestamos") + 1
DtaConsecutivos.Recordset.Update
'NumPrestamo = DtaConsecutivos.Recordset("prestamos")

Me.DtaConsulta.RecordSource = "SELECT     NumPrestamo, CuentaDebito, CuentaCredito, Monto, CantCuotas, Interes, Saldo, FechaInicial, Cancelado, Moneda, CuotasIguales, CodEmpleado From Prestamo "
Me.DtaConsulta.Refresh
If Me.DtaConsulta.Recordset.EOF Then
  NumPrestamo = 1
Else
  Me.DtaConsulta.Recordset.MoveLast
  NumPrestamo = Me.DtaConsulta.Recordset("NumPrestamo") + 1
End If

'pregunto si ha escogido un empleado

If DBCodigoEmpleado.Text = "" Then
    MsgBox "No ha seleccionado el empleado"
    Exit Sub
End If
'verifico si no tiene un prestamo sin cancelar

If TxtCreditoPrestamo.Text = "" Then
    MsgBox "La Cuenta de Crédito no ha sido digitada"
    Exit Sub
End If

If TxtDebitoPrestamo.Text = "" Then
    MsgBox "La Cuenta de Crédito no ha sido digitada"
    Exit Sub
End If


DtaPrestamo.Refresh
Do While Not DtaPrestamo.Recordset.EOF
If DtaPrestamo.Recordset("CodEmpleado") = val(Me.TxtCodEmpleado.Text) And DtaPrestamo.Recordset("cancelado") = False Then
 If Not val(DtaPrestamo.Recordset("Saldo")) = 0 Then
   MsgBox ("Este Empleado ya tiene un préstamo y su saldo es de: " & DtaPrestamo.Recordset("Saldo"))
   Exit Sub
 End If
End If
DtaPrestamo.Recordset.MoveNext
Loop

CantCuotas = val(TxtCuotas.Text)
Saldo = val(TxtSaldo.Text)
Interes = val(TxtInteresprestamo.Text) / 100
If Not CantCuotas = 0 Then
CuotaPrincipal = Saldo / CantCuotas
Else
 MsgBox "Digite la Cantidad de Cuotas", vbCritical, "sistema de nominas"
 Exit Sub
End If
MontoTotalInteres = 0

For i = 1 To CantCuotas
    PagoInteres(i) = (Saldo * Interes)
    MontoTotalInteres = MontoTotalInteres + (Saldo * Interes)
    Saldo = Saldo - CuotaPrincipal
Next i

CodEmpleado = Me.TxtCodEmpleado.Text

Saldo = val(Me.TxtSaldo.Text)
CuotaIgual = (Saldo + MontoTotalInteres) / CantCuotas


'agrego los datos a Prestamo
DtaPrestamo.Recordset.AddNew
DtaPrestamo.Recordset("NumPrestamo") = NumPrestamo
DtaPrestamo.Recordset("Monto") = val(TxtMontoPrestamoUS.Text)
DtaPrestamo.Recordset("CantCuotas") = val(TxtCuotas.Text)
DtaPrestamo.Recordset("Interes") = val(TxtInteresprestamo.Text)
DtaPrestamo.Recordset("Saldo") = val(TxtSaldo.Text)
DtaPrestamo.Recordset("fechainicial") = CDate(Now)
DtaPrestamo.Recordset("CuentaDebito") = TxtDebitoPrestamo.Text
DtaPrestamo.Recordset("cuentacredito") = TxtCreditoPrestamo.Text
DtaPrestamo.Recordset("CodEmpleado") = CodEmpleado
If ChkTipoPago = 1 Then
DtaPrestamo.Recordset("CuotasIguales") = True
Else
DtaPrestamo.Recordset("CuotasIguales") = False
End If

If OptC Then
   DtaPrestamo.Recordset("Moneda") = "CS"
Else
   DtaPrestamo.Recordset("Moneda") = "US"
End If

DtaPrestamo.Recordset.Update


'agrego los datos a Nprestamo
Saldo = val(TxtSaldo.Text)
For i = 1 To CantCuotas
    DtaMovPrestamo.Recordset.AddNew
    DtaMovPrestamo.Recordset("ID") = i
    DtaMovPrestamo.Recordset("NumPrestamo") = NumPrestamo
    'DtaMovPrestamo.Recordset("CodEmpleado") = DBCodigoEmpleado.Text
    DtaMovPrestamo.Recordset("numcuota") = i
    DtaMovPrestamo.Recordset("Monto") = CuotaPrincipal
    DtaMovPrestamo.Recordset("Interes") = PagoInteres(i)
    DtaMovPrestamo.Recordset("saldocuota") = Saldo
    DtaMovPrestamo.Recordset("CuotaIgual") = CuotaIgual
    DtaMovPrestamo.Recordset.Update
    Saldo = Saldo - CuotaPrincipal
Next i

DbgrLibreta.Columns(0).Visible = False
DbgrLibreta.Columns(1).Visible = False
DbgrLibreta.Columns(7).Visible = False

res = Bitacora(Now, NombreUsuario, "Empleados", "Se genero un Prestamo: " & DBCodigoEmpleado.Text & " " & Me.TxtNombre1.Text)

MsgBox "El préstamo fue realizado con éxito"



Exit Sub
TipoErrs:
ControlErrores

End Sub

Private Sub CmdAgregarDeduccion_Click()
'On Error GoTo TipoErrs
Dim CodDeduccion As String
Dim Numdeduccion As Double, NumeroNomina As Double
Dim CodTipoNomina As String, NUmDetalleDed As Double

CodTipoNomina = Me.TxtCodTipoNomina.Text

Me.DtaConsulta.RecordSource = "SELECT NumNomina, CodTipoNomina, Activa, Cerrada, Anulada From Nomina WHERE     (Activa = 1) AND (Cerrada = 0) AND (Anulada = 0) AND (CodTipoNomina = '" & CodTipoNomina & "') "
Me.DtaConsulta.Refresh
If Me.DtaConsulta.Recordset.EOF Then
  MsgBox "No Existe, Nomina Activa para este Empleado", vbCritical, "Sistema de Nominas"
  Exit Sub
Else
  NumeroNomina = Me.DtaConsulta.Recordset("NumNomina")
End If


If DbcDeducciones.Text = "Tipos de Deducciones" Then
    MsgBox "No ha Seleleccionado el tipo de Deduccion"
    DbcDeducciones.SetFocus
    Exit Sub
End If

If DBCodigoEmpleado.Text = "" Then
    MsgBox "No ha seleccionado el empleado"
    Exit Sub
End If

DtaConsecutivos.Refresh
'DtaConsecutivos.Recordset.Edit
DtaConsecutivos.Recordset("Deducciones") = DtaConsecutivos.Recordset("Deducciones") + 1
DtaConsecutivos.Recordset.Update
Numdeduccion = DtaConsecutivos.Recordset("Deducciones")



If TxtVecesDeduccion.Text <> "N" And TxtVecesDeduccion.Text <> "n" Then
    If Not IsNumeric(TxtVecesDeduccion.Text) Then
      MsgBox "El número de veces digitado es erróneo"
      Exit Sub
    End If
End If

If Not IsNumeric(TxtMontoDeduccion.Text) Then
   MsgBox "EL monto digitado es erróneo"
   TxtMontoDeduccion.SetFocus
   Exit Sub
End If


DtaTipoDeduccion.Refresh
Do While Not DtaTipoDeduccion.Recordset.EOF
If DtaTipoDeduccion.Recordset("deduccion") = DbcDeducciones.Text Then
   CodDeduccion = DtaTipoDeduccion.Recordset("codtipodeduccion")
   Exit Do
End If
DtaTipoDeduccion.Recordset.MoveNext
Loop

DtaDeduccion.RecordSource = "SELECT NumDeduccion, CodEmpleado, CodTipoDeduccion, NumVeces, Pagado, NUmNomina From Deduccion"
DtaDeduccion.Refresh
If DtaDeduccion.Recordset.EOF Then
 Numdeduccion = 0
Else
  Me.DtaDeduccion.Recordset.MoveLast
  Numdeduccion = Me.DtaDeduccion.Recordset("NumDeduccion") + 1

End If

DtaDeduccion.Recordset.AddNew
DtaDeduccion.Recordset("NumDeduccion") = Numdeduccion
DtaDeduccion.Recordset("CodEmpleado") = val(Me.TxtCodEmpleado.Text)
DtaDeduccion.Recordset("codtipodeduccion") = CodDeduccion
DtaDeduccion.Recordset("numveces") = TxtVecesDeduccion.Text
DtaDeduccion.Recordset("pagado") = False
DtaDeduccion.Recordset("NumNomina") = NumeroNomina
DtaDeduccion.Recordset.Update

Me.DtaConsulta.RecordSource = "SELECT Id, NumDeduccion, Valor, NumVez, Pagado, NumNomina From DetalleDeduccion"
Me.DtaConsulta.Refresh
If Me.DtaConsulta.Recordset.EOF Then
  NUmDetalleDed = 1
Else
 Me.DtaConsulta.Recordset.MoveLast
 NUmDetalleDed = Me.DtaConsulta.Recordset("Id") + 1

End If
DtaDetalleDeduccion2.Refresh

If Not IsNumeric(TxtVecesDeduccion.Text) Then

       DtaDetalleDeduccion2.Recordset.AddNew
       ' DtaDetalleDeduccion2.Recordset("ID") = Numdeduccion + 1
        DtaDetalleDeduccion2.Recordset("NumDeduccion") = Numdeduccion
        DtaDetalleDeduccion2.Recordset("valor") = val(TxtMontoDeduccion.Text)
        DtaDetalleDeduccion2.Recordset("NumVez") = 9999
        DtaDetalleDeduccion2.Recordset("pagado") = False
        DtaDetalleDeduccion2.Recordset("NumNomina") = NumeroNomina
        
        DtaDetalleDeduccion2.Recordset.Update
Else
For i = 1 To val(TxtVecesDeduccion.Text)
        DtaDetalleDeduccion2.Recordset.AddNew
        'DtaDetalleDeduccion2.Recordset("ID") = NUmDetalleDed
        DtaDetalleDeduccion2.Recordset("NumDeduccion") = Numdeduccion
        DtaDetalleDeduccion2.Recordset("valor") = val(TxtMontoDeduccion.Text)
        DtaDetalleDeduccion2.Recordset("NumVez") = i
        DtaDetalleDeduccion2.Recordset("pagado") = False
        If i = 1 Then
         DtaDetalleDeduccion2.Recordset("NumNomina") = NumeroNomina
        Else
         DtaDetalleDeduccion2.Recordset("NumNomina") = 0
        End If
             
        DtaDetalleDeduccion2.Recordset.Update
        NUmDetalleDed = NUmDetalleDed + 1
Next
End If
DtaDetalleDeduccion.Refresh
DbgDeducciones.Columns(0).Visible = False
DbgDeducciones.Columns(2).Visible = False
'DbgDeducciones.Columns(5).Visible = False


res = Bitacora(Now, NombreUsuario, "Empleados", "Se Agrego una Deduccion: " & DBCodigoEmpleado.Text & " " & Me.TxtNombre1.Text & " Por: " & val(TxtMontoDeduccion.Text) & " para la nomina: " & NumeroNomina)


Exit Sub
TipoErrs:
 ControlErrores
 Unload Me


End Sub

Private Sub CmdAgregarIncentivo_Click()

On Error GoTo TipoErrs

Dim CodIncentivo As String
Dim NumIncentivo As Long

If DBCodigoEmpleado.Text = "" Then
    MsgBox "No ha seleccionado el empleado"
    Exit Sub
End If


If DbCTipoIncentivo.Text = "Tipos de Incentivos" Then
    MsgBox "No ha Seleleccionado el tipo de Incentivo"
    DbCTipoIncentivo.SetFocus
    Exit Sub
End If

'DtaConsecutivos.Refresh
''DtaConsecutivos.Recordset.Edit
'DtaConsecutivos.Recordset("incentivos") = DtaConsecutivos.Recordset("incentivos") + 1
'DtaConsecutivos.Recordset.Update
'Numincentivo = DtaConsecutivos.Recordset("incentivos")



If TxtNumVeces.Text <> "N" And TxtNumVeces.Text <> "n" Then
    If Not IsNumeric(TxtNumVeces.Text) Then
      MsgBox "El número de veces digitado es erróneo"
      Exit Sub
    End If
End If

If Not IsNumeric(TxtMonto.Text) Then
   MsgBox "EL monto digitado es erróneo"
   TxtMonto.SetFocus
   Exit Sub
End If

CodTipoNomina = Me.TxtCodTipoNomina.Text

Me.DtaConsulta.RecordSource = "SELECT NumNomina, CodTipoNomina, Activa, Cerrada, Anulada From Nomina WHERE     (Activa = 1) AND (Cerrada = 0) AND (Anulada = 0) AND (CodTipoNomina = '" & CodTipoNomina & "') "
Me.DtaConsulta.Refresh
If Me.DtaConsulta.Recordset.EOF Then
  MsgBox "No Existe, Nomina Activa para este Empleado", vbCritical, "Sistema de Nominas"
  Exit Sub
Else
  NumeroNomina = Me.DtaConsulta.Recordset("NumNomina")
End If



DtaTipoIncentivo.Refresh
Do While Not DtaTipoIncentivo.Recordset.EOF
If DtaTipoIncentivo.Recordset("incentivo") = DbCTipoIncentivo.Text Then
   CodIncentivo = DtaTipoIncentivo.Recordset("CodtipoIncentivo")
   Exit Do
End If
DtaTipoIncentivo.Recordset.MoveNext
Loop


Me.DtaIncentivo.RecordSource = "SELECT NumIncentivo, CodEmpleado, CodTipoIncentivo, NumVeces, Pagado From Incentivo"
Me.DtaIncentivo.Refresh
If DtaIncentivo.Recordset.EOF Then
  NumIncentivo = 0
Else
 Me.DtaIncentivo.Recordset.MoveLast
 NumIncentivo = Me.DtaIncentivo.Recordset("NumIncentivo") + 1
End If

Me.DtaIncentivo.Recordset.AddNew
DtaIncentivo.Recordset("NumIncentivo") = NumIncentivo
DtaIncentivo.Recordset("CodEmpleado") = val(Me.TxtCodEmpleado.Text)
DtaIncentivo.Recordset("CodtipoIncentivo") = CodIncentivo
DtaIncentivo.Recordset("numveces") = TxtNumVeces.Text
DtaIncentivo.Recordset("pagado") = False
DtaIncentivo.Recordset.Update

 If Me.TxtNumVeces.Text = "n" Then
            DtaDetalleIncentivo2.Recordset.AddNew
                DtaDetalleIncentivo2.Recordset("ID") = i
                DtaDetalleIncentivo2.Recordset("NumIncentivo") = NumIncentivo
                DtaDetalleIncentivo2.Recordset("valor") = val(TxtMonto.Text)
                DtaDetalleIncentivo2.Recordset("NumVez") = "n"
                DtaDetalleIncentivo2.Recordset("NumNomina") = NumeroNomina
                DtaDetalleIncentivo2.Recordset("pagado") = False
            DtaDetalleIncentivo2.Recordset.Update
 Else

        Me.DtaDetalleIncentivo2.Refresh
        If Not IsNumeric(TxtNumVeces.Text) Then
            DtaDetalleIncentivo2.Recordset.AddNew
                DtaDetalleIncentivo2.Recordset("ID") = i
                DtaDetalleIncentivo2.Recordset("NumIncentivo") = NumIncentivo
                DtaDetalleIncentivo2.Recordset("valor") = val(TxtMonto.Text)
                DtaDetalleIncentivo2.Recordset("NumVez") = i
                DtaDetalleIncentivo2.Recordset("NumNomina") = NumeroNomina
                DtaDetalleIncentivo2.Recordset("pagado") = False
            DtaDetalleIncentivo2.Recordset.Update
        Else
       
          For i = 1 To val(TxtNumVeces.Text)
            
                DtaDetalleIncentivo2.Recordset.AddNew
                    DtaDetalleIncentivo2.Recordset("ID") = i
                    DtaDetalleIncentivo2.Recordset("NumIncentivo") = NumIncentivo
                    DtaDetalleIncentivo2.Recordset("valor") = val(TxtMonto.Text)
                    DtaDetalleIncentivo2.Recordset("NumVez") = i
                    If i = 1 Then
                     DtaDetalleIncentivo2.Recordset("NumNomina") = NumeroNomina
                    Else
                     DtaDetalleIncentivo2.Recordset("NumNomina") = 0
                    End If
                    DtaDetalleIncentivo2.Recordset("pagado") = False
                DtaDetalleIncentivo2.Recordset.Update
            Next
        End If
        
  End If
  
  
DtaDetalleIncentivo.Refresh

DbGIncentivos.Columns(0).Visible = False
DbGIncentivos.Columns(2).Visible = False
DbGIncentivos.Columns(5).Visible = False


res = Bitacora(Now, NombreUsuario, "Empleados", "Se Agrego un Incentivo: " & DBCodigoEmpleado.Text & " " & Me.TxtNombre1.Text & " Por: " & val(TxtMonto.Text) & " para la nomina: " & NumeroNomina)

Exit Sub
TipoErrs:
ControlErrores
Unload Me


End Sub

Private Sub CmdAgregarSubsidio_Click()
'On Error GoTo TipoErrs

Dim CodSubsidio As String
Dim NumSubsidio As Long



If DBCTipoSubsidio.Text = "Tipos de Subsidios" Then
    MsgBox "No ha Seleleccionado el tipo de Subsidio"
    DBCTipoSubsidio.SetFocus
    Exit Sub
End If

If DBCodigoEmpleado.Text = "" Then
    MsgBox "No ha seleccionado el empleado"
    Exit Sub
End If

'DtaConsecutivos.Refresh
''DtaConsecutivos.Recordset.Edit
'DtaConsecutivos.Recordset("Subsidios") = DtaConsecutivos.Recordset("Subsidios") + 1
'DtaConsecutivos.Recordset.Update

NumSubsidio = ConsecutivoSubsidio("Subsidio")


If TxtNumVecesSubsidio.Text <> "N" And TxtNumVecesSubsidio.Text <> "n" Then
    If Not IsNumeric(TxtNumVecesSubsidio.Text) Then
      MsgBox "El número de veces digitado es erróneo"
      Exit Sub
    End If
End If

If Not IsNumeric(TxtMontoSubsidio.Text) Then
   MsgBox "EL monto digitado es erróneo"
   TxtMontoSubsidio.SetFocus
   Exit Sub
End If


DtaTipoSubsidio.Refresh
Do While Not DtaTipoSubsidio.Recordset.EOF
If DtaTipoSubsidio.Recordset("Subsidio") = DBCTipoSubsidio.Text Then
   CodSubsidio = DtaTipoSubsidio.Recordset("CodtipoSubsidio")
   Exit Do
End If
DtaTipoSubsidio.Recordset.MoveNext
Loop
Me.DtaSubsidio.Refresh
DtaSubsidio.Recordset.AddNew
DtaSubsidio.Recordset("NumSubsidio") = NumSubsidio
DtaSubsidio.Recordset("CodEmpleado") = Me.TxtCodEmpleado.Text
DtaSubsidio.Recordset("CodtipoSubsidio") = CodSubsidio
DtaSubsidio.Recordset("numveces") = TxtNumVecesSubsidio.Text
DtaSubsidio.Recordset("pagado") = False
DtaSubsidio.Recordset.Update

Me.DtadetalleSubsidio2.Refresh
If Not IsNumeric(TxtNumVecesSubsidio.Text) Then
       DtadetalleSubsidio2.Recordset.MoveLast
 
        DtadetalleSubsidio2.Recordset.AddNew
        DtadetalleSubsidio2.Recordset("ID") = NumSubsidio
        DtadetalleSubsidio2.Recordset("NumSubsidio") = NumSubsidio
        DtadetalleSubsidio2.Recordset("valor") = val(TxtMontoSubsidio.Text)
        If Not IsNull(TxtDescripcion.Text) Then
            DtadetalleSubsidio2.Recordset("Descripcion") = TxtDescripcion.Text
        Else
            DtadetalleSubsidio2.Recordset("Descripcion") = " "
        End If
        DtadetalleSubsidio2.Recordset("NumVez") = "n"
        DtadetalleSubsidio2.Recordset("pagado") = False
        DtadetalleSubsidio2.Recordset.Update
Else
'For i = 1 To val(TxtNumVecesSubsidio.Text)
       DtadetalleSubsidio2.Recordset.AddNew
        DtadetalleSubsidio2.Recordset("ID") = NumSubsidio
        DtadetalleSubsidio2.Recordset("NumSubsidio") = NumSubsidio
        DtadetalleSubsidio2.Recordset("valor") = val(TxtMontoSubsidio.Text)
        If Not IsNull(TxtDescripcion.Text) Then
            DtadetalleSubsidio2.Recordset("Descripcion") = TxtDescripcion.Text
        Else
            DtadetalleSubsidio2.Recordset("Descripcion") = " "
        End If
        DtadetalleSubsidio2.Recordset("NumVez") = TxtNumVecesSubsidio.Text
        DtadetalleSubsidio2.Recordset("pagado") = False
        DtadetalleSubsidio2.Recordset.Update
'Next
End If
DtaDetalleSubsidio.Refresh

DbgrSubsidios.Columns(0).Visible = False
DbgrSubsidios.Columns(1).Visible = False
DbgrSubsidios.Columns(2).Visible = False
DbgrSubsidios.Columns(7).Visible = False
DbgrSubsidios.Columns(5).Width = 1200
DbgrSubsidios.Columns(6).Width = 500


Exit Sub
TipoErrs:
 ControlErrores
 Unload Me

End Sub

Private Sub CmdAnterior_Click()
' On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Empleado")
If Contesta Then
  CmdGrabar.Value = True
End If
  DtaEmpleados.Recordset.MovePrevious
       If DtaEmpleados.Recordset.BOF Then
           DtaEmpleados.Recordset.MoveNext
           MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
       Else
           DBCodigoEmpleado.Text = DtaEmpleados.Recordset("CodEmpleado1")
           Llenado
       End If
 
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub cmdborrar_Click()
On Error GoTo TipoErrs
If IsNull(DtaEmpleado.Recordset("CodEmpleado")) Then
 MsgBox "No existen Registros", vbInformation, "Sistema de Nominas"
 Exit Sub
End If
Dim Respuesta, Rsp
'Elimino el registro activo en la pantalla
  Set Rsp = DtaEmpleado.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?, Si Borra al Empleado puede que Descuadre las nóminas Existentes", vbYesNo, "Borrando a el Empleado: " & TxtNombre1.Text)
   If Respuesta = 6 Then
     Rsp.Delete
       DBCodigoEmpleado.Text = ""
      DtaEmpleado.Recordset.MoveLast
      DtaEmpleado.Recordset.MovePrevious
  End If


With Me.DtaEmpleados
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT CodEmpleado, CodEmpleado1, Activo, Nombre1 + ' '+ Nombre2 +' '+Apellido1+' '+Apellido2 as Nombres From Empleado Where (Activo = 1) ORDER BY CodEmpleado1"
   .Refresh
End With
'Me.DBCodigoEmpleado.Columns(0).Visible = False
'Me.DBCodigoEmpleado.Columns(1).Caption = "Codigo"
'Me.DBCodigoEmpleado.Columns(1).Width = 800
'Me.DBCodigoEmpleado.Columns(2).Visible = False

Salida = False
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdBuscarEmpleado_Click()
QueProducto = "CodigoEmpleado"
FrmConsulta.Show 1
Me.DBCodigoEmpleado.Text = FrmConsulta.CodigoEmpleado1
DBCodigoEmpleado_Change

'FrmBuscaEmpleado.Show 1
End Sub

Private Sub CmdCancelarPrestamo_Click()

On Error GoTo TipoErrs

k% = MsgBox("Desea Realmente Borrar el préstamo?", vbYesNo)
If k% <> 6 Then
   Exit Sub
End If

SqlDetallePrestamo = "SELECT MovPrestamo.NumPrestamo, Prestamo.CodEmpleado, MovPrestamo.NumCuota, MovPrestamo.Monto, MovPrestamo.Interes, MovPrestamo.CuotaIgual, MovPrestamo.SaldoCuota,Movprestamo.cancelado FROM Prestamo INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo WHERE Prestamo.CodEmpleado='" & Me.TxtCodEmpleado.Text & "' AND MovPrestamo.Cancelado = 0"
DtaMovPrestamo.RecordSource = SqlDetallePrestamo
DtaMovPrestamo.Refresh

SQlPrestamo = "SELECT Prestamo.* From Prestamo WHERE Prestamo.CodEmpleado='" & Me.TxtCodEmpleado.Text & "' AND Prestamo.Cancelado=0"
DtaPrestamo.RecordSource = SQlPrestamo
DtaPrestamo.Refresh


Do While Not DtaPrestamo.Recordset.EOF
 'DtaPrestamo.Recordset.Edit
  DtaPrestamo.Recordset("cancelado") = 1
 DtaPrestamo.Recordset.Update

  DtaPrestamo.Recordset.MoveNext
Loop

Do While Not DtaMovPrestamo.Recordset.EOF
 'DtaMovPrestamo.Recordset.Edit
  DtaMovPrestamo.Recordset("cancelado") = 1
 DtaMovPrestamo.Recordset.Update

 DtaMovPrestamo.Recordset.MoveNext
Loop

DtaPrestamo.Refresh
DtaMovPrestamo.Refresh
MsgBox "El Prestamo ha sido borrado"
DbgrLibreta.Columns(0).Visible = False
DbgrLibreta.Columns(1).Visible = False
DbgrLibreta.Columns(7).Visible = False

Exit Sub
TipoErrs:
 ControlErrores

End Sub

Private Sub CmdCarnet_Click()
'CodEmpleado = DBCodigoEmpleado.Text
'FrmCarnet.LblCedula.Caption = Me.TxtNumCedula.Text
'FrmCarnet.Show 1

' ArepDetalleDeduccion.DataControl1.ConnectionString = ConexionReporte
' ArepDetalleDeduccion.LblTitulo.Caption = Titulo
' ArepDetalleDeduccion.LblSubtitulo.Caption = "REPORTE DETALLADO DEDUCCIONES SEGUN NOMINA"
' ArepDetalleDeduccion.LblDesde.Caption = "Impreso desde: " & Me.TxtFecha1.Value & " Hasta: " & Me.TxtFecha2.Value
' ArepCarnet.LblSubtitulo.Caption = SubTitulo
 ArepCarnet.ImgLogo.Picture = LoadPicture(RutaLogo)
 ArepCarnet.lbltitulo.Caption = Titulo
' ArepCarnet.LblArea.Caption = Me.DBCDepartamento.Text
' ArepCarnet.lblCodigo.Caption = Me.DBCodigoEmpleado.Text
' ArepCarnet.LblApellidos.Caption = Me.TxtApellido1.Text & " " & Me.TxtApellido2.Text
' ArepCarnet.LblNombres.Caption = Me.TxtNombre1.Text & " " & Me.TxtNombre2.Text
' ArepCarnet.ImgFoto.Picture = Me.Image1.Picture
' ArepCarnet.LblCodigo2.Caption = Me.DBCodigoEmpleado.Text
ArepCarnet.DataControl1.ConnectionString = Conexion
SqlString = "SELECT  Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Cargo.Cargo, departamento.departamento , Empleado.Activo, Departamento.CodDepartamento FROM  Empleado INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  Where (Empleado.Activo = 1) AND (Empleado.CodEmpleado = " & Me.TxtCodEmpleado.Text & ") "
ArepCarnet.DataControl1.Source = SqlString
ArepCarnet.Show
End Sub

Private Sub CmdCuentas_Click()
 FrmCuentasContables.TxtCodEmpleado.Text = Me.TxtCodEmpleado.Text
 FrmCuentasContables.lbltitulo.Caption = "Empleado: " & Me.DBCodigoEmpleado.Text & " " & Me.TxtNombre1.Text & " " & Me.TxtNombre2.Text & Me.TxtApellido1.Text
 FrmCuentasContables.lbltitulo.Refresh
 FrmCuentasContables.Show 1
End Sub

Private Sub CmdDisponibles_Click()
Dim Numero As Integer, CodigoNuevo As String
Dim Respuesta As Double, Mes As Integer, CodEmpleado As Double
'//////////////////LIMPIO PANTALLA///////////////////////////
 Me.DBCodigoEmpleado.Text = ""
 LimpiaEmpleado
 LimpiaHistorico
 LimpiaInfNomina
 CodEmpleado = -1
 
 Respuesta = MsgBox("Desea reutilizar Codigos?", vbYesNo, "Zeus Facturacion")
 
 
 If Respuesta = 6 Then
        Me.AdoNumerosDisponibles.RecordSource = "SELECT DISTINCT CodEmpleado1 From dbo.Empleado WHERE(CodEmpleado1 NOT IN(SELECT CodEmpleado1 From Empleado WHERE  Activo = 1))ORDER BY CodEmpleado1"
        Me.AdoNumerosDisponibles.Refresh
        If Not Me.AdoNumerosDisponibles.Recordset.EOF Then
          Do While Not Me.AdoNumerosDisponibles.Recordset.EOF
           If Not Me.AdoNumerosDisponibles.Recordset("CodEmpleado1") = "" Then
             Me.DBCodigoEmpleado.Text = Me.AdoNumerosDisponibles.Recordset("CodEmpleado1")
             Exit Do
           End If
           Me.AdoNumerosDisponibles.Recordset.MoveNext
          Loop
        Else
          Me.DtaConsulta.RecordSource = "SELECT DISTINCT CodEmpleado1 From Empleado ORDER BY CodEmpleado1"
          Me.DtaConsulta.Refresh
          If Not DtaConsulta.Recordset.EOF Then
            Me.DtaConsulta.Recordset.MoveLast
            Numero = Me.DtaConsulta.Recordset("CodEmpleado1") + 1
            CodigoNuevo = Format(Numero, "00000#")
            Me.DBCodigoEmpleado.Text = CodigoNuevo
          Else
            CodigoNuevo = "000001"
            Me.DBCodigoEmpleado.Text = CodigoNuevo
          End If
        End If
        
  Else
    Mes = Month(Now)
    CodigoNuevo = "S1" & Mid(Year(Now), Len(Year(Now)) - 1, Len(Year(Now)) - 2) & Format(Mes, "0#")
    Me.DtaConsulta.RecordSource = "SELECT CodEmpleado1 From Empleado WHERE (CodEmpleado1 LIKE '%" & CodigoNuevo & "%') ORDER BY CodEmpleado1 DESC"
    Me.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Numero = Mid(Me.DtaConsulta.Recordset("CodEmpleado1"), 7, 4)
      CodigoNuevo = CodigoNuevo & Format(Numero + 1, "000#")
    Else
      CodigoNuevo = CodigoNuevo & "0001"
    End If
      Me.DBCodigoEmpleado.Text = CodigoNuevo
  
  End If
  
  '///////////////////////////////////////////////////////////////////////////////////////////////////////////
  '///////////////////////////////////AGREGO AL EMPLEADO CON EL CODIGO NUEVO /////////////////////////////////
  '///////////////////////////////////////////////////////////////////////////////////////////////////////////
    Me.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo, FechaAntiguedad, antiguedad,Dolarizado,CuentaBanco,SueldoActualBasico,HorasTurno,SalPorcentaje,CantPts,DiasBasico, AumentoBasico, ViaticoxDia, DeducirPorciento  From Empleado WHERE (CodEmpleado1 = '" & CodigoNuevo & "') And (Activo = 1)"
    Me.DtaEmpleado.Refresh
    If Me.DtaEmpleado.Recordset.EOF Then
    
          DtaEmpleado.Recordset.AddNew
            DtaEmpleado.Recordset("CodEmpleado1") = DBCodigoEmpleado.Text
          DtaEmpleado.Recordset.Update
          Me.DtaEmpleado.Refresh
          
          
          CodEmpleado1 = DtaEmpleado.Recordset("CodEmpleado1")
          CodEmpleado = DtaEmpleado.Recordset("CodEmpleado")
    
       Me.DtaHorarioEmpleado.RecordSource = "SELECT CodEmpleado, LEntrada, LSalida, MEntrada, MSalida, MCEntrada, MCSalida, JEntrada, JSalida, VEntrada, VSalida, TComida, TurnoLunes,TurnoMartes , TurnoMiercoles, TurnoJueves, TurnoViernes, TurnoSabado, TurnoDomingo, SEntrada, SSalida, DEntrada, DSalida From dbo.HorarioEmpleado WHERE(CodEmpleado = '" & CodEmpleado1 & "')"
       Me.DtaHorarioEmpleado.Refresh
       If Me.DtaHorarioEmpleado.Recordset.EOF Then
         Me.DtaTurnos.Refresh
         If Not Me.DtaTurnos.Recordset.EOF Then
           CodTurno = Me.DtaTurnos.Recordset("CodTurno")
           Me.DtaHorarioEmpleado.Recordset.AddNew
           Me.DtaHorarioEmpleado.Recordset("CodEmpleado") = CodEmpleado1
           Me.DtaHorarioEmpleado.Recordset("LEntrada") = Me.DtaTurnos.Recordset("LEntrada")
           Me.DtaHorarioEmpleado.Recordset("LSalida") = Me.DtaTurnos.Recordset("LSalida")
           Me.DtaHorarioEmpleado.Recordset("MEntrada") = Me.DtaTurnos.Recordset("MEntrada")
           Me.DtaHorarioEmpleado.Recordset("MSalida") = Me.DtaTurnos.Recordset("MSalida")
           Me.DtaHorarioEmpleado.Recordset("MCEntrada") = Me.DtaTurnos.Recordset("MCEntrada")
           Me.DtaHorarioEmpleado.Recordset("MCSalida") = Me.DtaTurnos.Recordset("MCSalida")
           Me.DtaHorarioEmpleado.Recordset("JEntrada") = Me.DtaTurnos.Recordset("JEntrada")
           Me.DtaHorarioEmpleado.Recordset("JSalida") = Me.DtaTurnos.Recordset("JSalida")
           Me.DtaHorarioEmpleado.Recordset("VEntrada") = Me.DtaTurnos.Recordset("VEntrada")
           Me.DtaHorarioEmpleado.Recordset("VSalida") = Me.DtaTurnos.Recordset("VSalida")
           Me.DtaHorarioEmpleado.Recordset("TComida") = Me.DtaTurnos.Recordset("TComida")
           Me.DtaHorarioEmpleado.Recordset("TurnoLunes") = CodTurno
           Me.DtaHorarioEmpleado.Recordset("TurnoMartes") = CodTurno
           Me.DtaHorarioEmpleado.Recordset("TurnoMiercoles") = CodTurno
           Me.DtaHorarioEmpleado.Recordset("TurnoJueves") = CodTurno
           Me.DtaHorarioEmpleado.Recordset("TurnoViernes") = CodTurno
           Me.DtaHorarioEmpleado.Recordset("TurnoSabado") = CodTurno
           Me.DtaHorarioEmpleado.Recordset("TurnoDomingo") = CodTurno
           Me.DtaHorarioEmpleado.Recordset("SEntrada") = Me.DtaTurnos.Recordset("SEntrada")
           Me.DtaHorarioEmpleado.Recordset("SSalida") = Me.DtaTurnos.Recordset("SEntrada")
           Me.DtaHorarioEmpleado.Recordset("DEntrada") = Me.DtaTurnos.Recordset("SEntrada")
           Me.DtaHorarioEmpleado.Recordset("DSalida") = Me.DtaTurnos.Recordset("SEntrada")
    
         Me.DtaHorarioEmpleado.Recordset.Update
         End If
       End If
    
    End If
  
  
  
  
  
   
'   Me.CmdAcercade.Caption = CodigoNuevo
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    
End Sub

Private Sub CmdEditarPresLinea_Click()
On Error GoTo TipoErr
FrmEditarPrestamo.TxtNumPrestamo.Text = DtaMovPrestamo.Recordset("NumPrestamo")
FrmEditarPrestamo.txtNumCuota.Text = DtaMovPrestamo.Recordset("numcuota")
FrmEditarPrestamo.TxtMontoOld.Text = Format(DtaMovPrestamo.Recordset("CuotaIgual"), "###,##0.00")
FrmEditarPrestamo.lblNombre = TxtNombre1.Text & " " & TxtNombre2.Text & " " & TxtApellido1.Text & " " & TxtApellido2.Text
FrmEditarPrestamo.Show 1

Exit Sub
TipoErr:
    ControlErrores
End Sub

Private Sub CmdEliminarDeduccion_Click()
On Error GoTo TipoErrs
Dim Numdeduccion As Double

Numdeduccion = Me.DbgDeducciones.Columns(0).Text
Me.DtaConsulta.RecordSource = "SELECT NumDeduccion, CodEmpleado, CodTipoDeduccion, NumVeces, Pagado, NUmNomina From Deduccion WHERE (NumDeduccion = " & Numdeduccion & ")"
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
 Me.DtaConsulta.Recordset.Delete
End If
Me.DtaDetalleDeduccion.Refresh
CmdEliminarDeduccion.Enabled = False
DbgDeducciones.Columns(0).Visible = False
DbgDeducciones.Columns(2).Visible = False
'DbgDeducciones.Columns(5).Visible = False


Exit Sub
TipoErrs:
  ControlErrores
  Unload Me

End Sub

Private Sub CmdEliminarIncentivo_Click()
On Error GoTo TipoErrs
Dim NumIncentivo As Integer

NumIncentivo = Me.DbGIncentivos.Columns(0).Text
Me.DtaConsulta.RecordSource = "SELECT CodEmpleado, CodTipoIncentivo, NumVeces, Pagado, NumIncentivo From Incentivo Where (NumIncentivo = " & NumIncentivo & ")"
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
 Me.DtaConsulta.Recordset.Delete
End If
Me.DtaDetalleIncentivo.Refresh
CmdEliminarIncentivo.Enabled = False

Exit Sub
TipoErrs:
  ControlErrores
  Unload Me

End Sub

Private Sub CmdEliminarSubsidio_Click()

On Error GoTo TipoErrs
Dim NumSubsidio As Integer

NumSubsidio = Me.DbgrSubsidios.Columns(0).Text
Me.DtaConsulta.RecordSource = "SELECT NumSubsidio, CodEmpleado, CodTipoSubsidio, NumVeces, Pagado From Subsidio Where (Numsubsidio = " & NumSubsidio & ")"
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
 Me.DtaConsulta.Recordset.Delete
End If
Me.DtaDetalleSubsidio.Refresh
CmdEliminarSubsidio.Enabled = False
Exit Sub
TipoErrs:
  ControlErrores
  Unload Me

End Sub

Private Sub CmdEstadoPrestamo_Click()

On Error GoTo TipoErr

 Dim CodigoEmpleado As String

    Me.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo,Dolarizado,CuentaBanco,SueldoActualBasico,HorasTurno,SalPorcentaje,CantPts,DiasBasico,AumentoBasico From Empleado WHERE (CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') And (Activo = 1)"
    Me.DtaEmpleado.Refresh

  If Not Me.DtaEmpleado.Recordset.EOF Then
    CodigoEmpleado = Me.DtaEmpleado.Recordset("CodEmpleado")
  End If

ARPagoPrestamo.DataControl1.ConnectionString = Conexion
ARPagoPrestamo.DataControl1.Source = "SELECT MovPrestamo.NumPrestamo, Prestamo.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Prestamo.Monto AS Prestamo, Prestamo.Saldo, Prestamo.CuotasIguales, MovPrestamo.NumCuota, MovPrestamo.Monto, MovPrestamo.Interes, MovPrestamo.CuotaIgual, MovPrestamo.SaldoCuota, MovPrestamo.Cancelado, MovPrestamo.NumNomina, Prestamo.Cancelado AS Activo FROM Empleado INNER JOIN  Prestamo ON Empleado.CodEmpleado = Prestamo.CodEmpleado INNER JOIN  MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo  " & _
                                     "Where (Prestamo.CodEmpleado = " & CodigoEmpleado & ") And (Prestamo.Cancelado = 0) ORDER BY MovPrestamo.NumPrestamo"

ARPagoPrestamo.Show

Exit Sub
TipoErr:
ControlErrores

End Sub

Private Sub CmdExportar_Click()
On Error GoTo TipoErr

DtaPrestamo.Refresh
Do While Not DtaPrestamo.Recordset.EOF
    If DtaPrestamo.Recordset("CodEmpleado") = DBCodigoEmpleado.Text Then
        Exit Do
    End If
DtaPrestamo.Recordset.MoveNext
Loop
    
Ruta2 = TxtRuta.Text

Open Ruta2 For Output As #1
    TextoMonto = Format(DtaPrestamo.Recordset("Monto"), "####0.00")
    For i = 1 To 15 - Len(TextoMonto)
       TextoMonto = " " + TextoMonto
    Next i
    
    Cadena = Trim(Str(Month(DtaPrestamo.Recordset("fechainicial"))))
    Cadena = Cadena + Trim(Str(Day(DtaPrestamo.Recordset("fechainicial"))))
    Cadena = Cadena + Trim(Str(Year(DtaPrestamo.Recordset("fechainicial"))))
    Cadena = Cadena + "        "
    Cadena = Cadena + "ZEUS"
    Cadena = Cadena + Trim(Str(TxtDebitoPrestamo.Text))
    For i = 1 To 36 - Len(Cadena)
    Cadena = Cadena + " "
    Next i
    Cadena = Cadena + "  "
    Cadena = Cadena + "                "
    Cadena = Cadena + "               "
    Cadena = Cadena + "          "
    Cadena = Cadena + "03"
    Cadena = Cadena + "      "
    Cadena = Cadena + "Pago de Prestamo " + Trim(Str(DtaPrestamo.Recordset("NumPrestamo"))) + "               "
    Cadena = Cadena + "                 "
    Cadena = Cadena + "   " + TextoMonto
    For i = 1 To 34
        Cadena = Cadena + " "
    Next i
    Cadena = Cadena + "00"
    Print #1, Cadena
    
    
    For i = 1 To 15 - Len(TextoMonto)
       TextoMonto = " " + TextoMonto
    Next i
    
    Cadena = Trim(Str(Month(DtaPrestamo.Recordset("fechainicial"))))
    Cadena = Cadena + Trim(Str(Day(DtaPrestamo.Recordset("fechainicial"))))
    Cadena = Cadena + Trim(Str(Year(DtaPrestamo.Recordset("fechainicial"))))
    Cadena = Cadena + "        "
    Cadena = Cadena + "ZEUS"
    Cadena = Cadena + Trim(Str(TxtCreditoPrestamo.Text))
    For i = 1 To 36 - Len(Cadena)
    Cadena = Cadena + " "
    Next i
    Cadena = Cadena + "  "
    Cadena = Cadena + "                "
    Cadena = Cadena + "               "
    Cadena = Cadena + "          "
    Cadena = Cadena + "07"
    Cadena = Cadena + "      "
    Cadena = Cadena + "Pago de Prestamo " + Trim(Str(DtaPrestamo.Recordset("NumPrestamo"))) + "               "
    Cadena = Cadena + "                 "
    Cadena = Cadena + "   " + TextoMonto
    For i = 1 To 34
        Cadena = Cadena + " "
    Next i
    Cadena = Cadena + "00"
    Print #1, Cadena
    
   
 Close #1

Exit Sub
TipoErr:
 ControlErrores
End Sub

Private Sub CmdGrabar_Click()
Dim Historico As Boolean
Dim Id As Integer, CodTurno As String
Dim CodEmpleado1 As String, Dolarizado As Boolean
Dim CodEmpleado As Double

'On Error GoTo TipoErrs
CmdAnterior.Enabled = False
CmdSiguiente.Enabled = False
CmdPrimero.Enabled = False
CmdUltimo.Enabled = False
CmdBorrar.Enabled = False

Historico = False

    Me.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo,Dolarizado,CuentaBanco,SueldoActualBasico,HorasTurno,SalPorcentaje,CantPts,DiasBasico,AumentoBasico From Empleado WHERE (CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') And (Activo = 1)"
    Me.DtaEmpleado.Refresh
  If Not Me.DtaEmpleado.Recordset.EOF Then
     CodEmpleado = Me.DtaEmpleado.Recordset("CodEmpleado")
  End If

Salida = False
frmEmpleado.MousePointer = 11
  If frmEmpleado.DBCodigoEmpleado.Text = "" Then
    MsgBox "Se necesita Código Empleado", vbCritical, "Error:Sistema de Nominas"
    frmEmpleado.MousePointer = 0
    Exit Sub
  End If
  
 
  If TxtNumHijos.Text = "" Then
    MsgBox "El número de hijos no puede quedar vacío", vbCritical, "Error:Sistema de Nominas"
    TxtNumHijos.SetFocus
    frmEmpleado.MousePointer = 0
    Exit Sub
  End If
  
  
  If frmEmpleado.TxtNombre1.Text = "" Then
    MsgBox "Se necesita Nombre del Empleado", vbCritical, "Error:Sistema de Nominas"
    frmEmpleado.MousePointer = 0
    Exit Sub
  End If
  
  If Me.txtAumentoBasico.Text = "" Then
    Me.txtAumentoBasico.Text = "0"
  End If
  
  If Me.TxtReembolso.Text = "" Then
    Me.TxtReembolso.Text = "0"
  End If
  
  If Me.TxtDiasBasico.Text = "" Then
    Me.TxtDiasBasico.Text = "0"
  End If
  
  If frmEmpleado.TxtNumCedula.Text = "" Then
    MsgBox "Se necesita el Numero de cédula", vbCritical, "Error:Sistema de Nominas"
    TxtNumCedula.SetFocus
    frmEmpleado.MousePointer = 0
    Exit Sub
  End If
  
  If frmEmpleado.TxtApellido1.Text = "" Then
      MsgBox "Se necesita Apellido del Empleado", vbCritical, "Error:Sistema de Nominas"
      frmEmpleado.MousePointer = 0
      Exit Sub
  End If
  
  If frmEmpleado.TxtDireccion.Text = "" Then
       MsgBox "Se necesita Dirección del Empleado", vbCritical, "Error:Sistema de Nominas"
       frmEmpleado.MousePointer = 0
       Exit Sub
  End If
  
  If frmEmpleado.TxtNacionalidad.Text = "" Then
      MsgBox "Se necesita Nacionalidad del Empleado", vbCritical, "Error:Sistema de Nominas"
      frmEmpleado.MousePointer = 0
      Exit Sub
  End If
  
  If frmEmpleado.CmbSexo.Text = "" Then
      MsgBox "Se necesita Sexo del Empleado", vbCritical, "Error:Sistema de Nominas"
      frmEmpleado.MousePointer = 0
      Exit Sub
  End If
  
  If frmEmpleado.DBCDepartamento.Text = "" Then
    MsgBox "Se necesita el Departamento del Empleado", vbCritical, "Error:Sistema de Nominas"
    frmEmpleado.MousePointer = 0
    Exit Sub
  End If
  
  If frmEmpleado.TxtCodCargo.Text = "" Then
   MsgBox "Se necesita el Cargo del Empleado", vbCritical, "Error:Sistema de Nominas"
   frmEmpleado.MousePointer = 0
   Exit Sub
  End If
  
 If frmEmpleado.MaskEdNacimiento.Value = "" Then
   MsgBox "Se necesita Fecha de Nacimiento", vbCritical, "Error:Sistema de Nominas"
   frmEmpleado.MousePointer = 0
   Exit Sub
  End If
  
' If frmEmpleado.TxtDebito.Text = "" Then
'   MsgBox "Se necesita Cuenta de Debito", vbCritical, "Error:Sistema de Nominas"
'   frmEmpleado.MousePointer = 0
'   Exit Sub
'  End If
'
' If frmEmpleado.TxtCredito.Text = "" Then
'   MsgBox "Se necesita Cuenta de Credito", vbCritical, "Error:Sistema de Nominas"
'   frmEmpleado.MousePointer = 0
'   Exit Sub
'  End If
  
If TxtCodGrupo.Text = "" Then
   MsgBox "Se necesita que selccione el grupo de la Nómina", vbCritical, "Error:Sistema de Nominas"
   frmEmpleado.MousePointer = 0
   Exit Sub
End If

  
 Select Case CmbTipoPago.Text
 
     Case ""
          MsgBox "Se necesita Tipo de Pago", vbCritical, "Error:Sistema de Nominas"
          frmEmpleado.MousePointer = 0
          Exit Sub
    Case "Salario Destajo"
    
        If TxtTarifaHoraria.Text = "" Then
           MsgBox "No ha gravado la tarifa Horaria"
           frmEmpleado.MousePointer = 0
           Exit Sub
        ElseIf Not IsNumeric(TxtTarifaHoraria.Text) Then
           MsgBox "La Tarifa Horaria es errónea"
           frmEmpleado.MousePointer = 0
           Exit Sub
 
        End If
           
                   
     Case "Salario Destajo y Comision"
     
        If TxtTarifaHoraria.Text = "" Or TxtComision.Text = "" Then
           MsgBox "La tarifa Horaria o la Comisión no ha sido Gravada"
           frmEmpleado.MousePointer = 0
           Exit Sub
        ElseIf Not IsNumeric(TxtTarifaHoraria.Text) Or Not IsNumeric(TxtComision.Text) Then
           MsgBox "La tarifa Horaria o la Comisión es errónea"
           frmEmpleado.MousePointer = 0
           Exit Sub
        End If
           
      Case "Sueldo Fijo"
        If TxtSueldoPeriodo.Text = "" Then
           MsgBox "El Sueldo del período no puede estar vacío"
           frmEmpleado.MousePointer = 0
           Exit Sub
        ElseIf Not IsNumeric(TxtSueldoPeriodo.Text) Then
           MsgBox "El Sueldo del período es erróneo"
           frmEmpleado.MousePointer = 0
           Exit Sub

        End If
                
       Case "Salario Fijo y Comision"
        If TxtSueldoPeriodo.Text = "" Or TxtComision.Text = "" Then
           MsgBox "El Sueldo del perídodo o la Comisión no puede estar vacío"
           frmEmpleado.MousePointer = 0
           Exit Sub
        ElseIf Not IsNumeric(TxtSueldoPeriodo.Text) Or Not IsNumeric(TxtComision.Text) Then
            MsgBox "El Sueldo del perídodo o la Comisión es erróneo"
            frmEmpleado.MousePointer = 0
            Exit Sub
        End If
        
 End Select
 
  If frmEmpleado.CmbExentoInss.Text = "" Then
   MsgBox "Se necesita Exento Inss", vbCritical, "Error:Sistema de Nominas"
   frmEmpleado.MousePointer = 0
   Exit Sub
  End If
  
 If frmEmpleado.CmbExentoIr.Text = "" Then
   MsgBox "Se necesita Exento Ir", vbCritical, "Error:Sistema de Nominas"
   frmEmpleado.MousePointer = 0
   Exit Sub
  End If
  
 If frmEmpleado.CmbPagoInssPatronal.Text = "" Then
   MsgBox "Se necesita Pago Inss Patronal", vbCritical, "Error:Sistema de Nominas"
   frmEmpleado.MousePointer = 0
   Exit Sub
  End If
 
 If frmEmpleado.TxtCodTipoNomina.Text = "" Then
   MsgBox "Se necesita Tipo de Nómina", vbCritical, "Error:Sistema de Nominas"
   frmEmpleado.MousePointer = 0
   Exit Sub
  End If
  
        TxtSueldoInicial.Text = Format((TxtSueldoInicial.Text), "##,##0.0000")
        TxtSueldoAnterior.Text = Format((TxtSueldoAnterior.Text), "##,##0.0000")
        TxtSueldoActual.Text = Format((TxtSueldoActual.Text), "##,##0.0000")
        TxtSueldoPeriodo.Text = Format((TxtSueldoPeriodo.Text), "##,##0.0000")
        TxtTarifaHoraria.Text = Format((TxtTarifaHoraria.Text), "##,##0.000000000")
  
  'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
     Valida = 0
     
  If Me.TxtSalarioPorciento.Text = "" Then
     Me.TxtSalarioPorciento.Text = 0
  End If
  

Me.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo, FechaAntiguedad, antiguedad,Dolarizado,CuentaBanco,SueldoActualBasico,HorasTurno,SalPorcentaje,CantPts,DiasBasico, AumentoBasico, ViaticoxDia, DeducirPorciento, Reembolso, Telefono  From Empleado WHERE (CodEmpleado = " & CodEmpleado & ") And (Activo = 1)"
Me.DtaEmpleado.Refresh
If Not Me.DtaEmpleado.Recordset.EOF Then



'      Do While Not DtaEmpleado.Recordset.EOF
'        If DtaEmpleado.Recordset("") = DBCodigoEmpleado.Text Then
            'DtaEmpleado.Recordset.Edit
            DtaEmpleado.Recordset("Nombre1") = TxtNombre1.Text
            DtaEmpleado.Recordset("Nombre2") = TxtNombre2.Text
            DtaEmpleado.Recordset("Apellido1") = TxtApellido1.Text
            DtaEmpleado.Recordset("Apellido2") = TxtApellido2.Text
            DtaEmpleado.Recordset("Direccion") = TxtDireccion.Text
            DtaEmpleado.Recordset("Nacionalidad") = TxtNacionalidad.Text
            DtaEmpleado.Recordset("CodigoPostal") = TxtCodPostal.Text
            DtaEmpleado.Recordset("sexo") = CmbSexo.Text
            DtaEmpleado.Recordset("NumeroInss") = TxtNInss.Text
            DtaEmpleado.Recordset("numcedula") = TxtNumCedula.Text
            DtaEmpleado.Recordset("numeroruc") = TxtNRuc.Text
            DtaEmpleado.Recordset("CodDepartamento") = TxtCodDepartamento.Text
            DtaEmpleado.Recordset("CodCargo") = TxtCodCargo.Text
            DtaEmpleado.Recordset("Codgrupo") = TxtCodGrupo.Text
            DtaEmpleado.Recordset("Sindicalista") = CmbSindicalista.Text
            DtaEmpleado.Recordset("CodTipoNomina") = CodTipoNomina
            DtaEmpleado.Recordset("numhijos") = TxtNumHijos.Text
            DtaEmpleado.Recordset("SalPorcentaje") = Me.TxtSalarioPorciento.Text
            DtaEmpleado.Recordset("AumentoBasico") = Me.txtAumentoBasico.Text
            DtaEmpleado.Recordset("Reembolso") = Me.TxtReembolso.Text
            DtaEmpleado.Recordset("Telefono") = Me.TxtTelefono.Text
            DtaEmpleado.Recordset("ViaticoxDia") = Me.TxtViatico.Text
            
             If Me.ChkSueldoActual.Value = xtpChecked Then
              DtaEmpleado.Recordset("SueldoActualBasico") = True
             Else
              DtaEmpleado.Recordset("SueldoActualBasico") = False
             End If
            
            
            If Me.TxtDiasAdicionales.Text = 0 Then
             DtaEmpleado.Recordset("CantPts") = 0
            Else
             DtaEmpleado.Recordset("CantPts") = Me.TxtDiasAdicionales.Text
            End If
            
            If Me.ChkDeducirPorcentaje.Value = 0 Then
              DtaEmpleado.Recordset("DeducirPorciento") = False
            Else
              DtaEmpleado.Recordset("DeducirPorciento") = True
            End If
            
            
            If Me.ChkDolarizado.Value = 0 Then
              DtaEmpleado.Recordset("Dolarizado") = False
            Else
              DtaEmpleado.Recordset("Dolarizado") = True
            End If
            
            If Me.ChkHorasTurno.Value = False Then
              DtaEmpleado.Recordset("HorasTurno") = False
            Else
              DtaEmpleado.Recordset("HorasTurno") = True
            End If
            
'            If Me.ChkHorasTurno.Value = False Then
'              DtaEmpleado.Recordset("SueldoActualBasico") = False
'            Else
'              DtaEmpleado.Recordset("SueldoActualBasico") = True
'            End If
            
            If Me.TxtCuentaBanco.Text <> "" Then
             DtaEmpleado.Recordset("CuentaBanco") = Me.TxtCuentaBanco.Text
            End If
            
            
            If Me.Check1.Value = 0 Then
             DtaEmpleado.Recordset("PorcientoIncentivo") = 0
            Else
             If Not Me.TxtPorcientoHora.Text = "" Then
               DtaEmpleado.Recordset("PorcientoIncentivo") = Me.TxtPorcientoHora.Text
             End If
            End If
            
           
            
            If Me.ChkSalarioFijo.Value = 1 Then
             DtaEmpleado.Recordset("SalarioFijo") = "S"
            Else
             DtaEmpleado.Recordset("SalarioFijo") = "N"
            End If
            
            'verifico si ya se le quito o puso la suspencion
            If ChkSuspendido.Value = 1 Then
              DtaEmpleado.Recordset("ausente") = True
            Else
              DtaEmpleado.Recordset("ausente") = False
                DtaSuspenciones.Refresh
                Do While Not DtaSuspenciones.Recordset.EOF
                   If DtaSuspenciones.Recordset("CodEmpleado") = CodEmpleado Then
                      'DtaSuspenciones.Recordset.Edit
                      DtaSuspenciones.Recordset("activo") = False
                      DtaSuspenciones.Recordset.Update
                      Exit Do
                   End If
                DtaSuspenciones.Recordset.MoveNext
                Loop
            
            End If
            LblSuspendido.Visible = False
            
                       
            'gravar los nuevos datos de la nómina
            
              If Not IsNull(TxtDiasDescuento.Text) Then
                 DtaEmpleado.Recordset("DiasDescuento") = TxtDiasDescuento.Text
              Else
                 DtaEmpleado.Recordset("DiasDescuento") = 0
              End If
              
              If Not IsNull(Me.TxtDiasBasico.Text) Then
                 DtaEmpleado.Recordset("DiasBasico") = Me.TxtDiasBasico.Text
              Else
                 DtaEmpleado.Recordset("DiasBasico") = 0
              End If
            
              DtaEmpleado.Recordset("SueldoPeriodo") = CDbl(TxtSueldoPeriodo.Text)
              DtaEmpleado.Recordset("TarifaHoraria") = CDbl(TxtTarifaHoraria.Text)
              DtaEmpleado.Recordset("PorcentajeComision") = CDbl(TxtComision.Text)
            
            If Not TxtOtrosIngresos.Text = "" Then
              DtaEmpleado.Recordset("OtrosIngresos") = CDbl(TxtOtrosIngresos.Text)
            End If
           If Not TxtDescripOtrIngre.Text = "" Then
              DtaEmpleado.Recordset("DescripOtrIngre") = TxtDescripOtrIngre.Text
           Else
              DtaEmpleado.Recordset("DescripOtrIngre") = "Sin Descrip"
            End If
            
            If CmbSalarioMinimo.Text = "Verdadero" Then
               DtaEmpleado.Recordset("salariominimo") = True
            Else
              DtaEmpleado.Recordset("salariominimo") = False
            End If
            
            If CmbExentoInss.Text = "Verdadero" Then
               DtaEmpleado.Recordset("ExentoInss") = True
            Else
              DtaEmpleado.Recordset("ExentoInss") = False
            End If
            
            If CmbExentoIr.Text = "Verdadero" Then
               DtaEmpleado.Recordset("ExentoIr") = True
            Else
              DtaEmpleado.Recordset("ExentoIr") = False
            End If

            If CmbPagoInssPatronal.Text = "Verdadero" Then
               DtaEmpleado.Recordset("PagoInssPatronal") = True
            Else
              DtaEmpleado.Recordset("PagoInssPatronal") = False
            End If
            
            DtaEmpleado.Recordset.Update
            
'/////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////AGREGA EMPLEADOS EN LA TABLA USUARIOS ////////////
'////////////////////////////////////////////////////////////////////////////////////
            Dim NumeroUser As Double
            If IsNumeric(DBCodigoEmpleado.Text) Then
                NumeroUser = DBCodigoEmpleado.Text
            Else
                Me.AdoUser.RecordSource = "SELECT  * From Userinfo ORDER BY Userid"
                Me.AdoUser.Refresh
                If Not Me.AdoUser.Recordset.EOF Then
                   Me.AdoUser.Recordset.MoveLast
                   NumeroUser = Me.AdoUser.Recordset("UserId") + 1
                Else
                    NumeroUser = 1
                End If
               
            End If
                
                Me.AdoUser.RecordSource = "SELECT  * From Userinfo WHERE (Userid = " & NumeroUser & ")"
                Me.AdoUser.Refresh
                If Me.AdoUser.Recordset.EOF Then
                   NumeroUser = ConsecutivoUser(DBCodigoEmpleado.Text)
                   Me.AdoUser.Recordset.AddNew
                     Me.AdoUser.Recordset("Userid") = NumeroUser
                     Me.AdoUser.Recordset("UserCode") = DBCodigoEmpleado.Text
                     Me.AdoUser.Recordset("Name") = Nombre1 + " " + Nombre2 + " " + Apellido1 + " " + Apellido2
                    Me.AdoUser.Recordset.Update
                Else
                    Me.AdoUser.Recordset("Name") = Nombre1 + " " + Nombre2 + " " + Apellido1 + " " + Apellido2
                    Me.AdoUser.Recordset.Update
                End If

            
            
            
          '  UbicaEmpleado
            Valida = 1
'            Exit Sub
            End If
            
            res = Bitacora(Now, NombreUsuario, "Empleados", "Editando al empleado: " & Me.TxtNombre1.Text)
'       DtaEmpleado.Recordset.MoveNext
'      Loop
' End If
      
   If Valida = 0 Then
   
   res = Bitacora(Now, NombreUsuario, "Empleados", "agrego al empleado: " & DBCodigoEmpleado.Text & " " & Me.TxtNombre1.Text)
   
            DtaEmpleado.Recordset.AddNew
            DtaEmpleado.Recordset("CodEmpleado1") = DBCodigoEmpleado.Text
            DtaEmpleado.Recordset("Nombre1") = TxtNombre1.Text
            DtaEmpleado.Recordset("Nombre2") = TxtNombre2.Text
            DtaEmpleado.Recordset("Apellido1") = TxtApellido1.Text
            DtaEmpleado.Recordset("Apellido2") = TxtApellido2.Text
            DtaEmpleado.Recordset("Direccion") = TxtDireccion.Text
            DtaEmpleado.Recordset("Nacionalidad") = TxtNacionalidad.Text
            DtaEmpleado.Recordset("CodigoPostal") = TxtCodPostal.Text
            DtaEmpleado.Recordset("numcedula") = TxtNumCedula.Text
            DtaEmpleado.Recordset("sexo") = CmbSexo.Text
            DtaEmpleado.Recordset("NumeroInss") = TxtNInss.Text
            DtaEmpleado.Recordset("numeroruc") = TxtNRuc.Text
            DtaEmpleado.Recordset("CodDepartamento") = TxtCodDepartamento.Text
            DtaEmpleado.Recordset("CodCargo") = TxtCodCargo.Text
            DtaEmpleado.Recordset("Codgrupo") = TxtCodGrupo.Text
            DtaEmpleado.Recordset("Sindicalista") = CmbSindicalista.Text
            DtaEmpleado.Recordset("CodTipoNomina") = CodTipoNomina
            DtaEmpleado.Recordset("numhijos") = TxtNumHijos.Text
            DtaEmpleado.Recordset("Telefono") = Me.TxtTelefono.Text
            If txtAumentoBasico.Text = "" Then
            txtAumentoBasico.Text = 0
            End If
            
            
            DtaEmpleado.Recordset("AumentoBasico") = CDbl(Me.txtAumentoBasico.Text)
            
            
            
            'grabar los nuevos datos de la nómina
              DtaEmpleado.Recordset("SueldoPeriodo") = CDbl(TxtSueldoPeriodo.Text)
              DtaEmpleado.Recordset("TarifaHoraria") = CDbl(TxtTarifaHoraria.Text)
              DtaEmpleado.Recordset("PorcentajeComision") = CDbl(TxtComision.Text)
              
            If Me.ChkDolarizado.Value = 0 Then
              DtaEmpleado.Recordset("Dolarizado") = False
            Else
              DtaEmpleado.Recordset("Dolarizado") = True
            End If
            
            If Me.TxtCuentaBanco.Text <> "" Then
             DtaEmpleado.Recordset("CuentaBanco") = Me.TxtCuentaBanco.Text
            End If
            
            
            If Me.TxtDiasAdicionales.Text = 0 Then
             DtaEmpleado.Recordset("CantPts") = 0
            Else
             DtaEmpleado.Recordset("CantPts") = Me.TxtDiasAdicionales.Text
            End If
            
            
        If Me.Check1.Value = 0 Then
         DtaEmpleado.Recordset("PorcientoIncentivo") = 0
        Else
         If Not Me.TxtPorcientoHora.Text = "" Then
           DtaEmpleado.Recordset("PorcientoIncentivo") = Me.TxtPorcientoHora.Text
         End If
 
        End If
        
            If Me.ChkHorasTurno.Value = False Then
              DtaEmpleado.Recordset("HorasTurno") = False
            Else
              DtaEmpleado.Recordset("HorasTurno") = True
            End If
            
            
            If Not TxtOtrosIngresos.Text = "" Then
               DtaEmpleado.Recordset("OtrosIngresos") = CDbl(TxtOtrosIngresos.Text)
            End If
            If Not TxtDescripOtrIngre.Text = "" Then
               DtaEmpleado.Recordset("DescripOtrIngre") = TxtDescripOtrIngre.Text
            Else
               DtaEmpleado.Recordset("DescripOtrIngre") = "Sin Descrip"
            End If
            
            If CmbSalarioMinimo.Text = "Verdadero" Then
               DtaEmpleado.Recordset("salariominimo") = True
            Else
              DtaEmpleado.Recordset("salariominimo") = False
            End If
            
            If CmbExentoInss.Text = "Verdadero" Then
               DtaEmpleado.Recordset("ExentoInss") = True
            Else
              DtaEmpleado.Recordset("ExentoInss") = False
            End If
            
            If CmbExentoIr.Text = "Verdadero" Then
               DtaEmpleado.Recordset("ExentoIr") = True
            Else
              DtaEmpleado.Recordset("ExentoIr") = False
            End If

            If CmbPagoInssPatronal.Text = "Verdadero" Then
               DtaEmpleado.Recordset("PagoInssPatronal") = True
            Else
              DtaEmpleado.Recordset("PagoInssPatronal") = False
            End If
            
             If Not IsNull(Me.TxtDiasBasico.Text) Then
                 DtaEmpleado.Recordset("DiasBasico") = Me.TxtDiasBasico.Text
             Else
                 DtaEmpleado.Recordset("DiasBasico") = 0
             End If
             
             If Me.ChkSueldoActual.Value = xtpChecked Then
              DtaEmpleado.Recordset("SueldoActualBasico") = 1
             Else
              DtaEmpleado.Recordset("SueldoActualBasico") = 0
             End If
            
      DtaEmpleado.Recordset.Update
            
            
    CodEmpleado1 = DtaEmpleado.Recordset("CodEmpleado1")
    CodEmpleado = DtaEmpleado.Recordset("CodEmpleado")
    
    Me.DtaHorarioEmpleado.RecordSource = "SELECT CodEmpleado, LEntrada, LSalida, MEntrada, MSalida, MCEntrada, MCSalida, JEntrada, JSalida, VEntrada, VSalida, TComida, TurnoLunes,TurnoMartes , TurnoMiercoles, TurnoJueves, TurnoViernes, TurnoSabado, TurnoDomingo, SEntrada, SSalida, DEntrada, DSalida From dbo.HorarioEmpleado WHERE(CodEmpleado ='" & CodEmpleado1 & "')"
    Me.DtaHorarioEmpleado.Refresh
    If Me.DtaHorarioEmpleado.Recordset.EOF Then
      Me.DtaTurnos.Refresh
      If Not Me.DtaTurnos.Recordset.EOF Then
        CodTurno = Me.DtaTurnos.Recordset("CodTurno")
        Me.DtaHorarioEmpleado.Recordset.AddNew
        Me.DtaHorarioEmpleado.Recordset("CodEmpleado") = CodEmpleado1
        Me.DtaHorarioEmpleado.Recordset("LEntrada") = Me.DtaTurnos.Recordset("LEntrada")
        Me.DtaHorarioEmpleado.Recordset("LSalida") = Me.DtaTurnos.Recordset("LSalida")
        Me.DtaHorarioEmpleado.Recordset("MEntrada") = Me.DtaTurnos.Recordset("MEntrada")
        Me.DtaHorarioEmpleado.Recordset("MSalida") = Me.DtaTurnos.Recordset("MSalida")
        Me.DtaHorarioEmpleado.Recordset("MCEntrada") = Me.DtaTurnos.Recordset("MCEntrada")
        Me.DtaHorarioEmpleado.Recordset("MCSalida") = Me.DtaTurnos.Recordset("MCSalida")
        Me.DtaHorarioEmpleado.Recordset("JEntrada") = Me.DtaTurnos.Recordset("JEntrada")
        Me.DtaHorarioEmpleado.Recordset("JSalida") = Me.DtaTurnos.Recordset("JSalida")
        Me.DtaHorarioEmpleado.Recordset("VEntrada") = Me.DtaTurnos.Recordset("VEntrada")
        Me.DtaHorarioEmpleado.Recordset("VSalida") = Me.DtaTurnos.Recordset("VSalida")
        Me.DtaHorarioEmpleado.Recordset("TComida") = Me.DtaTurnos.Recordset("TComida")
        Me.DtaHorarioEmpleado.Recordset("TurnoLunes") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("TurnoMartes") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("TurnoMiercoles") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("TurnoJueves") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("TurnoViernes") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("TurnoSabado") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("TurnoDomingo") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("SEntrada") = Me.DtaTurnos.Recordset("SEntrada")
        Me.DtaHorarioEmpleado.Recordset("SSalida") = Me.DtaTurnos.Recordset("SEntrada")
        Me.DtaHorarioEmpleado.Recordset("DEntrada") = Me.DtaTurnos.Recordset("SEntrada")
        Me.DtaHorarioEmpleado.Recordset("DSalida") = Me.DtaTurnos.Recordset("SEntrada")
 
      Me.DtaHorarioEmpleado.Recordset.Update
      End If
    End If
  

         
          frmEmpleado.MousePointer = 0
          UbicaEmpleado
Salida = False
 
  
  
  
  End If 'del valida = 0
'Gravo el histórico


       CodEmpleado1 = DtaEmpleado.Recordset("CodEmpleado1")
    
    
    Me.DtaHorarioEmpleado.RecordSource = "SELECT CodEmpleado, LEntrada, LSalida, MEntrada, MSalida, MCEntrada, MCSalida, JEntrada, JSalida, VEntrada, VSalida, TComida, TurnoLunes,TurnoMartes , TurnoMiercoles, TurnoJueves, TurnoViernes, TurnoSabado, TurnoDomingo, SEntrada, SSalida, DEntrada, DSalida From dbo.HorarioEmpleado WHERE(CodEmpleado = '" & CodEmpleado1 & "')"
    Me.DtaHorarioEmpleado.Refresh
    If Me.DtaHorarioEmpleado.Recordset.EOF Then
      Me.DtaTurnos.Refresh
      If Not Me.DtaTurnos.Recordset.EOF Then
        CodTurno = Me.DtaTurnos.Recordset("CodTurno")
      Me.DtaHorarioEmpleado.Recordset.AddNew
        Me.DtaHorarioEmpleado.Recordset("CodEmpleado") = CodEmpleado1
        Me.DtaHorarioEmpleado.Recordset("LEntrada") = Me.DtaTurnos.Recordset("LEntrada")
        Me.DtaHorarioEmpleado.Recordset("LSalida") = Me.DtaTurnos.Recordset("LSalida")
        Me.DtaHorarioEmpleado.Recordset("MEntrada") = Me.DtaTurnos.Recordset("MEntrada")
        Me.DtaHorarioEmpleado.Recordset("MSalida") = Me.DtaTurnos.Recordset("MSalida")
        Me.DtaHorarioEmpleado.Recordset("MCEntrada") = Me.DtaTurnos.Recordset("MCEntrada")
        Me.DtaHorarioEmpleado.Recordset("MCSalida") = Me.DtaTurnos.Recordset("MCSalida")
        Me.DtaHorarioEmpleado.Recordset("JEntrada") = Me.DtaTurnos.Recordset("JEntrada")
        Me.DtaHorarioEmpleado.Recordset("JSalida") = Me.DtaTurnos.Recordset("JSalida")
        Me.DtaHorarioEmpleado.Recordset("VEntrada") = Me.DtaTurnos.Recordset("VEntrada")
        Me.DtaHorarioEmpleado.Recordset("VSalida") = Me.DtaTurnos.Recordset("VSalida")
        Me.DtaHorarioEmpleado.Recordset("TComida") = Me.DtaTurnos.Recordset("TComida")
        Me.DtaHorarioEmpleado.Recordset("TurnoLunes") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("TurnoMartes") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("TurnoMiercoles") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("TurnoJueves") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("TurnoViernes") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("TurnoSabado") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("TurnoDomingo") = CodTurno
        Me.DtaHorarioEmpleado.Recordset("SEntrada") = Me.DtaTurnos.Recordset("SEntrada")
        Me.DtaHorarioEmpleado.Recordset("SSalida") = Me.DtaTurnos.Recordset("SEntrada")
        Me.DtaHorarioEmpleado.Recordset("DEntrada") = Me.DtaTurnos.Recordset("SEntrada")
        Me.DtaHorarioEmpleado.Recordset("DSalida") = Me.DtaTurnos.Recordset("SEntrada")
 
      Me.DtaHorarioEmpleado.Recordset.Update
      End If
    End If

      Me.DtaHistorico.RecordSource = "SELECT Id, Codempleado, FechaBaja, MotivoBaja, FechaAumento, MotivoAumento, FechaInicialSusp, FechaFinalSusp, MotivoSuspencion, FechaNacimiento,FechaContrato,FechaContratoVac , CargoInicial, CargoActual, CargoAnterior, SueldoInicial, SueldoAnterior, SueldoActual, CuentaDebito, CuentaCredito,CuentaPrestamo,CuentaOtrosIngresos,CuentaINSS,CuentaIR From Historico WHERE (Codempleado = " & CodEmpleado & ")"
      DtaHistorico.Refresh
      If Not Me.DtaHistorico.Recordset.EOF Then

      
       If DtaHistorico.Recordset("CodEmpleado") = CodEmpleado Then
            Historico = False
            'DtaHistorico.Recordset.Edit
                DtaHistorico.Recordset("CodEmpleado") = CodEmpleado
           
                DtaHistorico.Recordset.Fields("FechaNacimiento") = Format(MaskEdNacimiento.Value, "dd/MM/yyyy")
                DtaHistorico.Recordset.Fields("FechaContrato") = Format(MaskEdContrato.Value, "dd/MM/yyyy")
            
            
            DtaHistorico.Recordset.Fields("FechaContratoVac") = Me.DTPFechaContratoVaca.Value
            DtaHistorico.Recordset.Fields("CargoInicial") = DBCargoInicial.Text
            DtaHistorico.Recordset.Fields("CargoActual") = DBCargoActual.Text
            DtaHistorico.Recordset.Fields("CargoAnterior") = DBCargoAnterior.Text
            If TxtSueldoInicial.Text <> "" Then
            DtaHistorico.Recordset.Fields("SueldoInicial") = TxtSueldoInicial.Text
            End If
            If TxtSueldoAnterior.Text <> "" Then
             DtaHistorico.Recordset.Fields("SueldoAnterior") = TxtSueldoAnterior.Text
            End If
            If TxtSueldoActual.Text <> "" Then
             DtaHistorico.Recordset.Fields("SueldoActual") = TxtSueldoActual.Text
            End If
            If Not frmEmpleado.MaskEdBaja.Text = "__/__/____" Then
                DtaHistorico.Recordset.Fields("FechaBaja") = MaskEdBaja.Text
            End If
            DtaHistorico.Recordset.Fields("MotivoBaja") = TxtMotivoBaja.Text
                If Not frmEmpleado.MaskEdAumento.Text = "__/__/____" Then
            DtaHistorico.Recordset.Fields("FechaAumento") = MaskEdAumento.Text
            End If
            DtaHistorico.Recordset.Fields("MotivoAumento") = TxtMotivoAumento.Text
            If Not MaskEdSuspencion.Text = "__/__/____" Then
                DtaHistorico.Recordset.Fields("FechaInicialSusp") = MaskEdSuspencion.Text
            End If
            
            If Not MaskEdFinalSusp.Text = "__/__/____" Then
                DtaHistorico.Recordset.Fields("FechaFinalSusp") = MaskEdFinalSusp.Text
            End If
            DtaHistorico.Recordset.Fields("MotivoSuspencion") = TxtMotivoSuspencion.Text
'            DtaHistorico.Recordset.Fields("CuentaDebito") = TxtDebito.Text
'            DtaHistorico.Recordset.Fields("CuentaCredito") = TxtCredito.Text
            
'            If Me.TxtCtaPrestamo.Text <> "" Then
'              DtaHistorico.Recordset("CuentaPrestamo") = Me.TxtCtaPrestamo.Text
'            End If
'
'            If Me.TxtCtaOtrosIngresos.Text <> "" Then
'              DtaHistorico.Recordset("CuentaOtrosIngresos") = Me.TxtCtaOtrosIngresos.Text
'            End If
'
'            If Me.TxtCuentaInss.Text <> "" Then
'              DtaHistorico.Recordset("CuentaINSS") = Me.TxtCuentaInss.Text
'            End If
'
'            If Me.TxtCuentaIR.Text <> "" Then
'              DtaHistorico.Recordset("CuentaIR") = Me.TxtCuentaIR.Text
'            End If
            
            DtaHistorico.Recordset.Update
            
'            DtaEmpleado.Recordset!FechaAntiguedad = Format(MaskEdContrato.Value, "dd/MM/yyyy")
'            DtaEmpleado.Recordset!Antiguedad = 0
            DtaEmpleado.Recordset.Update
            
            Valida = 1
        End If
 
     Else
     
        Me.DtaConsulta.RecordSource = "SELECT Id From Historico"
        Me.DtaConsulta.Refresh
        If DtaConsulta.Recordset.EOF Then
          Id = 1
        Else
          Me.DtaConsulta.Recordset.MoveLast
          Id = Me.DtaConsulta.Recordset("id") + 1
        End If
      'si no existe un historico lo creo

            DtaHistorico.Recordset.AddNew
                      Me.DtaHistorico.Recordset("id") = Id
                      DtaHistorico.Recordset("CodEmpleado") = CodEmpleado
                  
                      DtaHistorico.Recordset.Fields("FechaNacimiento") = Format(MaskEdNacimiento.Value, "dd/MM/YYYY")

                 
                      DtaHistorico.Recordset.Fields("FechaContrato") = Format(MaskEdContrato.Value, "dd/MM/yyyy")
                      DtaHistorico.Recordset.Fields("FechaContratoVac") = Format(MaskEdContrato.Value, "dd/MM/yyyy")

                  DtaHistorico.Recordset.Fields("FechaContratoVac") = Me.DTPFechaContratoVaca.Value
                  DtaHistorico.Recordset.Fields("CargoInicial") = DBCargoInicial.Text
                  DtaHistorico.Recordset.Fields("CargoActual") = DBCargoActual.Text
                  DtaHistorico.Recordset.Fields("CargoAnterior") = DBCargoAnterior.Text
                  If TxtSueldoInicial.Text <> "" Then
                    DtaHistorico.Recordset.Fields("SueldoInicial") = TxtSueldoInicial.Text
                  End If
                  If TxtSueldoAnterior.Text <> "" Then
                    DtaHistorico.Recordset.Fields("SueldoAnterior") = TxtSueldoAnterior.Text
                  End If
                  If TxtSueldoActual.Text <> "" Then
                    DtaHistorico.Recordset.Fields("SueldoActual") = TxtSueldoActual.Text
                  End If
                  If Not frmEmpleado.MaskEdBaja.Text = "__/__/____" Then
                      DtaHistorico.Recordset.Fields("FechaBaja") = MaskEdBaja.Text
                  End If
                  If TxtMotivoBaja.Text <> "" Then
                    DtaHistorico.Recordset.Fields("MotivoBaja") = TxtMotivoBaja.Text
                  End If
                      If Not frmEmpleado.MaskEdAumento.Text = "__/__/____" Then
                  DtaHistorico.Recordset.Fields("FechaAumento") = MaskEdAumento.Text
                  End If
                  If TxtMotivoAumento.Text <> "" Then
                   DtaHistorico.Recordset.Fields("MotivoAumento") = TxtMotivoAumento.Text
                  End If
                  If Not MaskEdSuspencion.Text = "__/__/____" Then
                      DtaHistorico.Recordset.Fields("FechaInicialSusp") = MaskEdSuspencion.Text
                  End If
                  
                  If Not MaskEdFinalSusp.Text = "__/__/____" Then
                      DtaHistorico.Recordset.Fields("FechaFinalSusp") = MaskEdFinalSusp.Text
                  End If
                  If TxtMotivoSuspencion.Text <> "" Then
                   DtaHistorico.Recordset.Fields("MotivoSuspencion") = TxtMotivoSuspencion.Text
                  End If
'                  DtaHistorico.Recordset.Fields("CuentaDebito") = TxtDebito.Text
'                  DtaHistorico.Recordset.Fields("CuentaCredito") = TxtCredito.Text
                  
'                    If Me.TxtCtaPrestamo.Text <> "" Then
'                      DtaHistorico.Recordset("CuentaPrestamo") = Me.TxtCtaPrestamo.Text
'                    End If
'
'                    If Me.TxtCtaOtrosIngresos.Text <> "" Then
'                      DtaHistorico.Recordset("CuentaOtrosIngresos") = Me.TxtCtaOtrosIngresos.Text
'                    End If
'
'                    If Me.TxtCuentaInss.Text <> "" Then
'                      DtaHistorico.Recordset("CuentaINSS") = Me.TxtCuentaInss.Text
'                    End If
'
'                    If Me.TxtCuentaIR.Text <> "" Then
'                      DtaHistorico.Recordset("CuentaIR") = Me.TxtCuentaIR.Text
'                    End If
                  DtaHistorico.Recordset.Update
                  
'            DtaEmpleado.Recordset!FechaAntiguedad = Format(MaskEdContrato.Value, "dd/MM/yyyy")
'            DtaEmpleado.Recordset!Antiguedad = 0
'            DtaEmpleado.Recordset.Update
            
      End If
'

Me.DBCodigoEmpleado.Text = ""
If Not Me.DtaEmpleado.Recordset.EOF Then
  DtaEmpleado.Recordset.MoveLast
  DtaEmpleado.Recordset.MovePrevious
End If



frmEmpleado.MousePointer = 0

CmdAnterior.Enabled = True
CmdSiguiente.Enabled = True
CmdPrimero.Enabled = True
CmdUltimo.Enabled = True
CmdBorrar.Enabled = True
SSTab1.Tab = 0
LimpiaEmpleado
With Me.DtaEmpleados
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT CodEmpleado, CodEmpleado1, Activo, Nombre1 + ' '+ Nombre2 +' '+Apellido1+' '+Apellido2 as Nombres From Empleado Where (Activo = 1) ORDER BY CodEmpleado1"
   .Refresh
End With

'Me.DBCodigoEmpleado.Columns(0).Visible = False
'Me.DBCodigoEmpleado.Columns(1).Caption = "Codigo"
'Me.DBCodigoEmpleado.Columns(1).Width = 800
'Me.DBCodigoEmpleado.Columns(2).Visible = False

Exit Sub

TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdHistoDeducciones_Click()
CodEmpleado = DBCodigoEmpleado.Text
NumReport = 13
FrmVerReportes.Show
End Sub

Private Sub CmdHistoIncentivos_Click()
CodEmpleado = DBCodigoEmpleado.Text
NumReport = 12
FrmVerReportes.Show

End Sub

Private Sub CmdHistoSubsidios_Click()
CodEmpleado = DBCodigoEmpleado.Text
NumReport = 14
FrmVerReportes.Show
End Sub

Private Sub CmdIncapacidad_Click(Index As Integer)
 FrmIncapacidades.DbCCodEmpleado = DBCodigoEmpleado
 FrmIncapacidades.Show
End Sub

Private Sub CmdPrestamos_Click(Index As Integer)
 FrmPrestamos.DbCCodEmpleado.Text = DBCodigoEmpleado.Text
 FrmPrestamos.Show
End Sub

Private Sub CmdPrimero_Click()
On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Empleado")
If Contesta Then
  CmdGrabar.Value = True
 
End If
 DtaEmpleados.Recordset.MoveFirst
       If DtaEmpleados.Recordset.BOF Then
           DtaEmpleados.Recordset.MoveNext
           MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
       Else
           DBCodigoEmpleado.Text = DtaEmpleados.Recordset("CodEmpleado1")
           Llenado
       End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdRuta_Click()
QuienLlama = "Exporta Préstamo"
FrmRuta.Caption = "Buscando Ruta de Préstamo"
FrmRuta.Show 1
End Sub

Private Sub CmdSiguiente_Click()
'On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Empleado")
If Contesta Then
  CmdGrabar.Value = True
  
End If
 DtaEmpleados.Recordset.MoveNext
       If DtaEmpleados.Recordset.EOF Then
           DtaEmpleados.Recordset.MovePrevious
           MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
       Else
           DBCodigoEmpleado.Text = DtaEmpleados.Recordset("CodEmpleado1")
           Llenado
       End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdUltimo_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Empleado")
If Contesta Then
  CmdGrabar.Value = True
  
End If
 DtaEmpleados.Recordset.MoveLast
       If DtaEmpleados.Recordset.EOF Then
           DtaEmpleados.Recordset.MovePrevious
           MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
       Else
           DBCodigoEmpleado.Text = DtaEmpleados.Recordset("CodEmpleado1")
           Llenado
       End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub
Private Sub Salir_Click()
Unload Me
End Sub

Private Sub AgregaFoto_Click()
frmDrag.Show
End Sub

Private Sub CmdAnotaciones_Click()
If DBCodigoEmpleado.Text = "" Then Exit Sub
FrmAnotaciones.DBEmpleado.Text = DBCodigoEmpleado
FrmAnotaciones.Show
FrmAnotaciones.CmdBuscarEmpleado.Enabled = False
FrmAnotaciones.DBEmpleado.Enabled = False
End Sub

Private Sub CmdCerrar_Click()
Unload Me
End Sub


Private Sub Command2_Click()
'QueProducto = "CuentaContable"
'FrmConsulta.Show 1
'Me.TxtCredito.Text = CuentaContable
End Sub

Private Sub Command1_Click()
'QueProducto = "CuentaContable"
'FrmConsulta.Show 1
'Me.TxtDebito.Text = CuentaContable
End Sub

Private Sub Command3_Click()
'QueProducto = "CuentaContable"
'FrmConsulta.Show 1
'Me.TxtCtaPrestamo.Text = CuentaContable
End Sub

Private Sub Command4_Click()
'QueProducto = "CuentaContable"
'FrmConsulta.Show 1
'Me.TxtCtaOtrosIngresos.Text = CuentaContable
End Sub

Private Sub Command5_Click()
'QueProducto = "CuentaContable"
'FrmConsulta.Show 1
'Me.TxtCuentaInss.Text = CuentaContable
End Sub

Private Sub Command6_Click()
'QueProducto = "CuentaContable"
'FrmConsulta.Show 1
'Me.TxtCuentaIR.Text = CuentaContable
End Sub

Private Sub DBCargoActual_Click(Area As Integer)

Salida = True
End Sub

Private Sub DBCargoActual_Change()
PreparaSalida
End Sub

Private Sub DBCargoActual_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtSueldoInicial = Format((TxtSueldoInicial), "##,##0.00")
  TxtSueldoInicial.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub DBCActual_Click(Area As Integer)

End Sub

Private Sub DBCargoAnterior_Click(Area As Integer)

Salida = True
End Sub

Private Sub DBCargoAnterior_Change()
PreparaSalida
End Sub

Private Sub DBCargoAnterior_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  DBCargoActual.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub DBCargoInicial_Click(Area As Integer)

Salida = True
End Sub

Private Sub DBCargoInicial_Change()
PreparaSalida
End Sub

Private Sub DBCargoInicial_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  DBCargoAnterior.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub DBCCargo_Click(Area As Integer)
Salida = True
End Sub

Private Sub DBCCargo_Change()
 On Error GoTo TipoErrs
PreparaSalida
 'Al ejecutar algun cambio en el combo actualizo el nombre del departamento
   DtaCargo.Refresh
   Do While Not DtaCargo.Recordset.EOF
     If DtaCargo.Recordset("Cargo") = DBCCargo.Text Then
        TxtCodCargo.Text = DtaCargo.Recordset("CodCargo")
        
        Exit Do
     Else
        TxtCodCargo.Text = ""
     End If
       DtaCargo.Recordset.MoveNext
   Loop
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub


Private Sub DBCCargo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
CmbSindicalista.SetFocus
 Else
   Evaluar = False
  End If
End Sub



Private Sub DBCDepartamento_Click(Area As Integer)
Salida = True
'Evaluar = False
'PreparaSalida
End Sub

Private Sub DBCDepartamento_Change()
On Error GoTo TipoErrs
Salida = True
 ' Al ejecutar algun cambio en el combo actualizo el nombre del departamento
   DtaDepartamento.Refresh
   Do While Not DtaDepartamento.Recordset.EOF
     If DtaDepartamento.Recordset("departamento") = DBCDepartamento Then
        TxtCodDepartamento.Text = DtaDepartamento.Recordset("CodDepartamento")
        Exit Do
     Else
        TxtCodDepartamento.Text = ""
     End If
       DtaDepartamento.Recordset.MoveNext
   Loop
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub



Private Sub DBCDepartamento_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  DBCCargo.SetFocus
 Else
   Salida = True
  End If
End Sub



Private Sub DBCGrupo_Change()
DtaGrupo.Refresh
Do While Not DtaGrupo.Recordset.EOF
   If DBCGrupo.Text = DtaGrupo.Recordset("grupo") Then
      TxtCodGrupo.Text = DtaGrupo.Recordset("Codgrupo")
      Exit Sub
   End If
DtaGrupo.Recordset.MoveNext
Loop
End Sub



Private Sub DBCodigoEmpleado_Change()

 Me.TxtCodigoEmpleados.Text = Me.DBCodigoEmpleado.Text

'On Error GoTo TipoErrs
'Dim SQlIncentivos As String, SQlDeducciones As String, SqlDetallePrestamo As String, SQlPrestamo As String, SqlDetalleSubsidio As String
'Dim Salario As Boolean
'Dim CodEmpleado1 As String
'Dim numeroPrestamo As Double
'
'
'
'
'
'
'
'             Evaluar = True
'             'Al ejecutar algun cambio en el combo actualizo el nombre del Empleado
'             frmEmpleado.MousePointer = 11
'
'             LimpiaEmpleado
'             LimpiaHistorico
'             LimpiaInfNomina
'             'LimpiaInfNomina
'            ' DtaEmpleados.Refresh
'             ChkSuspendido.Visible = False
'            'Busco el codigo del empleado para que automaticamente ubique el nombre
'             'aunque no existe en la data consulta
'             CodEmpleado = -1
'
'            Me.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo,Dolarizado,CuentaBanco,SueldoActualBasico,HorasTurno From Empleado WHERE (CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') And (Activo = 1)"
'            Me.DtaEmpleado.Refresh
'
'If Not Me.DtaEmpleado.Recordset.EOF Then
'
'     CodEmpleado = Me.DtaEmpleado.Recordset("CodEmpleado")
'     TxtCodEmpleado.Text = Me.DtaEmpleado.Recordset("CodEmpleado")
'
'        If Not IsNull(DtaEmpleado.Recordset("numeroruc")) Then
'          TxtNRuc.Text = DtaEmpleado.Recordset("numeroruc")
'        End If
'        'busco el tipo del archivo
'        'Destino = ""
'        If Dir(RutaFoto & DBCodigoEmpleado.Text & ".jpg") <> "" Then
'           Destino = RutaFoto & DBCodigoEmpleado.Text & ".jpg"
'        ElseIf Dir(RutaFoto & DBCodigoEmpleado.Text & ".gif") <> "" Then
'           Destino = RutaFoto & DBCodigoEmpleado.Text & ".gif"
'        ElseIf Dir(RutaFoto & DBCodigoEmpleado.Text & ".bmp") <> "" Then
'           Destino = RutaFoto & DBCodigoEmpleado.Text & ".bmp"
'        End If
'
'        If (Dir(Destino) <> "") Then
'         Image1.Picture = LoadPicture(Destino)
'        Else
'          Destino = App.Path + "\Zw.bmp"
'         Image1.Picture = LoadPicture(Destino)
'        End If
'
'        If DtaEmpleado.Recordset("PorcientoIncentivo") = 0 Then
'         Me.Check1.Value = 0
'         Me.TxtPorcientoHora.Text = 0
'         Me.TxtPorcientoHora.Visible = False
'        Else
'         Me.Check1.Value = 1
'         Me.TxtPorcientoHora.Text = DtaEmpleado.Recordset("PorcientoIncentivo")
'         Me.TxtPorcientoHora.Visible = True
'        End If
'
'        If Not IsNull(DtaEmpleado.Recordset("SueldoActualBasico")) = True Then
'         If DtaEmpleado.Recordset("SueldoActualBasico") = True Then
'          Me.ChkSueldoActual.Value = 1
'        Else
'          Me.ChkSueldoActual.Value = 0
'         End If
'        End If
'
'        If Not IsNull(DtaEmpleado.Recordset("HorasTurno")) = True Then
'         If DtaEmpleado.Recordset("HorasTurno") = True Then
'          Me.ChkHorasTurno.Value = 1
'        Else
'          Me.ChkHorasTurno.Value = 0
'         End If
'        End If
'
'        If Not IsNull(DtaEmpleado.Recordset("CuentaBanco")) Then
'        Me.TxtCuentaBanco.Text = DtaEmpleado.Recordset("CuentaBanco")
'        End If
'
'        If Not IsNull(DtaEmpleado.Recordset("numcedula")) Then
'        TxtNumCedula.Text = DtaEmpleado.Recordset("numcedula")
'        End If
'        ChkSuspendido.Visible = True
'        txtNombre1.Text = DtaEmpleado.Recordset("Nombre1")
'
'        If Not IsNull(DtaEmpleado.Recordset("Nombre2")) Then
'        txtNombre2.Text = DtaEmpleado.Recordset("Nombre2")
'        End If
'        If Not IsNull(DtaEmpleado.Recordset("Apellido1")) Then
'          txtApellido1.Text = DtaEmpleado.Recordset("Apellido1")
'        End If
'        If Not IsNull(DtaEmpleado.Recordset("Apellido2")) Then
'         txtApellido2.Text = DtaEmpleado.Recordset("Apellido2")
'        End If
'        If Not IsNull(DtaEmpleado.Recordset("Direccion")) Then
'           TxtDireccion.Text = DtaEmpleado.Recordset("Direccion")
'        End If
'        If Not IsNull(DtaEmpleado.Recordset("Nacionalidad")) Then
'         TxtNacionalidad.Text = DtaEmpleado.Recordset("Nacionalidad")
'        End If
'        If Not IsNull(DtaEmpleado.Recordset("Codgrupo")) Then
'            TxtCodGrupo = DtaEmpleado.Recordset("Codgrupo")
'        Else
'            TxtCodGrupo = ""
'            DBCGrupo.Text = ""
'        End If
'        If Not IsNull(DtaEmpleado.Recordset("CodigoPostal")) Then
'          TxtCodPostal.Text = DtaEmpleado.Recordset("CodigoPostal")
'        End If
'        If Not IsNull(DtaEmpleado.Recordset("sexo")) Then
'          CmbSexo.Text = DtaEmpleado.Recordset("sexo")
'        End If
'        If Not IsNull(DtaEmpleado.Recordset("NumeroInss")) Then
'        TxtNInss.Text = DtaEmpleado.Recordset("NumeroInss")
'        End If
'        If Not IsNull(DtaEmpleado.Recordset("CodDepartamento")) Then
'        txtCodDepartamento.Text = DtaEmpleado.Recordset("CodDepartamento")
'        End If
'        If Not IsNull(DtaEmpleado.Recordset("CodCargo")) Then
'          TxtCodCargo.Text = DtaEmpleado.Recordset("CodCargo")
'        End If
'        If Not IsNull(DtaEmpleado.Recordset("Sindicalista")) Then
'          CmbSindicalista.Text = DtaEmpleado.Recordset("Sindicalista")
'        End If
'        If Not IsNull(DtaEmpleado.Recordset("numhijos")) Then
'          TxtNumHijos.Text = DtaEmpleado.Recordset("numhijos")
'        End If
'        frmEmpleado.Caption = "Registro del Empleado: " & DBCodigoEmpleado.Text & ": " & txtNombre1.Text & " " & txtNombre2.Text & " " & txtApellido1.Text & " " & txtApellido2.Text
''        Me.CmdAcercade.Caption = DBCodigoEmpleado.Text & ":   " & txtNombre1.Text & " " & txtNombre2.Text & " " & txtApellido1.Text & " " & txtApellido2.Text
''        Me.xp_canvas1.Caption = "Registro del Empleado: " & DBCodigoEmpleado.Text & ": " & TxtNombre1.Text & " " & txtNombre2.Text & " " & TxtApellido1.Text & " " & txtApellido2.Text
'
'        If Not IsNull(DtaEmpleado.Recordset("DiasDescuento")) Then
'            TxtDiasDescuento.Text = DtaEmpleado.Recordset("DiasDescuento")
'        Else
'            TxtDiasDescuento.Text = 0
'        End If
'        Bandera = False
'
'        If DtaEmpleado.Recordset("SalarioFijo") = "S" Then
'          Salario = True
'        Else
'          Salario = False
'        End If
'
'        If DtaEmpleado.Recordset("ausente") = True Then
'           ChkSuspendido.Value = 1
'           LblSuspendido.Visible = True
'        Else
'           LblSuspendido.Visible = False
'           ChkSuspendido.Value = 0
'        End If
'
'        If DtaEmpleado.Recordset("salariominimo") = True Then
'            CmbSalarioMinimo.Text = "Verdaderp"
'        Else
'           CmbSalarioMinimo.Text = "Falso"
'        End If
'
'        If DtaEmpleado.Recordset("ExentoInss") = True Then
'            CmbExentoInss.Text = "Verdadero"
'        Else
'           CmbExentoInss.Text = "Falso"
'        End If
'
'        If DtaEmpleado.Recordset("ExentoIr") = True Then
'            CmbExentoIr.Text = "Verdadero"
'        Else
'           CmbExentoIr.Text = "Falso"
'        End If
'
'        If DtaEmpleado.Recordset("PagoInssPatronal") = True Then
'            CmbPagoInssPatronal.Text = "Verdadero"
'        Else
'           CmbPagoInssPatronal.Text = "Falso"
'        End If
'
'
'        If DtaEmpleado.Recordset("Dolarizado") = True Then
'           Me.ChkDolarizado.Value = xtpChecked
'        Else
'           Me.ChkDolarizado.Value = xtpUnchecked
'        End If
'        Bandera = True
'
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = True
'    SSTab1.TabEnabled(2) = True
'    SSTab1.TabEnabled(3) = True
'    SSTab1.TabEnabled(4) = True
'    SSTab1.TabEnabled(5) = True
'    SSTab1.TabEnabled(6) = True
'
'    ' datos de la Nómina
'
''no olvidar los valores nomina
'
'        DtaTipoNomina.Refresh
'        Do While Not DtaTipoNomina.Recordset.EOF
'           If DtaTipoNomina.Recordset("CodTipoNomina") = DtaEmpleado.Recordset("CodTipoNomina") Then
'              DBCTipoNomina.Text = DtaTipoNomina.Recordset("nomina")
'              Exit Do
'            End If
'        DtaTipoNomina.Recordset.MoveNext
'        Loop
'
'            If Not IsNull(DtaEmpleado.Recordset("SueldoPeriodo")) Then
'            TxtSueldoPeriodo.Text = DtaEmpleado.Recordset("SueldoPeriodo")
'            End If
'
'        If Not IsNull(DtaEmpleado.Recordset("TarifaHoraria")) Then
'            txtTarifaHoraria.Text = DtaEmpleado.Recordset("TarifaHoraria")
'        End If
'
'        If Not IsNull(DtaEmpleado.Recordset("PorcentajeComision")) Then
'            TxtComision.Text = DtaEmpleado.Recordset("PorcentajeComision")
'        End If
'
'
'       If Not IsNull(DtaEmpleado.Recordset("OtrosIngresos")) Then
'          txtOtrosIngresos.Text = DtaEmpleado.Recordset("OtrosIngresos")
'       End If
'
'       If Not IsNull(DtaEmpleado.Recordset("DescripOtrIngre")) Then
'          TxtDescripOtrIngre.Text = DtaEmpleado.Recordset("DescripOtrIngre")
'       End If
'
'
'
'    Evaluar = True
'    Me.DtaHistorico.RecordSource = "SELECT  Codempleado, FechaBaja, MotivoBaja, FechaAumento, MotivoAumento, FechaInicialSusp, FechaFinalSusp, MotivoSuspencion, FechaNacimiento, FechaContrato,FechaContratoVac , CargoInicial, CargoActual, CargoAnterior, SueldoInicial, SueldoAnterior, SueldoActual, CuentaDebito, CuentaCredito,CuentaPrestamo,CuentaOtrosIngresos,CuentaINSS,CuentaIR From Historico Where (CodEmpleado = " & CodEmpleado & " )"
'    DtaHistorico.Refresh
'    Do While Not DtaHistorico.Recordset.EOF
'         If DtaHistorico.Recordset("CodEmpleado") = CodEmpleado Then
'            If Not IsNull(DtaHistorico.Recordset("FechaNacimiento")) Then
'                MaskEdNacimiento.Value = Format(DtaHistorico.Recordset("FechaNacimiento"), "dd/mm/yyyy")
'            End If
'
'            If Not IsNull(DtaHistorico.Recordset("FechaContratoVac")) Then
'                Me.DTPFechaContratoVaca.Value = Format(DtaHistorico.Recordset("FechaContratoVac"), "dd/mm/yyyy")
'            End If
'
'            If Not IsNull(DtaHistorico.Recordset("FechaContrato")) Then
'                MaskEdContrato.Value = Format(DtaHistorico.Recordset("FechaContrato"), "dd/mm/yyyy")
'            End If
'
'            If Not IsNull(DtaHistorico.Recordset("CargoInicial")) Then
'              DBCargoInicial.Text = DtaHistorico.Recordset("CargoInicial")
'            End If
'            If Not IsNull(DtaHistorico.Recordset("CargoAnterior")) Then
'               DBCargoAnterior.Text = DtaHistorico.Recordset("CargoAnterior")
'            End If
'            If Not IsNull(DtaHistorico.Recordset("CargoActual")) Then
'                 DBCargoActual.Text = DtaHistorico.Recordset("CargoActual")
'            End If
'            If Not IsNull(DtaHistorico.Recordset("MOTIVOBAJA")) Then
'                  TxtMotivoBaja.Text = DtaHistorico.Recordset("MOTIVOBAJA")
'            End If
'            If Not IsNull(DtaHistorico.Recordset("MotivoAumento")) Then
'                 TxtMotivoAumento.Text = DtaHistorico.Recordset("MotivoAumento")
'            End If
'            If Not IsNull(DtaHistorico.Recordset("MotivoSuspencion")) Then
'                 TxtMotivoSuspencion.Text = DtaHistorico.Recordset("MotivoSuspencion")
'            End If
'
'            TxtSueldoInicial.Text = Format((DtaHistorico.Recordset("SueldoInicial")), "##,##0.00")
'            TxtSueldoAnterior.Text = Format((DtaHistorico.Recordset("SueldoAnterior")), "##,##0.00")
'            TxtSueldoActual.Text = Format((DtaHistorico.Recordset("SueldoActual")), "##,##0.00")
'
'            If Not IsNull(DtaHistorico.Recordset("fechabaja")) Then
'                 MaskEdBaja.Text = DtaHistorico.Recordset("fechabaja")
'            End If
'
'            If Not IsNull(DtaHistorico.Recordset("FechaAumento")) Then
'                 MaskEdAumento.Text = Format(DtaHistorico.Recordset("FechaAumento"), "dd/mm/yyyy")
'            End If
'
'             If Not IsNull(DtaHistorico.Recordset("FechaInicialSusp")) Then
'                MaskEdSuspencion.Text = DtaHistorico.Recordset("FechaInicialSusp")
'            End If
'
'            If Not IsNull(DtaHistorico.Recordset("FechaInicialSusp")) Then
'               MaskEdFinalSusp.Text = DtaHistorico.Recordset("FechaInicialSusp")
'            End If
'
''            If Not IsNull(DtaHistorico.Recordset("CuentaDebito")) Then
''             TxtDebito.Text = DtaHistorico.Recordset("CuentaDebito")
''            End If
''
''            If Not IsNull(DtaHistorico.Recordset("cuentacredito")) Then
''             TxtCredito.Text = DtaHistorico.Recordset("cuentacredito")
''            End If
''
''            If Not IsNull(DtaHistorico.Recordset("CuentaPrestamo")) Then
''             Me.TxtCtaPrestamo.Text = DtaHistorico.Recordset("CuentaPrestamo")
''            End If
''
''            If Not IsNull(DtaHistorico.Recordset("CuentaOtrosIngresos")) Then
''             Me.TxtCtaOtrosIngresos.Text = DtaHistorico.Recordset("CuentaOtrosIngresos")
''            End If
'
''            If Not IsNull(DtaHistorico.Recordset("CuentaINSS")) Then
''             Me.TxtCuentaInss.Text = DtaHistorico.Recordset("CuentaINSS")
''            End If
''
''            If Not IsNull(DtaHistorico.Recordset("CuentaIR")) Then
''             Me.TxtCuentaIR.Text = DtaHistorico.Recordset("CuentaIR")
''            End If
'
'            Exit Do
'       End If
'           DtaHistorico.Recordset.MoveNext
'       Loop
'
'    DtaInfNomina.Refresh
'    Do While Not DtaInfNomina.Recordset.EOF
'         If DtaInfNomina.Recordset("CodEmpleado") = CodEmpleado Then
'            If Not IsNull(DtaInfNomina.Recordset("salariominimo")) Then
'                 CmbSalarioMinimo.Text = DtaInfNomina.Recordset("salariominimo")
'            End If
'            If Not IsNull(DtaInfNomina.Recordset("ExentoInss")) Then
'                 CmbExentoInss.Text = DtaInfNomina.Recordset("ExentoInss")
'            End If
'            If Not IsNull(DtaInfNomina.Recordset("ExentoIr")) Then
'                 CmbExentoIr.Text = DtaInfNomina.Recordset("ExentoIr")
'            End If
'
'            If Not IsNull(DtaInfNomina.Recordset("PagoInssPatronal")) Then
'                CmbPagoInssPatronal.Text = DtaInfNomina.Recordset("PagoInssPatronal")
'            End If
'
'            TxtCodTipoNomina.Text = DtaInfNomina.Recordset("CodTipoNomina")
'
'            Evaluar = True
'            Exit Do
'
'         End If
'           DtaInfNomina.Recordset.MoveNext
'       Loop
'
'      DtaDepartamento.Refresh
'       Do While Not DtaDepartamento.Recordset.EOF
'         If DtaDepartamento.Recordset("CodDepartamento") = txtCodDepartamento.Text Then
'            DBCDepartamento.Text = DtaDepartamento.Recordset("departamento")
'            Exit Do
'         Else
'          'DBCDepartamento.Text = ""
'         End If
'           DtaDepartamento.Recordset.MoveNext
'       Loop
'
'       DtaCargo.Refresh
'       Do While Not DtaCargo.Recordset.EOF
'         If DtaCargo.Recordset("CodCargo") = TxtCodCargo.Text Then
'            DBCCargo.Text = DtaCargo.Recordset("Cargo")
'            Exit Do
'         Else
'            DBCCargo.Text = ""
'         End If
'           DtaCargo.Recordset.MoveNext
'       Loop
'
'      Evaluar = True
'     CmbIncapacidad.Text = "No"
'      DtaIncapacidades.Refresh
'       Do While Not DtaIncapacidades.Recordset.EOF
'         If DtaIncapacidades.Recordset("CodEmpleado") = CodEmpleado Then
'            CmbIncapacidad.Text = "Si"
'            Exit Do
'         Else
'            CmbIncapacidad.Text = "No"
'         End If
'           DtaIncapacidades.Recordset.MoveNext
'       Loop
'    Evaluar = True
'       Salida = False
'    End If
'
'
'
'
'    'realizo los sql's de los incentivos y de las deducciones
'
'    CodEmpleado1 = DBCodigoEmpleado.Text
'
'    SQlIncentivos = "SELECT Incentivo.NumIncentivo, TipoIncentivo.Incentivo, Incentivo.CodEmpleado, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado FROM TipoIncentivo INNER JOIN (Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo) ON TipoIncentivo.CodTipoIncentivo = Incentivo.CodTipoIncentivo WHERE (Incentivo.CodEmpleado= " & CodEmpleado & ") And (DetalleIncentivo.Pagado = " & 0 & ")"
'    DtaDetalleIncentivo.RecordSource = SQlIncentivos
'    DtaDetalleIncentivo.Refresh
'
'    DbGIncentivos.Columns(0).Visible = False
'    DbGIncentivos.Columns(2).Visible = False
'    DbGIncentivos.Columns(5).Visible = False
'
'    SQlDeducciones = "SELECT Deduccion.NumDeduccion, TipoDeduccion.Deduccion, Deduccion.CodEmpleado, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado FROM TipoDeduccion INNER JOIN (Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion) ON (TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion) WHERE Deduccion.CodEmpleado=" & CodEmpleado & " AND DetalleDeduccion.Pagado= " & 0 & " "
'    DtaDetalleDeduccion.RecordSource = SQlDeducciones
'    DtaDetalleDeduccion.Refresh
'
'    DbgDeducciones.Columns(0).Visible = False
'    DbgDeducciones.Columns(2).Visible = False
'    DbgDeducciones.Columns(5).Visible = False
'    SQlPrestamo = "SELECT NumPrestamo, CuentaDebito, CuentaCredito, Monto, CantCuotas, Interes, Saldo, FechaInicial, Cancelado, Moneda, CuotasIguales, CodEmpleado From Prestamo WHERE Prestamo.CodEmpleado=" & CodEmpleado & " AND Prestamo.Cancelado=0"
'    DtaPrestamo.RecordSource = SQlPrestamo
'    DtaPrestamo.Refresh
'    If Not DtaPrestamo.Recordset.EOF Then
'    numeroPrestamo = Me.DtaPrestamo.Recordset("NumPrestamo")
'    Else
'     numeroPrestamo = -100
'    End If
'    SqlDetallePrestamo = "SELECT MovPrestamo.ID,MovPrestamo.NumPrestamo, MovPrestamo.NumCuota, MovPrestamo.Monto, MovPrestamo.Interes, MovPrestamo.CuotaIgual,MovPrestamo.SaldoCuota , MovPrestamo.Cancelado FROM Prestamo INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo Where (MovPrestamo.Cancelado = 0) And (MovPrestamo.NumPrestamo = " & numeroPrestamo & ")"
'    DtaMovPrestamo.RecordSource = SqlDetallePrestamo
'    DtaMovPrestamo.Refresh
'
'
'
'
'
'
'    If Not DtaPrestamo.Recordset.EOF Then
'       TxtCreditoPrestamo.Text = DtaPrestamo.Recordset("cuentacredito")
'       TxtDebitoPrestamo.Text = DtaPrestamo.Recordset("CuentaDebito")
'       Me.DbgrLibreta.Columns(0).Visible = False
'    DbgrLibreta.Columns(1).Visible = False
'    DbgrLibreta.Columns(7).Visible = False
'
'    Else
'       TxtCreditoPrestamo.Text = " "
'       TxtDebitoPrestamo.Text = " "
'    End If
'
'
'
'
'    SqlDetalleSubsidio = "SELECT Subsidio.NumSubsidio, Subsidio.CodEmpleado, Subsidio.CodTipoSubsidio, TipoSubsidio.Subsidio,DetalleSubsidio.Descripcion, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.Pagado FROM TipoSubsidio INNER JOIN (Subsidio INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio) ON TipoSubsidio.CodTipoSubsidio = Subsidio.CodTipoSubsidio WHERE DetalleSubsidio.Pagado=0 And Subsidio.CodEmpleado=" & CodEmpleado & " "
'    DtaDetalleSubsidio.RecordSource = SqlDetalleSubsidio
'    DtaDetalleSubsidio.Refresh
'
'    DbgrSubsidios.Columns(0).Visible = False
'    DbgrSubsidios.Columns(1).Visible = False
'    DbgrSubsidios.Columns(2).Visible = False
'    DbgrSubsidios.Columns(7).Visible = False
'    DbgrSubsidios.Columns(5).Width = 1200
'    DbgrSubsidios.Columns(6).Width = 500
'
'
'
'    If txtNombre1.Text = "" Then
'        SSTab1.TabEnabled(1) = True
'        SSTab1.TabEnabled(2) = True
'        SSTab1.TabEnabled(3) = True
'        SSTab1.TabEnabled(4) = True
'        SSTab1.TabEnabled(5) = True
'        SSTab1.TabEnabled(6) = True
'    End If
'
'          If Salario = True Then
'             Me.ChkSalarioFijo.Value = 1
'             Me.TxtComision.Enabled = False
'            Else
'             Me.ChkSalarioFijo.Value = 0
'             Me.TxtComision.Enabled = True
'            End If
'
'
'        Me.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo From Empleado WHERE     (CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') "
'        Me.DtaEmpleado.Refresh
'
'        If Not Me.DtaEmpleado.Recordset.EOF Then
'         DBCodigoEmpleado.Text = Me.DtaEmpleado.Recordset("CodEmpleado1")
'        End If
'
'        Me.DBCodigoEmpleado.Columns(0).Visible = False
'        Me.DBCodigoEmpleado.Columns(1).Caption = "Codigo"
'        Me.DBCodigoEmpleado.Columns(1).Width = 800
'        Me.DBCodigoEmpleado.Columns(2).Visible = False
'
'
'        frmEmpleado.MousePointer = 0
'        frmEmpleado.AutoRedraw = True
'Exit Sub
'TipoErrs:
' MsgBox Err.Description
' Unload Me
End Sub

Private Sub DBCodigoEmpleado_SelChange(Cancel As Integer)
Me.Text1.Text = Me.DBCodigoEmpleado.Text
End Sub

Private Sub DBCTipoNomina_Click(Area As Integer)

Salida = True
End Sub

Private Sub DBCTipoNomina_Change()
On Error GoTo TipoErrs
Dim TipoPago As String
Dim Moneda As String

PreparaSalida
DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
     If DtaTipoNomina.Recordset("nomina") = DBCTipoNomina.Text Then
        CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
        TxtCodTipoNomina.Text = DtaTipoNomina.Recordset("CodTipoNomina")
        CmbTipoPago.Text = DtaTipoNomina.Recordset("TipoPago")
        
        If DtaTipoNomina.Recordset("TipoPago") = "Salario Destajo" Then
           TxtSueldoPeriodo.Text = "0.00"
           TxtComision.Text = "0.00"
           TxtSueldoPeriodo.Enabled = False
           TxtTarifaHoraria.Enabled = True
           TxtComision.Enabled = False
        ElseIf DtaTipoNomina.Recordset("TipoPago") = "Salario Destajo y Comision" Then
           TxtSueldoPeriodo.Text = "0.00"
           TxtSueldoPeriodo.Enabled = False
           TxtTarifaHoraria.Enabled = True
           TxtComision.Enabled = True
        ElseIf DtaTipoNomina.Recordset("TipoPago") = "Salario Fijo" Then
           TxtTarifaHoraria.Text = "0.00"
           TxtComision.Text = "0.00"
           TxtSueldoPeriodo.Enabled = True
           TxtTarifaHoraria.Enabled = False
           TxtComision.Enabled = False
        ElseIf DtaTipoNomina.Recordset("TipoPago") = "Salario Fijo,Destajo y Comision" Then
           TxtSueldoPeriodo.Enabled = True
           TxtTarifaHoraria.Enabled = True
           TxtComision.Enabled = True
        Else
           TxtTarifaHoraria.Text = "0.00"
           TxtSueldoPeriodo.Enabled = True
           TxtTarifaHoraria.Enabled = False
           TxtComision.Enabled = True
           
       End If
        
        Exit Do
     Else
        'LimpiaEmpleado
     End If
       DtaTipoNomina.Recordset.MoveNext
   Loop
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub DBCTipoNomina_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  CmbTipoPago.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub DbgDeducciones_Click()
On Error GoTo TipoErrs
CmdEliminarDeduccion.Enabled = True
Exit Sub
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub DbGIncentivos_Click()

On Error GoTo TipoErrs
CmdEliminarIncentivo.Enabled = True
Me.CmdAnular.Enabled = True

Exit Sub
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub DbgrSubsidios_Click()
On Error GoTo TipoErrs
CmdEliminarSubsidio.Enabled = True
Exit Sub
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub DtaTipoSubsidio_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub EliminaFoto_Click()
On Error GoTo TipoErrs
 Respuesta = MsgBox("Esta seguro de Borrar esta Foto?", vbYesNo, "Borrando la Foto de: " & TxtNombre1.Text & TxtNombre2)
   If Respuesta = 6 Then
     If Dir(RutaFoto & DBCodigoEmpleado.Text & ".bmp") <> "" Then
     Destino = RutaFoto & DBCodigoEmpleado.Text & ".bmp"
     ElseIf Dir(RutaFoto & DBCodigoEmpleado.Text & ".gif") <> "" Then
     Destino = RutaFoto & DBCodigoEmpleado.Text & ".gif"
     ElseIf Dir(RutaFoto & DBCodigoEmpleado.Text & ".jpg") <> "" Then
     Destino = RutaFoto & DBCodigoEmpleado.Text & ".jpg"
     End If
     If Destino = RutaFoto & "\Zw.bmp" Then
        MsgBox "El Empleado no tiene una foto agregada"
        Exit Sub
     End If
     
     If (Dir(Destino) <> "") Then
         Kill Destino
         Destino = RutaFoto & "\Zw.bmp"
         Image1.Picture = LoadPicture(Destino)
        Else
         MsgBox "No existe Ninguna Foto que eliminar", vbCritical, "Error:Sistema de Nominas"
       End If
   End If
Exit Sub
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Public Sub CargarDatos()
'On Error GoTo TipoErrs
Dim SQlIncentivos As String, SQlDeducciones As String, SqlDetallePrestamo As String, SQlPrestamo As String, SqlDetalleSubsidio As String
Dim Salario As Boolean
Dim CodEmpleado1 As String
Dim numeroPrestamo As Double
Dim CodEmpleado As Double



 Evaluar = True
 'Al ejecutar algun cambio en el combo actualizo el nombre del Empleado
 frmEmpleado.MousePointer = 11

 LimpiaEmpleado
 LimpiaHistorico
 LimpiaInfNomina
 'LimpiaInfNomina
' DtaEmpleados.Refresh
 ChkSuspendido.Visible = False
'Busco el codigo del empleado para que automaticamente ubique el nombre
 'aunque no existe en la data consulta
 CodEmpleado = -1
 
Me.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo,Dolarizado,CuentaBanco,SueldoActualBasico,HorasTurno,SalPorcentaje, AumentoBasico From Empleado WHERE (CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') And (Activo = 1)"
Me.DtaEmpleado.Refresh

If Not Me.DtaEmpleado.Recordset.EOF Then
'Do While Not DtaEmpleado.Recordset.EOF
'     If DtaEmpleado.Recordset("CodEmpleado1") = DBCodigoEmpleado.Text Then
'        If DtaEmpleado.Recordset("activo") = False Then
'           MsgBox "Este empleado ya fue dado de Baja"
'        End If
        
     CodEmpleado = Me.DtaEmpleado.Recordset("CodEmpleado")
     TxtCodEmpleado.Text = Me.DtaEmpleado.Recordset("CodEmpleado")
     Me.Text1.Text = Me.DtaEmpleado.Recordset("CodEmpleado1")
        If Not IsNull(DtaEmpleado.Recordset("numeroruc")) Then
          TxtNRuc.Text = DtaEmpleado.Recordset("numeroruc")
        End If
        'busco el tipo del archivo
        'Destino = ""
        If Dir(RutaFoto & DBCodigoEmpleado.Text & ".jpg") <> "" Then
           Destino = RutaFoto & DBCodigoEmpleado.Text & ".jpg"
        ElseIf Dir(RutaFoto & DBCodigoEmpleado.Text & ".gif") <> "" Then
           Destino = RutaFoto & DBCodigoEmpleado.Text & ".gif"
        ElseIf Dir(RutaFoto & DBCodigoEmpleado.Text & ".bmp") <> "" Then
           Destino = RutaFoto & DBCodigoEmpleado.Text & ".bmp"
        End If
        
        If (Dir(Destino) <> "") Then
         Image1.Picture = LoadPicture(Destino)
        Else
          Destino = App.Path + "\Zw.bmp"
'          Destino = RutaLogo
         Image1.Picture = LoadPicture(Destino)
        End If
        
        If DtaEmpleado.Recordset("PorcientoIncentivo") = 0 Then
         Me.Check1.Value = 0
         Me.TxtPorcientoHora.Text = 0
         Me.TxtPorcientoHora.Visible = False
        Else
         Me.Check1.Value = 1
         Me.TxtPorcientoHora.Text = DtaEmpleado.Recordset("PorcientoIncentivo")
         Me.TxtPorcientoHora.Visible = True
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("SueldoActualBasico")) = True Then
         If DtaEmpleado.Recordset("SueldoActualBasico") = True Then
          Me.ChkSueldoActual.Value = 1
        Else
          Me.ChkSueldoActual.Value = 0
         End If
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("HorasTurno")) = True Then
            If DtaEmpleado.Recordset("HorasTurno") = True Then
             Me.ChkHorasTurno.Value = 1
            Else
             Me.ChkHorasTurno.Value = 0
            End If
        Else
          Me.ChkHorasTurno.Value = 0
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("CuentaBanco")) Then
        Me.TxtCuentaBanco.Text = DtaEmpleado.Recordset("CuentaBanco")
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("AumentoBasico")) Then
        Me.txtAumentoBasico.Text = DtaEmpleado.Recordset("AumentoBasico")
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("numcedula")) Then
        TxtNumCedula.Text = DtaEmpleado.Recordset("numcedula")
        End If
        ChkSuspendido.Visible = True
        TxtNombre1.Text = DtaEmpleado.Recordset("Nombre1")
        
        If Not IsNull(DtaEmpleado.Recordset("SalPorcentaje")) Then
         Me.TxtSalarioPorciento.Text = DtaEmpleado.Recordset("SalPorcentaje")
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("Nombre2")) Then
        TxtNombre2.Text = DtaEmpleado.Recordset("Nombre2")
        End If
        If Not IsNull(DtaEmpleado.Recordset("Apellido1")) Then
          TxtApellido1.Text = DtaEmpleado.Recordset("Apellido1")
        End If
        If Not IsNull(DtaEmpleado.Recordset("Apellido2")) Then
         TxtApellido2.Text = DtaEmpleado.Recordset("Apellido2")
        End If
        If Not IsNull(DtaEmpleado.Recordset("Direccion")) Then
           TxtDireccion.Text = DtaEmpleado.Recordset("Direccion")
        End If
        If Not IsNull(DtaEmpleado.Recordset("Nacionalidad")) Then
         TxtNacionalidad.Text = DtaEmpleado.Recordset("Nacionalidad")
        End If
        If Not IsNull(DtaEmpleado.Recordset("Codgrupo")) Then
            TxtCodGrupo = DtaEmpleado.Recordset("Codgrupo")
        Else
            TxtCodGrupo = ""
            DBCGrupo.Text = ""
        End If
        If Not IsNull(DtaEmpleado.Recordset("CodigoPostal")) Then
          TxtCodPostal.Text = DtaEmpleado.Recordset("CodigoPostal")
        End If
        If Not IsNull(DtaEmpleado.Recordset("sexo")) Then
          CmbSexo.Text = DtaEmpleado.Recordset("sexo")
        End If
        If Not IsNull(DtaEmpleado.Recordset("NumeroInss")) Then
        TxtNInss.Text = DtaEmpleado.Recordset("NumeroInss")
        End If
        If Not IsNull(DtaEmpleado.Recordset("CodDepartamento")) Then
        TxtCodDepartamento.Text = DtaEmpleado.Recordset("CodDepartamento")
        End If
        If Not IsNull(DtaEmpleado.Recordset("CodCargo")) Then
          TxtCodCargo.Text = DtaEmpleado.Recordset("CodCargo")
        End If
        If Not IsNull(DtaEmpleado.Recordset("Sindicalista")) Then
          CmbSindicalista.Text = DtaEmpleado.Recordset("Sindicalista")
        End If
        If Not IsNull(DtaEmpleado.Recordset("numhijos")) Then
          TxtNumHijos.Text = DtaEmpleado.Recordset("numhijos")
        End If
        frmEmpleado.Caption = "Registro del Empleado: " & DBCodigoEmpleado.Text & ": " & TxtNombre1.Text & " " & TxtNombre2.Text & " " & TxtApellido1.Text & " " & TxtApellido2.Text
'        Me.CmdAcercade.Caption = DBCodigoEmpleado.Text & ":   " & txtNombre1.Text & " " & txtNombre2.Text & " " & txtApellido1.Text & " " & txtApellido2.Text
'        Me.xp_canvas1.Caption = "Registro del Empleado: " & DBCodigoEmpleado.Text & ": " & TxtNombre1.Text & " " & txtNombre2.Text & " " & TxtApellido1.Text & " " & txtApellido2.Text
        
        If Not IsNull(DtaEmpleado.Recordset("DiasDescuento")) Then
            TxtDiasDescuento.Text = DtaEmpleado.Recordset("DiasDescuento")
        Else
            TxtDiasDescuento.Text = 0
        End If
        Bandera = False
        
        If DtaEmpleado.Recordset("SalarioFijo") = "S" Then
          Salario = True
        Else
          Salario = False
        End If
        
        If DtaEmpleado.Recordset("ausente") = True Then
           ChkSuspendido.Value = 1
           LblSuspendido.Visible = True
        Else
           LblSuspendido.Visible = False
           ChkSuspendido.Value = 0
        End If
        
        If DtaEmpleado.Recordset("salariominimo") = True Then
            CmbSalarioMinimo.Text = "Verdaderp"
        Else
           CmbSalarioMinimo.Text = "Falso"
        End If
        
        If DtaEmpleado.Recordset("ExentoInss") = True Then
            CmbExentoInss.Text = "Verdadero"
        Else
           CmbExentoInss.Text = "Falso"
        End If
           
        If DtaEmpleado.Recordset("ExentoIr") = True Then
            CmbExentoIr.Text = "Verdadero"
        Else
           CmbExentoIr.Text = "Falso"
        End If
        
        If DtaEmpleado.Recordset("PagoInssPatronal") = True Then
            CmbPagoInssPatronal.Text = "Verdadero"
        Else
           CmbPagoInssPatronal.Text = "Falso"
        End If
        
        
        If DtaEmpleado.Recordset("Dolarizado") = True Then
           Me.ChkDolarizado.Value = xtpChecked
        Else
           Me.ChkDolarizado.Value = xtpUnchecked
        End If
        Bandera = True
    
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    SSTab1.TabEnabled(4) = True
    SSTab1.TabEnabled(5) = True
    SSTab1.TabEnabled(6) = True

    ' datos de la Nómina

'no olvidar los valores nomina

        DtaTipoNomina.Refresh
        Do While Not DtaTipoNomina.Recordset.EOF
           If DtaTipoNomina.Recordset("CodTipoNomina") = DtaEmpleado.Recordset("CodTipoNomina") Then
              DBCTipoNomina.Text = DtaTipoNomina.Recordset("nomina")
              Exit Do
            End If
        DtaTipoNomina.Recordset.MoveNext
        Loop
        
            If Not IsNull(DtaEmpleado.Recordset("SueldoPeriodo")) Then
            TxtSueldoPeriodo.Text = DtaEmpleado.Recordset("SueldoPeriodo")
            End If
            
        If Not IsNull(DtaEmpleado.Recordset("TarifaHoraria")) Then
            TxtTarifaHoraria.Text = DtaEmpleado.Recordset("TarifaHoraria")
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("PorcentajeComision")) Then
            TxtComision.Text = DtaEmpleado.Recordset("PorcentajeComision")
        End If
      
     
       If Not IsNull(DtaEmpleado.Recordset("OtrosIngresos")) Then
          TxtOtrosIngresos.Text = DtaEmpleado.Recordset("OtrosIngresos")
       End If
       
       If Not IsNull(DtaEmpleado.Recordset("DescripOtrIngre")) Then
          TxtDescripOtrIngre.Text = DtaEmpleado.Recordset("DescripOtrIngre")
       End If
'        Exit Do
        Else 'si no lo encuentra
  
              SSTab1.TabEnabled(0) = True
              SSTab1.TabEnabled(1) = True
              SSTab1.TabEnabled(2) = True
              SSTab1.TabEnabled(3) = False
              SSTab1.TabEnabled(4) = False
              SSTab1.TabEnabled(5) = False
              SSTab1.TabEnabled(6) = False
        'frmEmpleado.Caption = "Registro del Empleado: " & TxtNombre1.Text & " " & TxtNombre2.Text & " " & TxtApellido1.Text & " " & TxtApellido2.Text
     End If
'DtaEmpleado.Recordset.MoveNext
'Loop

Evaluar = True
Me.DtaHistorico.RecordSource = "SELECT  Codempleado, FechaBaja, MotivoBaja, FechaAumento, MotivoAumento, FechaInicialSusp, FechaFinalSusp, MotivoSuspencion, FechaNacimiento, FechaContrato,FechaContratoVac , CargoInicial, CargoActual, CargoAnterior, SueldoInicial, SueldoAnterior, SueldoActual, CuentaDebito, CuentaCredito,CuentaPrestamo,CuentaOtrosIngresos,CuentaINSS,CuentaIR From Historico Where (CodEmpleado = " & CodEmpleado & " )"
DtaHistorico.Refresh
Do While Not DtaHistorico.Recordset.EOF
     If DtaHistorico.Recordset("CodEmpleado") = CodEmpleado Then
        If Not IsNull(DtaHistorico.Recordset("FechaNacimiento")) Then
            MaskEdNacimiento.Value = Format(DtaHistorico.Recordset("FechaNacimiento"), "dd/mm/yyyy")
        End If
        
        If Not IsNull(DtaHistorico.Recordset("FechaContratoVac")) Then
            Me.DTPFechaContratoVaca.Value = Format(DtaHistorico.Recordset("FechaContratoVac"), "dd/mm/yyyy")
        End If
        
        If Not IsNull(DtaHistorico.Recordset("FechaContrato")) Then
            MaskEdContrato.Value = Format(DtaHistorico.Recordset("FechaContrato"), "dd/mm/yyyy")
        End If
        
        If Not IsNull(DtaHistorico.Recordset("CargoInicial")) Then
          DBCargoInicial.Text = DtaHistorico.Recordset("CargoInicial")
        End If
        If Not IsNull(DtaHistorico.Recordset("CargoAnterior")) Then
           DBCargoAnterior.Text = DtaHistorico.Recordset("CargoAnterior")
        End If
        If Not IsNull(DtaHistorico.Recordset("CargoActual")) Then
             DBCargoActual.Text = DtaHistorico.Recordset("CargoActual")
        End If
        If Not IsNull(DtaHistorico.Recordset("MOTIVOBAJA")) Then
              TxtMotivoBaja.Text = DtaHistorico.Recordset("MOTIVOBAJA")
        End If
        If Not IsNull(DtaHistorico.Recordset("MotivoAumento")) Then
             TxtMotivoAumento.Text = DtaHistorico.Recordset("MotivoAumento")
        End If
        If Not IsNull(DtaHistorico.Recordset("MotivoSuspencion")) Then
             TxtMotivoSuspencion.Text = DtaHistorico.Recordset("MotivoSuspencion")
        End If
        
        TxtSueldoInicial.Text = Format((DtaHistorico.Recordset("SueldoInicial")), "##,##0.00")
        TxtSueldoAnterior.Text = Format((DtaHistorico.Recordset("SueldoAnterior")), "##,##0.00")
        TxtSueldoActual.Text = Format((DtaHistorico.Recordset("SueldoActual")), "##,##0.00")
        
        If Not IsNull(DtaHistorico.Recordset("fechabaja")) Then
             MaskEdBaja.Text = DtaHistorico.Recordset("fechabaja")
        End If
        
        If Not IsNull(DtaHistorico.Recordset("FechaAumento")) Then
             MaskEdAumento.Text = Format(DtaHistorico.Recordset("FechaAumento"), "dd/mm/yyyy")
        End If
         
         If Not IsNull(DtaHistorico.Recordset("FechaInicialSusp")) Then
            MaskEdSuspencion.Text = DtaHistorico.Recordset("FechaInicialSusp")
        End If
        
        If Not IsNull(DtaHistorico.Recordset("FechaInicialSusp")) Then
           MaskEdFinalSusp.Text = DtaHistorico.Recordset("FechaInicialSusp")
        End If
        
'        If Not IsNull(DtaHistorico.Recordset("CuentaDebito")) Then
'         TxtDebito.Text = DtaHistorico.Recordset("CuentaDebito")
'        End If
'
'        If Not IsNull(DtaHistorico.Recordset("cuentacredito")) Then
'         TxtCredito.Text = DtaHistorico.Recordset("cuentacredito")
'        End If
        
'        If Not IsNull(DtaHistorico.Recordset("CuentaPrestamo")) Then
'         Me.TxtCtaPrestamo.Text = DtaHistorico.Recordset("CuentaPrestamo")
'        End If
'
'        If Not IsNull(DtaHistorico.Recordset("CuentaOtrosIngresos")) Then
'         Me.TxtCtaOtrosIngresos.Text = DtaHistorico.Recordset("CuentaOtrosIngresos")
'        End If
'
'        If Not IsNull(DtaHistorico.Recordset("CuentaINSS")) Then
'         Me.TxtCuentaInss.Text = DtaHistorico.Recordset("CuentaINSS")
'        End If
'
'        If Not IsNull(DtaHistorico.Recordset("CuentaIR")) Then
'         Me.TxtCuentaIR.Text = DtaHistorico.Recordset("CuentaIR")
'        End If
        
        Exit Do
   End If
       DtaHistorico.Recordset.MoveNext
   Loop

DtaInfNomina.Refresh
Do While Not DtaInfNomina.Recordset.EOF
     If DtaInfNomina.Recordset("CodEmpleado") = CodEmpleado Then
        If Not IsNull(DtaInfNomina.Recordset("salariominimo")) Then
             CmbSalarioMinimo.Text = DtaInfNomina.Recordset("salariominimo")
        End If
        If Not IsNull(DtaInfNomina.Recordset("ExentoInss")) Then
             CmbExentoInss.Text = DtaInfNomina.Recordset("ExentoInss")
        End If
        If Not IsNull(DtaInfNomina.Recordset("ExentoIr")) Then
             CmbExentoIr.Text = DtaInfNomina.Recordset("ExentoIr")
        End If
       
        If Not IsNull(DtaInfNomina.Recordset("PagoInssPatronal")) Then
            CmbPagoInssPatronal.Text = DtaInfNomina.Recordset("PagoInssPatronal")
        End If
       
        TxtCodTipoNomina.Text = DtaInfNomina.Recordset("CodTipoNomina")
          
        Evaluar = True
        Exit Do
 
     End If
       DtaInfNomina.Recordset.MoveNext
   Loop

  DtaDepartamento.Refresh
   Do While Not DtaDepartamento.Recordset.EOF
     If DtaDepartamento.Recordset("CodDepartamento") = TxtCodDepartamento.Text Then
        DBCDepartamento.Text = DtaDepartamento.Recordset("departamento")
        Exit Do
     Else
      'DBCDepartamento.Text = ""
     End If
       DtaDepartamento.Recordset.MoveNext
   Loop

   DtaCargo.Refresh
   Do While Not DtaCargo.Recordset.EOF
     If DtaCargo.Recordset("CodCargo") = TxtCodCargo.Text Then
        DBCCargo.Text = DtaCargo.Recordset("Cargo")
        Exit Do
     Else
        DBCCargo.Text = ""
     End If
       DtaCargo.Recordset.MoveNext
   Loop
   
  Evaluar = True
 CmbIncapacidad.Text = "No"
  DtaIncapacidades.Refresh
   Do While Not DtaIncapacidades.Recordset.EOF
     If DtaIncapacidades.Recordset("CodEmpleado") = CodEmpleado Then
        CmbIncapacidad.Text = "Si"
        Exit Do
     Else
        CmbIncapacidad.Text = "No"
     End If
       DtaIncapacidades.Recordset.MoveNext
   Loop
Evaluar = True
   Salida = False

   



'realizo los sql's de los incentivos y de las deducciones

CodEmpleado1 = DBCodigoEmpleado.Text

SQlIncentivos = "SELECT Incentivo.NumIncentivo, TipoIncentivo.Incentivo, Incentivo.CodEmpleado, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado FROM TipoIncentivo INNER JOIN (Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo) ON TipoIncentivo.CodTipoIncentivo = Incentivo.CodTipoIncentivo WHERE (Incentivo.CodEmpleado= " & CodEmpleado & ") And (DetalleIncentivo.Pagado = " & 0 & ")"
DtaDetalleIncentivo.RecordSource = SQlIncentivos
DtaDetalleIncentivo.Refresh

DbGIncentivos.Columns(0).Visible = False
DbGIncentivos.Columns(2).Visible = False
DbGIncentivos.Columns(5).Visible = False

SQlDeducciones = "SELECT Deduccion.NumDeduccion, TipoDeduccion.Deduccion, Deduccion.CodEmpleado, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado FROM TipoDeduccion INNER JOIN (Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion) ON (TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion) WHERE Deduccion.CodEmpleado=" & CodEmpleado & " AND DetalleDeduccion.Pagado= " & 0 & " "
DtaDetalleDeduccion.RecordSource = SQlDeducciones
DtaDetalleDeduccion.Refresh

DbgDeducciones.Columns(0).Visible = False
DbgDeducciones.Columns(2).Visible = False
DbgDeducciones.Columns(5).Visible = False
SQlPrestamo = "SELECT NumPrestamo, CuentaDebito, CuentaCredito, Monto, CantCuotas, Interes, Saldo, FechaInicial, Cancelado, Moneda, CuotasIguales, CodEmpleado From Prestamo WHERE Prestamo.CodEmpleado=" & CodEmpleado & " AND Prestamo.Cancelado=0"
DtaPrestamo.RecordSource = SQlPrestamo
DtaPrestamo.Refresh
If Not DtaPrestamo.Recordset.EOF Then
numeroPrestamo = Me.DtaPrestamo.Recordset("NumPrestamo")
Else
 numeroPrestamo = -100
End If
SqlDetallePrestamo = "SELECT MovPrestamo.ID,MovPrestamo.NumPrestamo, MovPrestamo.NumCuota, MovPrestamo.Monto, MovPrestamo.Interes, MovPrestamo.CuotaIgual,MovPrestamo.SaldoCuota , MovPrestamo.Cancelado FROM Prestamo INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo Where (MovPrestamo.Cancelado = 0) And (MovPrestamo.NumPrestamo = " & numeroPrestamo & ")"
DtaMovPrestamo.RecordSource = SqlDetallePrestamo
DtaMovPrestamo.Refresh






If Not DtaPrestamo.Recordset.EOF Then
   TxtCreditoPrestamo.Text = DtaPrestamo.Recordset("cuentacredito")
   TxtDebitoPrestamo.Text = DtaPrestamo.Recordset("CuentaDebito")
   Me.DbgrLibreta.Columns(0).Visible = False
DbgrLibreta.Columns(1).Visible = False
DbgrLibreta.Columns(7).Visible = False

Else
   TxtCreditoPrestamo.Text = " "
   TxtDebitoPrestamo.Text = " "
End If




SqlDetalleSubsidio = "SELECT Subsidio.NumSubsidio, Subsidio.CodEmpleado, Subsidio.CodTipoSubsidio, TipoSubsidio.Subsidio,DetalleSubsidio.Descripcion, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.Pagado FROM TipoSubsidio INNER JOIN (Subsidio INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio) ON TipoSubsidio.CodTipoSubsidio = Subsidio.CodTipoSubsidio WHERE DetalleSubsidio.Pagado=0 And Subsidio.CodEmpleado=" & CodEmpleado & " "
DtaDetalleSubsidio.RecordSource = SqlDetalleSubsidio
DtaDetalleSubsidio.Refresh

DbgrSubsidios.Columns(0).Visible = False
DbgrSubsidios.Columns(1).Visible = False
DbgrSubsidios.Columns(2).Visible = False
DbgrSubsidios.Columns(7).Visible = False
DbgrSubsidios.Columns(5).Width = 1200
DbgrSubsidios.Columns(6).Width = 500



If TxtNombre1.Text = "" Then
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    SSTab1.TabEnabled(4) = True
    SSTab1.TabEnabled(5) = True
    SSTab1.TabEnabled(6) = True
End If

      If Salario = True Then
         Me.ChkSalarioFijo.Value = 1
         Me.TxtComision.Enabled = False
        Else
         Me.ChkSalarioFijo.Value = 0
         Me.TxtComision.Enabled = True
        End If


Me.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo From Empleado WHERE     (CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') "
Me.DtaEmpleado.Refresh

If Not Me.DtaEmpleado.Recordset.EOF Then
 DBCodigoEmpleado.Text = Me.DtaEmpleado.Recordset("CodEmpleado1")
End If

'Me.DBCodigoEmpleado.Columns(0).Visible = False
'Me.DBCodigoEmpleado.Columns(1).Caption = "Codigo"
'Me.DBCodigoEmpleado.Columns(1).Width = 800
'Me.DBCodigoEmpleado.Columns(2).Visible = False


frmEmpleado.MousePointer = 0
frmEmpleado.AutoRedraw = True
Exit Sub
TipoErrs:
 MsgBox Err.Description
' ControlErrores
 Unload Me
' End If

End Sub









Private Sub Form_Activate()
Salida = False
DtaDepartamento.Refresh
DtaCargo.Refresh
DtaTipoNomina.Refresh
DtaTipoIncentivo.Refresh
DtaTipoDeduccion.Refresh
DtaTipoSubsidio.Refresh
'DBCodigoEmpleado.Text = CodEmpleado
 'If Not BEmpleado = True Then
  ' frmEmpleado.CmdBorrar.Enabled = False
 'End If
 'If Not GEmpleado = True Then
  ' frmEmpleado.CmdGrabar.Enabled = False
 'End If
End Sub

Private Sub Form_Deactivate()
If Me.DBCodigoEmpleado.Text = "" Then
MsgBox "Se necesita el codigo del empleado", vbCritical, "Sistema Contable"
 Exit Sub
Else
'CodEmpleado = DBCodigoEmpleado.Text
End If
End Sub

Private Sub Form_Load()
Dim SqlSuspenciones As String
MDIPrimero.Skin1.ApplySkin hWnd
With Me.AdoNumerosDisponibles
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.AdoUser
   .ConnectionString = Conexion
End With

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaTurnos
   .ConnectionString = Conexion
   .RecordSource = "Turno"
   .Refresh
End With

With Me.DtaHorarioEmpleado
   .ConnectionString = Conexion
End With


With Me.DtaCargo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Cargo"
   .Refresh
End With

With Me.DtaConsecutivos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Consecutivos"
   .Refresh
End With

With Me.DtaDeduccion
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Deduccion"
End With

With Me.DtaDepartamento
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Departamento"
   .Refresh
End With

With Me.DtaDetalleDeduccion
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "DetalleDeduccion"
   .Refresh
End With

With Me.DtaDetalleDeduccion2
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "DetalleDeduccion"
   .Refresh
End With

With Me.DtaDetalleIncentivo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaDetalleIncentivo2
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "DetalleIncentivo"
   .Refresh
End With

With Me.DtaDetalleSubsidio
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtadetalleSubsidio2
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "DetalleSubsidio"
   .Refresh
End With

With Me.DtaEmpleados
   .ConnectionString = Conexion
   .RecordSource = "SELECT CodEmpleado, CodEmpleado1, Activo, Nombre1 + ' '+ Nombre2 +' '+Apellido1+' '+Apellido2 as Nombres From Empleado Where (Activo = 1) ORDER BY CodEmpleado1"
   .Refresh
End With

With Me.DtaEmpleado
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
'   .RecordSource = "Empleado"
'   .Refresh
End With

With Me.DtaGrupo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Grupo"
   .Refresh
End With

With Me.DtaHistorico
   .ConnectionString = Conexion
   .RecordSource = "Historico"
   .Refresh
End With

With Me.DtaIncapacidades
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Incapacidad"
   .Refresh
End With

With Me.DtaIncentivo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaInfNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Empleado"
   .Refresh
End With

With Me.DtaMovPrestamo
   '.DatabaseName = Ruta
    .ConnectionString = Conexion
End With

With Me.DtaPrestamo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaSubsidio
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Subsidio"
   .Refresh
End With

With Me.DtaSuspenciones
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaTipoDeduccion
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoDeduccion"
   .Refresh
End With

With Me.DtaTipoIncentivo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoIncentivo"
   .Refresh
End With

With Me.DtaTipoNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoNomina"
   .Refresh
End With

With Me.DtaTipoSubsidio
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoSubsidio"
End With


'Me.DBCodigoEmpleado.Columns(0).Visible = False
'Me.DBCodigoEmpleado.Columns(1).Caption = "Codigo"
'Me.DBCodigoEmpleado.Columns(1).Width = 800
'Me.DBCodigoEmpleado.Columns(2).Visible = False

Me.MaskEdNacimiento.Value = Now
Me.MaskEdContrato.Value = Now

frmEmpleado.CmdCerrar.MousePointer = 99
frmEmpleado.CmdGrabar.MousePointer = 99
frmEmpleado.MousePointer = 99
 ChkSuspendido.Visible = False
'objeto.TabVisible(ficha) [ = booleano ]
'objeto.TabEnabled(ficha)[ = booleano ]
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
SSTab1.TabEnabled(4) = False
SSTab1.TabEnabled(5) = False
SSTab1.TabEnabled(6) = False

Me.DTPFechaContratoVaca.Value = Format(Now, "dd/mm/yyyy")
'SqlSuspenciones = "SELECT Subsidios.CodEmpleado, Subsidios.Fechaini, Subsidios.FechaFin, Subsidios.Motivo, Subsidios.Activo, Subsidios.Ultimo From Subsidios WHERE (((Subsidios.Activo)=1))"
SqlSuspenciones = "SELECT Subsidios.CodEmpleado, Subsidios.Fechaini, Subsidios.FechaFin, Subsidios.Motivo, Subsidios.Activo, Subsidios.Ultimo From Subsidios WHERE (((Subsidios.Activo)=1))"
DtaSuspenciones.RecordSource = SqlSuspenciones
DtaSuspenciones.Refresh

Me.Top = 200
Me.Left = 1000



End Sub



Private Sub Form_LostFocus()
If DBCodigoEmpleado.Text <> "" Then
' CodEmpleado = DBCodigoEmpleado.Text
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo TipoErrs
If GEmpleado = True Then
ValidaSalida ("en la Tabla Empleado")
  If Contesta Then
    CmdGrabar.Value = True
    Salida = False
    Unload Me
  Else
    Salida = False
    Unload Me
  End If
End If

Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
  On Error GoTo TipoErrs
  If frmEmpleado.DBCodigoEmpleado.Text = "" Then
    MsgBox "Para Agregar Foto se Necesita el Codigo del Empleado", vbInformation, "Error:Sistema de Nominas"
    frmEmpleado.DBCodigoEmpleado.SetFocus
    frmDrag.CmdSalir.Value = True
    Exit Sub
  End If
    
    ' Obtiene las tres últimas letras del nombre del archivo arrastrado.
    temp = Right$(frmDrag.File1.FileName, 3)

    ' Si el archivo arrastrado se encuentra en la raíz, agrega el nombre del archivo.
    If Mid$(frmDrag.File1.Path, Len(frmDrag.File1.Path)) = "\" Then
      dropfile = frmDrag.File1.Path & frmDrag.File1.FileName
      
    ' Si el archivo arrastrado no se encuentra en la raíz, agrega "\" al nombre del archivo.
    Else
      dropfile = frmDrag.File1.Path & "\" & frmDrag.File1.FileName
      
    End If
    
    Guarda = DBCodigoEmpleado
    'InputBox("Digite el nombre")
    Origen = frmDrag.File1.Path & "\" & frmDrag.File1.FileName
       
       
    temp = Right$(frmDrag.File1.FileName, 3)
    Guarda = DBCodigoEmpleado.Text + "." + temp
    Destino = RutaFoto & Guarda
    
    Image1.Picture = LoadPicture("")
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
frmDrag.CmdSalir.Value = True
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub




Private Sub MacButton1_Click()

End Sub

Private Sub MaskEdAumento_Change()
Salida = True
End Sub

Private Sub MaskEdAumento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  MaskEdSuspencion.SetFocus
Else
   Evaluar = False
  End If
End Sub

Private Sub MaskEdBaja_Change()
Salida = True
End Sub

Private Sub MaskEdBaja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  TxtMotivoBaja.SetFocus
Else
   Evaluar = False
  End If
End Sub

Private Sub MaskEdBox4_LostFocus()
TxtSaldo.Text = MaskEdBox4.Text
End Sub

Private Sub MaskEdContrato_Change()

Salida = True
End Sub

Private Sub MaskEdContrato_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  DBCargoInicial.SetFocus
Else
   Evaluar = False
  End If
End Sub

Private Sub MaskEdFinalSusp_Change()
Salida = True
End Sub

Private Sub MaskEdFinalSusp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  TxtMotivoSuspencion.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub MaskEdNacimiento_Change()
Salida = True
End Sub

Private Sub MaskEdNacimiento_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  MaskEdContrato.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub MaskEdSuspencion_Change()
Salida = True
End Sub

Private Sub MaskEdSuspencion_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  MaskEdFinalSusp.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Select Case PreviousTab

Case 3
    DbgrLibreta.BorderStyle = 1
    DtaMovPrestamo.Refresh
    DbgrLibreta.Columns(0).Caption = "# Prestamo"
    DbgrLibreta.Columns(1).Caption = "Código Empleado"
    DbgrLibreta.Columns(2).Caption = "# de Cuota"
    DbgrLibreta.Columns(5).Caption = "Cuota Igualada"
    DbgrLibreta.Columns(6).Caption = "Saldo de Prestamo"
    DbgrLibreta.BorderStyle = 1

Case 4
    DbGIncentivos.BorderStyle = 1
    DtaDetalleIncentivo.Refresh
    DbGIncentivos.Columns(0).Caption = "# de Incentivo"
    DbGIncentivos.Columns(2).Caption = "Código Empleado"
    DbGIncentivos.Columns(4).Caption = "Pago Número"
    DbGIncentivos.BorderStyle = 1
Case 5
    DbgDeducciones.BorderStyle = 1
    DtaDetalleDeduccion.Refresh
    DbgDeducciones.Columns(0).Caption = "# de Deducción"
    DbgDeducciones.Columns(2).Caption = "Código Empleado"
    DbgDeducciones.Columns(4).Caption = "Pago Número"
    DbgDeducciones.BorderStyle = 1
Case 6
    DbgrSubsidios.BorderStyle = 1
    DtaDetalleSubsidio.Refresh
    DbgrSubsidios.Columns(0).Caption = "# de Subsidio"
    DbgrSubsidios.Columns(2).Visible = False
    DbgrSubsidios.Columns(4).Caption = "Valor Pago"
    DbgrSubsidios.Columns(6).Visible = False
    DbgrSubsidios.BorderStyle = 1
End Select

'DbgrLibreta.Columns(0).Visible = False
'DbgrLibreta.Columns(1).Visible = False
'DbgrLibreta.Columns(7).Visible = False
''DbGIncentivos.Columns(0).Visible = False

If PreviousTab > 2 Then
DbGIncentivos.Columns(2).Visible = False
DbGIncentivos.Columns(5).Visible = False
DbgDeducciones.Columns(0).Visible = False
DbgDeducciones.Columns(2).Visible = False
'DbgDeducciones.Columns(5).Visible = False
DbgrSubsidios.Columns(0).Visible = False
DbgrSubsidios.Columns(1).Visible = False
DbgrSubsidios.Columns(2).Visible = False
DbgrSubsidios.Columns(7).Visible = False
DbgrSubsidios.Columns(5).Width = 1200
DbgrSubsidios.Columns(6).Width = 500
End If


End Sub

Private Sub Text1_Change()
Me.DBCodigoEmpleado.Text = Me.Text1.Text
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

' End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)



'On Error GoTo TipoErrs
Dim SQlIncentivos As String, SQlDeducciones As String, SqlDetallePrestamo As String, SQlPrestamo As String, SqlDetalleSubsidio As String
Dim Salario As Boolean
Dim CodEmpleado1 As String
Dim numeroPrestamo As Double





If KeyCode = 13 Then

 Evaluar = True
 'Al ejecutar algun cambio en el combo actualizo el nombre del Empleado
 frmEmpleado.MousePointer = 11

 LimpiaEmpleado
 LimpiaHistorico
 LimpiaInfNomina
 'LimpiaInfNomina
' DtaEmpleados.Refresh
 ChkSuspendido.Visible = False
'Busco el codigo del empleado para que automaticamente ubique el nombre
 'aunque no existe en la data consulta
 CodEmpleado = -1
 
Me.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo,Dolarizado,CuentaBanco,SueldoActualBasico,HorasTurno,SalPorcentaje,CantPts,DiasBasico, AumentoBasico From Empleado WHERE (CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') And (Activo = 1)"
Me.DtaEmpleado.Refresh

If Not Me.DtaEmpleado.Recordset.EOF Then
'Do While Not DtaEmpleado.Recordset.EOF
'     If DtaEmpleado.Recordset("CodEmpleado1") = DBCodigoEmpleado.Text Then
'        If DtaEmpleado.Recordset("activo") = False Then
'           MsgBox "Este empleado ya fue dado de Baja"
'        End If
        
     CodEmpleado = Me.DtaEmpleado.Recordset("CodEmpleado")
     TxtCodEmpleado.Text = Me.DtaEmpleado.Recordset("CodEmpleado")
     
        If Not IsNull(DtaEmpleado.Recordset("numeroruc")) Then
          TxtNRuc.Text = DtaEmpleado.Recordset("numeroruc")
        End If
        
        
        If Not IsNull(DtaEmpleado.Recordset("AumentoBasico")) Then
          txtAumentoBasico.Text = DtaEmpleado.Recordset("AumentoBasico")
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("DiasBasico")) Then
          Me.TxtDiasBasico.Text = DtaEmpleado.Recordset("DiasBasico")
        End If
        'busco el tipo del archivo
        'Destino = ""
        If Dir(RutaFoto & DBCodigoEmpleado.Text & ".jpg") <> "" Then
           Destino = RutaFoto & DBCodigoEmpleado.Text & ".jpg"
        ElseIf Dir(RutaFoto & DBCodigoEmpleado.Text & ".gif") <> "" Then
           Destino = RutaFoto & DBCodigoEmpleado.Text & ".gif"
        ElseIf Dir(RutaFoto & DBCodigoEmpleado.Text & ".bmp") <> "" Then
           Destino = RutaFoto & DBCodigoEmpleado.Text & ".bmp"
        End If
        
        If (Dir(Destino) <> "") Then
         Image1.Picture = LoadPicture(Destino)
        Else
          Destino = App.Path + "\Zw.bmp"
'          Destino = RutaLogo
          If Dir(Destino) <> "" Then
            Image1.Picture = LoadPicture(Destino)
          End If
        End If
        
        If DtaEmpleado.Recordset("PorcientoIncentivo") = 0 Then
         Me.Check1.Value = 0
         Me.TxtPorcientoHora.Text = 0
         Me.TxtPorcientoHora.Visible = False
        Else
         Me.Check1.Value = 1
         Me.TxtPorcientoHora.Text = DtaEmpleado.Recordset("PorcientoIncentivo")
         Me.TxtPorcientoHora.Visible = True
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("SueldoActualBasico")) = True Then
         If DtaEmpleado.Recordset("SueldoActualBasico") = True Then
          Me.ChkSueldoActual.Value = 1
        Else
          Me.ChkSueldoActual.Value = 0
         End If
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("HorasTurno")) = True Then
            If DtaEmpleado.Recordset("HorasTurno") = True Then
             Me.ChkHorasTurno.Value = 1
            Else
             Me.ChkHorasTurno.Value = 0
            End If
        Else
          Me.ChkHorasTurno.Value = 0
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("CantPts")) Then
         Me.TxtDiasAdicionales.Text = DtaEmpleado.Recordset("CantPts")
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("CuentaBanco")) Then
        Me.TxtCuentaBanco.Text = DtaEmpleado.Recordset("CuentaBanco")
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("numcedula")) Then
        TxtNumCedula.Text = DtaEmpleado.Recordset("numcedula")
        End If
        ChkSuspendido.Visible = True
        TxtNombre1.Text = DtaEmpleado.Recordset("Nombre1")
        
        If Not IsNull(DtaEmpleado.Recordset("Nombre2")) Then
        TxtNombre2.Text = DtaEmpleado.Recordset("Nombre2")
        End If
        If Not IsNull(DtaEmpleado.Recordset("Apellido1")) Then
          TxtApellido1.Text = DtaEmpleado.Recordset("Apellido1")
        End If
        If Not IsNull(DtaEmpleado.Recordset("Apellido2")) Then
         TxtApellido2.Text = DtaEmpleado.Recordset("Apellido2")
        End If
        If Not IsNull(DtaEmpleado.Recordset("Direccion")) Then
           TxtDireccion.Text = DtaEmpleado.Recordset("Direccion")
        End If
        If Not IsNull(DtaEmpleado.Recordset("Nacionalidad")) Then
         TxtNacionalidad.Text = DtaEmpleado.Recordset("Nacionalidad")
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("SalPorcentaje")) Then
          Me.TxtSalarioPorciento.Text = DtaEmpleado.Recordset("SalPorcentaje")
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("Codgrupo")) Then
            TxtCodGrupo = DtaEmpleado.Recordset("Codgrupo")
        Else
            TxtCodGrupo = ""
            DBCGrupo.Text = ""
        End If
        If Not IsNull(DtaEmpleado.Recordset("CodigoPostal")) Then
          TxtCodPostal.Text = DtaEmpleado.Recordset("CodigoPostal")
        End If
        If Not IsNull(DtaEmpleado.Recordset("sexo")) Then
          CmbSexo.Text = DtaEmpleado.Recordset("sexo")
        End If
        If Not IsNull(DtaEmpleado.Recordset("NumeroInss")) Then
        TxtNInss.Text = DtaEmpleado.Recordset("NumeroInss")
        End If
        If Not IsNull(DtaEmpleado.Recordset("CodDepartamento")) Then
        TxtCodDepartamento.Text = DtaEmpleado.Recordset("CodDepartamento")
        End If
        If Not IsNull(DtaEmpleado.Recordset("CodCargo")) Then
          TxtCodCargo.Text = DtaEmpleado.Recordset("CodCargo")
        End If
        If Not IsNull(DtaEmpleado.Recordset("Sindicalista")) Then
          CmbSindicalista.Text = DtaEmpleado.Recordset("Sindicalista")
        End If
        If Not IsNull(DtaEmpleado.Recordset("numhijos")) Then
          TxtNumHijos.Text = DtaEmpleado.Recordset("numhijos")
        End If
        frmEmpleado.Caption = "Registro del Empleado: " & DBCodigoEmpleado.Text & ": " & TxtNombre1.Text & " " & TxtNombre2.Text & " " & TxtApellido1.Text & " " & TxtApellido2.Text
'        Me.CmdAcercade.Caption = DBCodigoEmpleado.Text & ":   " & txtNombre1.Text & " " & txtNombre2.Text & " " & txtApellido1.Text & " " & txtApellido2.Text
'        Me.xp_canvas1.Caption = "Registro del Empleado: " & DBCodigoEmpleado.Text & ": " & TxtNombre1.Text & " " & txtNombre2.Text & " " & TxtApellido1.Text & " " & txtApellido2.Text
        
        If Not IsNull(DtaEmpleado.Recordset("DiasDescuento")) Then
            TxtDiasDescuento.Text = DtaEmpleado.Recordset("DiasDescuento")
        Else
            TxtDiasDescuento.Text = 0
        End If
        
        Bandera = False
        
        If DtaEmpleado.Recordset("SalarioFijo") = "S" Then
          Salario = True
        Else
          Salario = False
        End If
        
        If DtaEmpleado.Recordset("ausente") = True Then
           ChkSuspendido.Value = 1
           LblSuspendido.Visible = True
        Else
           LblSuspendido.Visible = False
           ChkSuspendido.Value = 0
        End If
        
        If DtaEmpleado.Recordset("salariominimo") = True Then
            CmbSalarioMinimo.Text = "Verdaderp"
        Else
           CmbSalarioMinimo.Text = "Falso"
        End If
        
        If DtaEmpleado.Recordset("ExentoInss") = True Then
            CmbExentoInss.Text = "Verdadero"
        Else
           CmbExentoInss.Text = "Falso"
        End If
           
        If DtaEmpleado.Recordset("ExentoIr") = True Then
            CmbExentoIr.Text = "Verdadero"
        Else
           CmbExentoIr.Text = "Falso"
        End If
        
        If DtaEmpleado.Recordset("PagoInssPatronal") = True Then
            CmbPagoInssPatronal.Text = "Verdadero"
        Else
           CmbPagoInssPatronal.Text = "Falso"
        End If
        
        
        If DtaEmpleado.Recordset("Dolarizado") = True Then
           Me.ChkDolarizado.Value = xtpChecked
        Else
           Me.ChkDolarizado.Value = xtpUnchecked
        End If
        Bandera = True
    
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    SSTab1.TabEnabled(4) = True
    SSTab1.TabEnabled(5) = True
    SSTab1.TabEnabled(6) = True

    ' datos de la Nómina

'no olvidar los valores nomina

        DtaTipoNomina.Refresh
        Do While Not DtaTipoNomina.Recordset.EOF
           If DtaTipoNomina.Recordset("CodTipoNomina") = DtaEmpleado.Recordset("CodTipoNomina") Then
              DBCTipoNomina.Text = DtaTipoNomina.Recordset("nomina")
              Exit Do
            End If
        DtaTipoNomina.Recordset.MoveNext
        Loop
        
            If Not IsNull(DtaEmpleado.Recordset("SueldoPeriodo")) Then
            TxtSueldoPeriodo.Text = DtaEmpleado.Recordset("SueldoPeriodo")
            End If
            
        If Not IsNull(DtaEmpleado.Recordset("TarifaHoraria")) Then
            TxtTarifaHoraria.Text = DtaEmpleado.Recordset("TarifaHoraria")
        End If
        
        If Not IsNull(DtaEmpleado.Recordset("PorcentajeComision")) Then
            TxtComision.Text = DtaEmpleado.Recordset("PorcentajeComision")
        End If
      
     
       If Not IsNull(DtaEmpleado.Recordset("OtrosIngresos")) Then
          TxtOtrosIngresos.Text = DtaEmpleado.Recordset("OtrosIngresos")
       End If
       
       If Not IsNull(DtaEmpleado.Recordset("DescripOtrIngre")) Then
          TxtDescripOtrIngre.Text = DtaEmpleado.Recordset("DescripOtrIngre")
       End If
'        Exit Do
        Else 'si no lo encuentra
  
              SSTab1.TabEnabled(0) = True
              SSTab1.TabEnabled(1) = True
              SSTab1.TabEnabled(2) = True
              SSTab1.TabEnabled(3) = False
              SSTab1.TabEnabled(4) = False
              SSTab1.TabEnabled(5) = False
              SSTab1.TabEnabled(6) = False
        'frmEmpleado.Caption = "Registro del Empleado: " & TxtNombre1.Text & " " & TxtNombre2.Text & " " & TxtApellido1.Text & " " & TxtApellido2.Text
     End If
'DtaEmpleado.Recordset.MoveNext
'Loop

Evaluar = True
Me.DtaHistorico.RecordSource = "SELECT  Codempleado, FechaBaja, MotivoBaja, FechaAumento, MotivoAumento, FechaInicialSusp, FechaFinalSusp, MotivoSuspencion, FechaNacimiento, FechaContrato,FechaContratoVac , CargoInicial, CargoActual, CargoAnterior, SueldoInicial, SueldoAnterior, SueldoActual, CuentaDebito, CuentaCredito,CuentaPrestamo,CuentaOtrosIngresos,CuentaINSS,CuentaIR From Historico Where (CodEmpleado = " & CodEmpleado & " )"
DtaHistorico.Refresh
Do While Not DtaHistorico.Recordset.EOF
     If DtaHistorico.Recordset("CodEmpleado") = CodEmpleado Then
        If Not IsNull(DtaHistorico.Recordset("FechaNacimiento")) Then
            MaskEdNacimiento.Value = Format(DtaHistorico.Recordset("FechaNacimiento"), "dd/mm/yyyy")
        End If
        
        If Not IsNull(DtaHistorico.Recordset("FechaContratoVac")) Then
            Me.DTPFechaContratoVaca.Value = Format(DtaHistorico.Recordset("FechaContratoVac"), "dd/mm/yyyy")
        End If
        
        If Not IsNull(DtaHistorico.Recordset("FechaContrato")) Then
            MaskEdContrato.Value = Format(DtaHistorico.Recordset("FechaContrato"), "dd/mm/yyyy")
        End If
        
        If Not IsNull(DtaHistorico.Recordset("CargoInicial")) Then
          DBCargoInicial.Text = DtaHistorico.Recordset("CargoInicial")
        End If
        If Not IsNull(DtaHistorico.Recordset("CargoAnterior")) Then
           DBCargoAnterior.Text = DtaHistorico.Recordset("CargoAnterior")
        End If
        If Not IsNull(DtaHistorico.Recordset("CargoActual")) Then
             DBCargoActual.Text = DtaHistorico.Recordset("CargoActual")
        End If
        If Not IsNull(DtaHistorico.Recordset("MOTIVOBAJA")) Then
              TxtMotivoBaja.Text = DtaHistorico.Recordset("MOTIVOBAJA")
        End If
        If Not IsNull(DtaHistorico.Recordset("MotivoAumento")) Then
             TxtMotivoAumento.Text = DtaHistorico.Recordset("MotivoAumento")
        End If
        If Not IsNull(DtaHistorico.Recordset("MotivoSuspencion")) Then
             TxtMotivoSuspencion.Text = DtaHistorico.Recordset("MotivoSuspencion")
        End If
        
        TxtSueldoInicial.Text = Format((DtaHistorico.Recordset("SueldoInicial")), "##,##0.00")
        TxtSueldoAnterior.Text = Format((DtaHistorico.Recordset("SueldoAnterior")), "##,##0.00")
        TxtSueldoActual.Text = Format((DtaHistorico.Recordset("SueldoActual")), "##,##0.00")
        
        If Not IsNull(DtaHistorico.Recordset("fechabaja")) Then
             MaskEdBaja.Text = DtaHistorico.Recordset("fechabaja")
        End If
        
        If Not IsNull(DtaHistorico.Recordset("FechaAumento")) Then
             MaskEdAumento.Text = Format(DtaHistorico.Recordset("FechaAumento"), "dd/mm/yyyy")
        End If
         
         If Not IsNull(DtaHistorico.Recordset("FechaInicialSusp")) Then
            MaskEdSuspencion.Text = DtaHistorico.Recordset("FechaInicialSusp")
        End If
        
        If Not IsNull(DtaHistorico.Recordset("FechaInicialSusp")) Then
           MaskEdFinalSusp.Text = DtaHistorico.Recordset("FechaInicialSusp")
        End If
        
'        If Not IsNull(DtaHistorico.Recordset("CuentaDebito")) Then
'         TxtDebito.Text = DtaHistorico.Recordset("CuentaDebito")
'        End If
'
'        If Not IsNull(DtaHistorico.Recordset("cuentacredito")) Then
'         TxtCredito.Text = DtaHistorico.Recordset("cuentacredito")
'        End If
        
'        If Not IsNull(DtaHistorico.Recordset("CuentaPrestamo")) Then
'         Me.TxtCtaPrestamo.Text = DtaHistorico.Recordset("CuentaPrestamo")
'        End If
'
'        If Not IsNull(DtaHistorico.Recordset("CuentaOtrosIngresos")) Then
'         Me.TxtCtaOtrosIngresos.Text = DtaHistorico.Recordset("CuentaOtrosIngresos")
'        End If
'
'        If Not IsNull(DtaHistorico.Recordset("CuentaINSS")) Then
'         Me.TxtCuentaInss.Text = DtaHistorico.Recordset("CuentaINSS")
'        End If
'
'        If Not IsNull(DtaHistorico.Recordset("CuentaIR")) Then
'         Me.TxtCuentaIR.Text = DtaHistorico.Recordset("CuentaIR")
'        End If
        
        Exit Do
   End If
       DtaHistorico.Recordset.MoveNext
   Loop

DtaInfNomina.Refresh
Do While Not DtaInfNomina.Recordset.EOF
     If DtaInfNomina.Recordset("CodEmpleado") = CodEmpleado Then
        If Not IsNull(DtaInfNomina.Recordset("salariominimo")) Then
             CmbSalarioMinimo.Text = DtaInfNomina.Recordset("salariominimo")
        End If
        If Not IsNull(DtaInfNomina.Recordset("ExentoInss")) Then
             CmbExentoInss.Text = DtaInfNomina.Recordset("ExentoInss")
        End If
        If Not IsNull(DtaInfNomina.Recordset("ExentoIr")) Then
             CmbExentoIr.Text = DtaInfNomina.Recordset("ExentoIr")
        End If
       
        If Not IsNull(DtaInfNomina.Recordset("PagoInssPatronal")) Then
            CmbPagoInssPatronal.Text = DtaInfNomina.Recordset("PagoInssPatronal")
        End If
       
        TxtCodTipoNomina.Text = DtaInfNomina.Recordset("CodTipoNomina")
          
        Evaluar = True
        Exit Do
 
     End If
       DtaInfNomina.Recordset.MoveNext
   Loop

  DtaDepartamento.Refresh
   Do While Not DtaDepartamento.Recordset.EOF
     If DtaDepartamento.Recordset("CodDepartamento") = TxtCodDepartamento.Text Then
        DBCDepartamento.Text = DtaDepartamento.Recordset("departamento")
        Exit Do
     Else
      'DBCDepartamento.Text = ""
     End If
       DtaDepartamento.Recordset.MoveNext
   Loop

   DtaCargo.Refresh
   Do While Not DtaCargo.Recordset.EOF
     If DtaCargo.Recordset("CodCargo") = TxtCodCargo.Text Then
        DBCCargo.Text = DtaCargo.Recordset("Cargo")
        Exit Do
     Else
        DBCCargo.Text = ""
     End If
       DtaCargo.Recordset.MoveNext
   Loop
   
  Evaluar = True
 CmbIncapacidad.Text = "No"
  DtaIncapacidades.Refresh
   Do While Not DtaIncapacidades.Recordset.EOF
     If DtaIncapacidades.Recordset("CodEmpleado") = CodEmpleado Then
        CmbIncapacidad.Text = "Si"
        Exit Do
     Else
        CmbIncapacidad.Text = "No"
     End If
       DtaIncapacidades.Recordset.MoveNext
   Loop
Evaluar = True
   Salida = False
End If
   



'realizo los sql's de los incentivos y de las deducciones

CodEmpleado1 = DBCodigoEmpleado.Text
SQlIncentivos = "SELECT MAX(Incentivo.NumIncentivo) AS NumIncentivo, TipoIncentivo.Incentivo, Incentivo.CodEmpleado, AVG(DetalleIncentivo.Valor) AS Valor, COUNT(DetalleIncentivo.NumVez) AS NumVez, DetalleIncentivo.Pagado FROM TipoIncentivo INNER JOIN Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo ON TipoIncentivo.CodTipoIncentivo = Incentivo.CodTipoIncentivo GROUP BY TipoIncentivo.Incentivo, Incentivo.CodEmpleado, DetalleIncentivo.Pagado Having (Incentivo.CodEmpleado = " & CodEmpleado & ") And (DetalleIncentivo.Pagado = 0) "
'SQlIncentivos = "SELECT Incentivo.NumIncentivo, TipoIncentivo.Incentivo, Incentivo.CodEmpleado, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado FROM TipoIncentivo INNER JOIN (Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo) ON TipoIncentivo.CodTipoIncentivo = Incentivo.CodTipoIncentivo WHERE (Incentivo.CodEmpleado= " & CodEmpleado & ") And (DetalleIncentivo.Pagado = " & 0 & ")"
DtaDetalleIncentivo.RecordSource = SQlIncentivos
DtaDetalleIncentivo.Refresh

DbGIncentivos.Columns(0).Visible = False
DbGIncentivos.Columns(2).Visible = False
DbGIncentivos.Columns(5).Visible = False

'SQlDeducciones = "SELECT Deduccion.NumDeduccion, TipoDeduccion.Deduccion, Deduccion.CodEmpleado, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado FROM TipoDeduccion INNER JOIN (Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion) ON (TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion) WHERE Deduccion.CodEmpleado=" & CodEmpleado & " AND DetalleDeduccion.Pagado= " & 0 & " "
SQlDeducciones = "SELECT  MAX(Deduccion.NumDeduccion) AS NumDeduccion, TipoDeduccion.Deduccion, Deduccion.CodEmpleado, AVG(DetalleDeduccion.Valor) AS Valor, COUNT(DetalleDeduccion.NumVez) As NumVez FROM TipoDeduccion INNER JOIN Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion ON TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion Where (DetalleDeduccion.Pagado = 0) GROUP BY TipoDeduccion.Deduccion, Deduccion.CodEmpleado Having (Deduccion.CodEmpleado = " & CodEmpleado & ") ORDER BY NumDeduccion"
DtaDetalleDeduccion.RecordSource = SQlDeducciones
DtaDetalleDeduccion.Refresh

DbgDeducciones.Columns(0).Visible = False
DbgDeducciones.Columns(2).Visible = False
'DbgDeducciones.Columns(5).Visible = False
SQlPrestamo = "SELECT NumPrestamo, CuentaDebito, CuentaCredito, Monto, CantCuotas, Interes, Saldo, FechaInicial, Cancelado, Moneda, CuotasIguales, CodEmpleado From Prestamo WHERE Prestamo.CodEmpleado=" & CodEmpleado & " AND Prestamo.Cancelado=0"
DtaPrestamo.RecordSource = SQlPrestamo
DtaPrestamo.Refresh
If Not DtaPrestamo.Recordset.EOF Then
numeroPrestamo = Me.DtaPrestamo.Recordset("NumPrestamo")
Else
 numeroPrestamo = -100
End If
SqlDetallePrestamo = "SELECT MovPrestamo.ID,MovPrestamo.NumPrestamo, MovPrestamo.NumCuota, MovPrestamo.Monto, MovPrestamo.Interes, MovPrestamo.CuotaIgual,MovPrestamo.SaldoCuota , MovPrestamo.Cancelado FROM Prestamo INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo Where (MovPrestamo.Cancelado = 0) And (MovPrestamo.NumPrestamo = " & numeroPrestamo & ")"
DtaMovPrestamo.RecordSource = SqlDetallePrestamo
DtaMovPrestamo.Refresh






If Not DtaPrestamo.Recordset.EOF Then
   TxtCreditoPrestamo.Text = DtaPrestamo.Recordset("cuentacredito")
   TxtDebitoPrestamo.Text = DtaPrestamo.Recordset("CuentaDebito")
   Me.DbgrLibreta.Columns(0).Visible = False
DbgrLibreta.Columns(1).Visible = False
DbgrLibreta.Columns(7).Visible = False

Else
   TxtCreditoPrestamo.Text = " "
   TxtDebitoPrestamo.Text = " "
End If




SqlDetalleSubsidio = "SELECT Subsidio.NumSubsidio, Subsidio.CodEmpleado, Subsidio.CodTipoSubsidio, TipoSubsidio.Subsidio,DetalleSubsidio.Descripcion, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.Pagado FROM TipoSubsidio INNER JOIN (Subsidio INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio) ON TipoSubsidio.CodTipoSubsidio = Subsidio.CodTipoSubsidio WHERE DetalleSubsidio.Pagado=0 And Subsidio.CodEmpleado=" & CodEmpleado & " "
DtaDetalleSubsidio.RecordSource = SqlDetalleSubsidio
DtaDetalleSubsidio.Refresh

DbgrSubsidios.Columns(0).Visible = False
DbgrSubsidios.Columns(1).Visible = False
DbgrSubsidios.Columns(2).Visible = False
DbgrSubsidios.Columns(7).Visible = False
DbgrSubsidios.Columns(5).Width = 1200
DbgrSubsidios.Columns(6).Width = 500



If TxtNombre1.Text = "" Then
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    SSTab1.TabEnabled(4) = True
    SSTab1.TabEnabled(5) = True
    SSTab1.TabEnabled(6) = True
End If

        If Salario = True Then
         Me.ChkSalarioFijo.Value = 1
         Me.TxtComision.Enabled = False
        Else
         Me.ChkSalarioFijo.Value = 0
         Me.TxtComision.Enabled = True
        End If


Me.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo From Empleado WHERE     (CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') "
Me.DtaEmpleado.Refresh

If Not Me.DtaEmpleado.Recordset.EOF Then
 DBCodigoEmpleado.Text = Me.DtaEmpleado.Recordset("CodEmpleado1")
End If

'Me.DBCodigoEmpleado.Columns(0).Visible = False
'Me.DBCodigoEmpleado.Columns(1).Caption = "Codigo"
'Me.DBCodigoEmpleado.Columns(1).Width = 800
'Me.DBCodigoEmpleado.Columns(2).Visible = False


frmEmpleado.MousePointer = 0
frmEmpleado.AutoRedraw = True
Exit Sub
TipoErrs:
 MsgBox Err.Description
' ControlErrores
 Unload Me


End Sub

Private Sub TxtApellido1_Change()
Salida = True
End Sub

Private Sub TxtApellido1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  TxtApellido2.SetFocus
Else
   Evaluar = False
  End If
End Sub





Private Sub TxtApellido2_Change()
Salida = True
End Sub

Private Sub TxtApellido2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  TxtDireccion.SetFocus
 Else
   Evaluar = False
  End If
End Sub




Private Sub TxtAumento_Change()
Salida = True
End Sub

Private Sub TxtAumento_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 
 Else
   Evaluar = False
  End If
End Sub

Private Sub TxtCodGrupo_Change()
DtaGrupo.Refresh
Do While Not DtaGrupo.Recordset.EOF
   If TxtCodGrupo.Text = DtaGrupo.Recordset("Codgrupo") Then
      DBCGrupo.Text = DtaGrupo.Recordset("grupo")
      Exit Sub
   End If
DtaGrupo.Recordset.MoveNext
Loop
End Sub

Private Sub TxtCodigoEmpleados_Change()
'On Error GoTo TipoErrs
Dim SQlIncentivos As String, SQlDeducciones As String, SqlDetallePrestamo As String, SQlPrestamo As String, SqlDetalleSubsidio As String
Dim Salario As Boolean
Dim CodEmpleado1 As String, CodEmpleado As Double
Dim numeroPrestamo As Double

RegistrarBitacora = False




    
     Evaluar = True
     'Al ejecutar algun cambio en el combo actualizo el nombre del Empleado
     frmEmpleado.MousePointer = 11
    

     'LimpiaInfNomina
    ' DtaEmpleados.Refresh
     ChkSuspendido.Visible = False
    'Busco el codigo del empleado para que automaticamente ubique el nombre
     'aunque no existe en la data consulta
     CodEmpleado = -1
     
    Me.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo,Dolarizado,CuentaBanco,SueldoActualBasico,HorasTurno,SalPorcentaje,CantPts,DiasBasico,AumentoBasico, ViaticoxDia, DeducirPorciento, Reembolso, Telefono From Empleado WHERE (CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') And (Activo = 1)"
    Me.DtaEmpleado.Refresh

  If Not Me.DtaEmpleado.Recordset.EOF Then
'
'                  LimpiaEmpleado
'                  LimpiaHistorico
'                  LimpiaInfNomina
        
             CodEmpleado = Me.DtaEmpleado.Recordset("CodEmpleado")
             TxtCodEmpleado.Text = Me.DtaEmpleado.Recordset("CodEmpleado")
             
                If Not IsNull(DtaEmpleado.Recordset("Telefono")) Then
                  Me.TxtTelefono.Text = DtaEmpleado.Recordset("Telefono")
                End If
             
                If Not IsNull(DtaEmpleado.Recordset("Reembolso")) Then
                  Me.TxtReembolso.Text = DtaEmpleado.Recordset("Reembolso")
                End If
             
                If Not IsNull(DtaEmpleado.Recordset("DeducirPorciento")) Then
                  If DtaEmpleado.Recordset("DeducirPorciento") = True Then
                     Me.ChkDeducirPorcentaje.Value = 1
                  Else
                    Me.ChkDeducirPorcentaje.Value = 0
                  End If
                Else
                  Me.ChkDeducirPorcentaje.Value = 1
                End If
             
             
                If Not IsNull(DtaEmpleado.Recordset("ViaticoxDia")) Then
                  Me.TxtViatico.Text = DtaEmpleado.Recordset("ViaticoxDia")
                Else
                  Me.TxtViatico.Text = 0
                End If
                
                If Not IsNull(DtaEmpleado.Recordset("numeroruc")) Then
                  TxtNRuc.Text = DtaEmpleado.Recordset("numeroruc")
                End If
                
                If Not IsNull(DtaEmpleado.Recordset("AumentoBasico")) Then
                  txtAumentoBasico.Text = DtaEmpleado.Recordset("AumentoBasico")
                End If
                
                If Not IsNull(DtaEmpleado.Recordset("DiasBasico")) Then
                  Me.TxtDiasBasico.Text = DtaEmpleado.Recordset("DiasBasico")
                End If
                'busco el tipo del archivo
                'Destino = ""
                If Dir(RutaFoto & DBCodigoEmpleado.Text & ".jpg") <> "" Then
                   Destino = RutaFoto & DBCodigoEmpleado.Text & ".jpg"
                ElseIf Dir(RutaFoto & DBCodigoEmpleado.Text & ".gif") <> "" Then
                   Destino = RutaFoto & DBCodigoEmpleado.Text & ".gif"
                ElseIf Dir(RutaFoto & DBCodigoEmpleado.Text & ".bmp") <> "" Then
                   Destino = RutaFoto & DBCodigoEmpleado.Text & ".bmp"
                End If
                
                If (Dir(Destino) <> "") Then
                 Image1.Picture = LoadPicture(Destino)
                Else
                  Destino = App.Path + "\Zw.bmp"
        '          Destino = RutaLogo
                  If Dir(Destino) <> "" Then
                    Image1.Picture = LoadPicture(Destino)
                  End If
                End If
                
                If DtaEmpleado.Recordset("PorcientoIncentivo") = 0 Then
                 Me.Check1.Value = 0
                 Me.TxtPorcientoHora.Text = 0
                 Me.TxtPorcientoHora.Visible = False
                Else
                 Me.Check1.Value = 1
                 Me.TxtPorcientoHora.Text = DtaEmpleado.Recordset("PorcientoIncentivo")
                 Me.TxtPorcientoHora.Visible = True
                End If
                
                If Not IsNull(DtaEmpleado.Recordset("SueldoActualBasico")) = True Then
                 If DtaEmpleado.Recordset("SueldoActualBasico") = True Then
                  Me.ChkSueldoActual.Value = 1
                Else
                  Me.ChkSueldoActual.Value = 0
                 End If
                End If
                
                If Not IsNull(DtaEmpleado.Recordset("HorasTurno")) = True Then
                    If DtaEmpleado.Recordset("HorasTurno") = True Then
                     Me.ChkHorasTurno.Value = 1
                    Else
                     Me.ChkHorasTurno.Value = 0
                    End If
                Else
                  Me.ChkHorasTurno.Value = 0
                End If
                
                If Not IsNull(DtaEmpleado.Recordset("CantPts")) Then
                 Me.TxtDiasAdicionales.Text = DtaEmpleado.Recordset("CantPts")
                End If
                
                If Not IsNull(DtaEmpleado.Recordset("CuentaBanco")) Then
                Me.TxtCuentaBanco.Text = DtaEmpleado.Recordset("CuentaBanco")
                End If
                
                If Not IsNull(DtaEmpleado.Recordset("numcedula")) Then
                TxtNumCedula.Text = DtaEmpleado.Recordset("numcedula")
                End If
                ChkSuspendido.Visible = True
                
                If Not IsNull(DtaEmpleado.Recordset("Nombre1")) Then
                  TxtNombre1.Text = DtaEmpleado.Recordset("Nombre1")
                End If
                
                If Not IsNull(DtaEmpleado.Recordset("Nombre2")) Then
                TxtNombre2.Text = DtaEmpleado.Recordset("Nombre2")
                End If
                If Not IsNull(DtaEmpleado.Recordset("Apellido1")) Then
                  TxtApellido1.Text = DtaEmpleado.Recordset("Apellido1")
                End If
                If Not IsNull(DtaEmpleado.Recordset("Apellido2")) Then
                 TxtApellido2.Text = DtaEmpleado.Recordset("Apellido2")
                End If
                If Not IsNull(DtaEmpleado.Recordset("Direccion")) Then
                   TxtDireccion.Text = DtaEmpleado.Recordset("Direccion")
                End If
                If Not IsNull(DtaEmpleado.Recordset("Nacionalidad")) Then
                 TxtNacionalidad.Text = DtaEmpleado.Recordset("Nacionalidad")
                End If
                
                If Not IsNull(DtaEmpleado.Recordset("SalPorcentaje")) Then
                  Me.TxtSalarioPorciento.Text = DtaEmpleado.Recordset("SalPorcentaje")
                End If
                
                If Not IsNull(DtaEmpleado.Recordset("Codgrupo")) Then
                    TxtCodGrupo = DtaEmpleado.Recordset("Codgrupo")
                Else
                    TxtCodGrupo = ""
                    DBCGrupo.Text = ""
                End If
                If Not IsNull(DtaEmpleado.Recordset("CodigoPostal")) Then
                  TxtCodPostal.Text = DtaEmpleado.Recordset("CodigoPostal")
                End If
                If Not IsNull(DtaEmpleado.Recordset("sexo")) Then
                  CmbSexo.Text = DtaEmpleado.Recordset("sexo")
                End If
                If Not IsNull(DtaEmpleado.Recordset("NumeroInss")) Then
                TxtNInss.Text = DtaEmpleado.Recordset("NumeroInss")
                End If
                If Not IsNull(DtaEmpleado.Recordset("CodDepartamento")) Then
                TxtCodDepartamento.Text = DtaEmpleado.Recordset("CodDepartamento")
                End If
                If Not IsNull(DtaEmpleado.Recordset("CodCargo")) Then
                  TxtCodCargo.Text = DtaEmpleado.Recordset("CodCargo")
                End If
                If Not IsNull(DtaEmpleado.Recordset("Sindicalista")) Then
                  CmbSindicalista.Text = DtaEmpleado.Recordset("Sindicalista")
                End If
                If Not IsNull(DtaEmpleado.Recordset("numhijos")) Then
                  TxtNumHijos.Text = DtaEmpleado.Recordset("numhijos")
                End If
                frmEmpleado.Caption = "Registro del Empleado: " & DBCodigoEmpleado.Text & ": " & TxtNombre1.Text & " " & TxtNombre2.Text & " " & TxtApellido1.Text & " " & TxtApellido2.Text
        '        Me.CmdAcercade.Caption = DBCodigoEmpleado.Text & ":   " & txtNombre1.Text & " " & txtNombre2.Text & " " & txtApellido1.Text & " " & txtApellido2.Text
        '        Me.xp_canvas1.Caption = "Registro del Empleado: " & DBCodigoEmpleado.Text & ": " & TxtNombre1.Text & " " & txtNombre2.Text & " " & TxtApellido1.Text & " " & txtApellido2.Text
                
                If Not IsNull(DtaEmpleado.Recordset("DiasDescuento")) Then
                    TxtDiasDescuento.Text = DtaEmpleado.Recordset("DiasDescuento")
                Else
                    TxtDiasDescuento.Text = 0
                End If
        
        
        
                Bandera = False
                
                If DtaEmpleado.Recordset("SalarioFijo") = "S" Then
                  Salario = True
                Else
                  Salario = False
                End If
                
                If DtaEmpleado.Recordset("ausente") = True Then
                   ChkSuspendido.Value = 1
                   LblSuspendido.Visible = True
                Else
                   LblSuspendido.Visible = False
                   ChkSuspendido.Value = 0
                End If
                
                If DtaEmpleado.Recordset("salariominimo") = True Then
                    CmbSalarioMinimo.Text = "Verdaderp"
                Else
                   CmbSalarioMinimo.Text = "Falso"
                End If
                
                If DtaEmpleado.Recordset("ExentoInss") = True Then
                    CmbExentoInss.Text = "Verdadero"
                Else
                   CmbExentoInss.Text = "Falso"
                End If
           
                If DtaEmpleado.Recordset("ExentoIr") = True Then
                    CmbExentoIr.Text = "Verdadero"
                Else
                   CmbExentoIr.Text = "Falso"
                End If
                
                If DtaEmpleado.Recordset("PagoInssPatronal") = True Then
                    CmbPagoInssPatronal.Text = "Verdadero"
                Else
                   CmbPagoInssPatronal.Text = "Falso"
                End If
                
                
                If DtaEmpleado.Recordset("Dolarizado") = True Then
                   Me.ChkDolarizado.Value = xtpChecked
                Else
                   Me.ChkDolarizado.Value = xtpUnchecked
                End If
                Bandera = True
    
                SSTab1.TabEnabled(0) = True
                SSTab1.TabEnabled(1) = True
                SSTab1.TabEnabled(2) = True
                SSTab1.TabEnabled(3) = True
                SSTab1.TabEnabled(4) = True
                SSTab1.TabEnabled(5) = True
                SSTab1.TabEnabled(6) = True
            
                ' datos de la Nómina
            
            'no olvidar los valores nomina

                DtaTipoNomina.Refresh
                Do While Not DtaTipoNomina.Recordset.EOF
                   If DtaTipoNomina.Recordset("CodTipoNomina") = DtaEmpleado.Recordset("CodTipoNomina") Then
                      DBCTipoNomina.Text = DtaTipoNomina.Recordset("nomina")
                      Exit Do
                    End If
                DtaTipoNomina.Recordset.MoveNext
                Loop
        
                If Not IsNull(DtaEmpleado.Recordset("SueldoPeriodo")) Then
                   TxtSueldoPeriodo.Text = DtaEmpleado.Recordset("SueldoPeriodo")
                End If
            
               If Not IsNull(DtaEmpleado.Recordset("TarifaHoraria")) Then
                   TxtTarifaHoraria.Text = DtaEmpleado.Recordset("TarifaHoraria")
               End If
               
               If Not IsNull(DtaEmpleado.Recordset("PorcentajeComision")) Then
                   TxtComision.Text = DtaEmpleado.Recordset("PorcentajeComision")
               End If
             
            
              If Not IsNull(DtaEmpleado.Recordset("OtrosIngresos")) Then
                 TxtOtrosIngresos.Text = DtaEmpleado.Recordset("OtrosIngresos")
              End If
       
               If Not IsNull(DtaEmpleado.Recordset("DescripOtrIngre")) Then
                  TxtDescripOtrIngre.Text = DtaEmpleado.Recordset("DescripOtrIngre")
               End If
               
               
                Evaluar = True
                Me.DtaHistorico.RecordSource = "SELECT  Codempleado, FechaBaja, MotivoBaja, FechaAumento, MotivoAumento, FechaInicialSusp, FechaFinalSusp, MotivoSuspencion, FechaNacimiento, FechaContrato,FechaContratoVac , CargoInicial, CargoActual, CargoAnterior, SueldoInicial, SueldoAnterior, SueldoActual, CuentaDebito, CuentaCredito,CuentaPrestamo,CuentaOtrosIngresos,CuentaINSS,CuentaIR From Historico Where (CodEmpleado = " & CodEmpleado & " )"
                DtaHistorico.Refresh
                Do While Not DtaHistorico.Recordset.EOF
                     If DtaHistorico.Recordset("CodEmpleado") = CodEmpleado Then
                        If Not IsNull(DtaHistorico.Recordset("FechaNacimiento")) Then
                            MaskEdNacimiento.Value = Format(DtaHistorico.Recordset("FechaNacimiento"), "dd/mm/yyyy")
                        End If
                        
                        If Not IsNull(DtaHistorico.Recordset("FechaContratoVac")) Then
                            Me.DTPFechaContratoVaca.Value = Format(DtaHistorico.Recordset("FechaContratoVac"), "dd/mm/yyyy")
                        End If
                        
                        If Not IsNull(DtaHistorico.Recordset("FechaContrato")) Then
                            MaskEdContrato.Value = Format(DtaHistorico.Recordset("FechaContrato"), "dd/mm/yyyy")
                        End If
                        
                        If Not IsNull(DtaHistorico.Recordset("CargoInicial")) Then
                          DBCargoInicial.Text = DtaHistorico.Recordset("CargoInicial")
                        End If
                        If Not IsNull(DtaHistorico.Recordset("CargoAnterior")) Then
                           DBCargoAnterior.Text = DtaHistorico.Recordset("CargoAnterior")
                        End If
                        If Not IsNull(DtaHistorico.Recordset("CargoActual")) Then
                             DBCargoActual.Text = DtaHistorico.Recordset("CargoActual")
                        End If
                        If Not IsNull(DtaHistorico.Recordset("MOTIVOBAJA")) Then
                              TxtMotivoBaja.Text = DtaHistorico.Recordset("MOTIVOBAJA")
                        End If
                        If Not IsNull(DtaHistorico.Recordset("MotivoAumento")) Then
                             TxtMotivoAumento.Text = DtaHistorico.Recordset("MotivoAumento")
                        End If
                        If Not IsNull(DtaHistorico.Recordset("MotivoSuspencion")) Then
                             TxtMotivoSuspencion.Text = DtaHistorico.Recordset("MotivoSuspencion")
                        End If
            
                        TxtSueldoInicial.Text = Format((DtaHistorico.Recordset("SueldoInicial")), "##,##0.00")
                        TxtSueldoAnterior.Text = Format((DtaHistorico.Recordset("SueldoAnterior")), "##,##0.00")
                        TxtSueldoActual.Text = Format((DtaHistorico.Recordset("SueldoActual")), "##,##0.00")
                        
                        If Not IsNull(DtaHistorico.Recordset("fechabaja")) Then
                             MaskEdBaja.Text = DtaHistorico.Recordset("fechabaja")
                        End If
                        
                        If Not IsNull(DtaHistorico.Recordset("FechaAumento")) Then
                             MaskEdAumento.Text = Format(DtaHistorico.Recordset("FechaAumento"), "dd/mm/yyyy")
                        End If
                         
                         If Not IsNull(DtaHistorico.Recordset("FechaInicialSusp")) Then
                            MaskEdSuspencion.Text = DtaHistorico.Recordset("FechaInicialSusp")
                        End If
                        
                        If Not IsNull(DtaHistorico.Recordset("FechaInicialSusp")) Then
                           MaskEdFinalSusp.Text = DtaHistorico.Recordset("FechaInicialSusp")
                        End If
    
            
                        Exit Do
                  End If
                  DtaHistorico.Recordset.MoveNext
                Loop

                DtaInfNomina.Refresh
                Do While Not DtaInfNomina.Recordset.EOF
                     If DtaInfNomina.Recordset("CodEmpleado") = CodEmpleado Then
                        If Not IsNull(DtaInfNomina.Recordset("salariominimo")) Then
                             CmbSalarioMinimo.Text = DtaInfNomina.Recordset("salariominimo")
                        End If
                        If Not IsNull(DtaInfNomina.Recordset("ExentoInss")) Then
                             CmbExentoInss.Text = DtaInfNomina.Recordset("ExentoInss")
                        End If
                        If Not IsNull(DtaInfNomina.Recordset("ExentoIr")) Then
                             CmbExentoIr.Text = DtaInfNomina.Recordset("ExentoIr")
                        End If
                       
                        If Not IsNull(DtaInfNomina.Recordset("PagoInssPatronal")) Then
                            CmbPagoInssPatronal.Text = DtaInfNomina.Recordset("PagoInssPatronal")
                        End If
                       
                        If Not IsNull(DtaInfNomina.Recordset("CodTipoNomina")) Then
                        TxtCodTipoNomina.Text = DtaInfNomina.Recordset("CodTipoNomina")
                        End If
                          
                        Evaluar = True
                        Exit Do
                 
                     End If
                       DtaInfNomina.Recordset.MoveNext
                   Loop
    
                    DtaDepartamento.Refresh
                     Do While Not DtaDepartamento.Recordset.EOF
                       If DtaDepartamento.Recordset("CodDepartamento") = TxtCodDepartamento.Text Then
                          DBCDepartamento.Text = DtaDepartamento.Recordset("departamento")
                          Exit Do
                       Else
                        'DBCDepartamento.Text = ""
                       End If
                         DtaDepartamento.Recordset.MoveNext
                     Loop
                     
                DtaCargo.Refresh
                Do While Not DtaCargo.Recordset.EOF
                  If DtaCargo.Recordset("CodCargo") = TxtCodCargo.Text Then
                     DBCCargo.Text = DtaCargo.Recordset("Cargo")
                     Exit Do
                  Else
                     DBCCargo.Text = ""
                  End If
                    DtaCargo.Recordset.MoveNext
                Loop
   
                Evaluar = True
                CmbIncapacidad.Text = "No"
                DtaIncapacidades.Refresh
  
                Do While Not DtaIncapacidades.Recordset.EOF
                  If DtaIncapacidades.Recordset("CodEmpleado") = CodEmpleado Then
                     CmbIncapacidad.Text = "Si"
                     Exit Do
                  Else
                     CmbIncapacidad.Text = "No"
                  End If
                    DtaIncapacidades.Recordset.MoveNext
                Loop
                
                Evaluar = True
                Salida = False
                
                'realizo los sql's de los incentivos y de las deducciones
                
                CodEmpleado1 = DBCodigoEmpleado.Text
                SQlIncentivos = "SELECT MAX(Incentivo.NumIncentivo) AS NumIncentivo, TipoIncentivo.Incentivo, Incentivo.CodEmpleado, AVG(DetalleIncentivo.Valor) AS Valor, COUNT(DetalleIncentivo.NumVez) AS NumVez, DetalleIncentivo.Pagado FROM TipoIncentivo INNER JOIN Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo ON TipoIncentivo.CodTipoIncentivo = Incentivo.CodTipoIncentivo GROUP BY TipoIncentivo.Incentivo, Incentivo.CodEmpleado, DetalleIncentivo.Pagado Having (Incentivo.CodEmpleado = " & CodEmpleado & ") And (DetalleIncentivo.Pagado = 0) "
                'SQlIncentivos = "SELECT Incentivo.NumIncentivo, TipoIncentivo.Incentivo, Incentivo.CodEmpleado, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado FROM TipoIncentivo INNER JOIN (Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo) ON TipoIncentivo.CodTipoIncentivo = Incentivo.CodTipoIncentivo WHERE (Incentivo.CodEmpleado= " & CodEmpleado & ") And (DetalleIncentivo.Pagado = " & 0 & ")"
                DtaDetalleIncentivo.RecordSource = SQlIncentivos
                DtaDetalleIncentivo.Refresh
                
                DbGIncentivos.Columns(0).Visible = False
                DbGIncentivos.Columns(2).Visible = False
                DbGIncentivos.Columns(5).Visible = False
                
                'SQlDeducciones = "SELECT Deduccion.NumDeduccion, TipoDeduccion.Deduccion, Deduccion.CodEmpleado, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado FROM TipoDeduccion INNER JOIN (Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion) ON (TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion) WHERE Deduccion.CodEmpleado=" & CodEmpleado & " AND DetalleDeduccion.Pagado= " & 0 & " "
                SQlDeducciones = "SELECT  MAX(Deduccion.NumDeduccion) AS NumDeduccion, TipoDeduccion.Deduccion, Deduccion.CodEmpleado, AVG(DetalleDeduccion.Valor) AS Valor, COUNT(DetalleDeduccion.NumVez) As NumVez FROM TipoDeduccion INNER JOIN Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion ON TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion Where (DetalleDeduccion.Pagado = 0) GROUP BY TipoDeduccion.Deduccion, Deduccion.CodEmpleado Having (Deduccion.CodEmpleado = " & CodEmpleado & ") ORDER BY NumDeduccion"
                DtaDetalleDeduccion.RecordSource = SQlDeducciones
                DtaDetalleDeduccion.Refresh
                
                DbgDeducciones.Columns(0).Visible = False
                DbgDeducciones.Columns(2).Visible = False
                'DbgDeducciones.Columns(5).Visible = False
                SQlPrestamo = "SELECT NumPrestamo, CuentaDebito, CuentaCredito, Monto, CantCuotas, Interes, Saldo, FechaInicial, Cancelado, Moneda, CuotasIguales, CodEmpleado From Prestamo WHERE Prestamo.CodEmpleado=" & CodEmpleado & " AND Prestamo.Cancelado=0"
                DtaPrestamo.RecordSource = SQlPrestamo
                DtaPrestamo.Refresh
                If Not DtaPrestamo.Recordset.EOF Then
                numeroPrestamo = Me.DtaPrestamo.Recordset("NumPrestamo")
                Else
                 numeroPrestamo = -100
                End If
                SqlDetallePrestamo = "SELECT MovPrestamo.ID,MovPrestamo.NumPrestamo, MovPrestamo.NumCuota, MovPrestamo.Monto, MovPrestamo.Interes, MovPrestamo.CuotaIgual,MovPrestamo.SaldoCuota , MovPrestamo.Cancelado FROM Prestamo INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo Where (MovPrestamo.Cancelado = 0) And (MovPrestamo.NumPrestamo = " & numeroPrestamo & ")"
                DtaMovPrestamo.RecordSource = SqlDetallePrestamo
                DtaMovPrestamo.Refresh
                
                                If Not DtaPrestamo.Recordset.EOF Then
                   TxtCreditoPrestamo.Text = DtaPrestamo.Recordset("cuentacredito")
                   TxtDebitoPrestamo.Text = DtaPrestamo.Recordset("CuentaDebito")
                   Me.DbgrLibreta.Columns(0).Visible = False
                DbgrLibreta.Columns(1).Visible = False
                DbgrLibreta.Columns(7).Visible = False
                
                Else
                   TxtCreditoPrestamo.Text = " "
                   TxtDebitoPrestamo.Text = " "
                End If
                
                
                
                
                SqlDetalleSubsidio = "SELECT Subsidio.NumSubsidio, Subsidio.CodEmpleado, Subsidio.CodTipoSubsidio, TipoSubsidio.Subsidio,DetalleSubsidio.Descripcion, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.Pagado FROM TipoSubsidio INNER JOIN (Subsidio INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio) ON TipoSubsidio.CodTipoSubsidio = Subsidio.CodTipoSubsidio WHERE DetalleSubsidio.Pagado=0 And Subsidio.CodEmpleado=" & CodEmpleado & " "
                DtaDetalleSubsidio.RecordSource = SqlDetalleSubsidio
                DtaDetalleSubsidio.Refresh
                
                DbgrSubsidios.Columns(0).Visible = False
                DbgrSubsidios.Columns(1).Visible = False
                DbgrSubsidios.Columns(2).Visible = False
                DbgrSubsidios.Columns(7).Visible = False
                DbgrSubsidios.Columns(5).Width = 1200
                DbgrSubsidios.Columns(6).Width = 500
                
                
                
                If TxtNombre1.Text = "" Then
                    SSTab1.TabEnabled(1) = True
                    SSTab1.TabEnabled(2) = True
                    SSTab1.TabEnabled(3) = True
                    SSTab1.TabEnabled(4) = True
                    SSTab1.TabEnabled(5) = True
                    SSTab1.TabEnabled(6) = True
                End If
                
                      If Salario = True Then
                         Me.ChkSalarioFijo.Value = 1
                         Me.TxtComision.Enabled = False
                        Else
                         Me.ChkSalarioFijo.Value = 0
                         Me.TxtComision.Enabled = True
                        End If
                
                
'                Me.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo From Empleado WHERE     (CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') "
'                Me.DtaEmpleado.Refresh
'
'                If Not Me.DtaEmpleado.Recordset.EOF Then
'                 DBCodigoEmpleado.Text = Me.DtaEmpleado.Recordset("CodEmpleado1")
'                End If
                
                'Me.DBCodigoEmpleado.Columns(0).Visible = False
                'Me.DBCodigoEmpleado.Columns(1).Caption = "Codigo"
                'Me.DBCodigoEmpleado.Columns(1).Width = 800
                'Me.DBCodigoEmpleado.Columns(2).Visible = False
                
                
                frmEmpleado.MousePointer = 0
                frmEmpleado.AutoRedraw = True

      
            Else 'si no lo encuentra
            
            
                LimpiaEmpleado
                LimpiaHistorico
                LimpiaInfNomina
          
                      SSTab1.TabEnabled(0) = True
                      SSTab1.TabEnabled(1) = True
                      SSTab1.TabEnabled(2) = True
                      SSTab1.TabEnabled(3) = False
                      SSTab1.TabEnabled(4) = False
                      SSTab1.TabEnabled(5) = False
                      SSTab1.TabEnabled(6) = False
                      
                      frmEmpleado.MousePointer = 0
                'frmEmpleado.Caption = "Registro del Empleado: " & TxtNombre1.Text & " " & TxtNombre2.Text & " " & TxtApellido1.Text & " " & TxtApellido2.Text
      End If
      
      
      RegistrarBitacora = True

Exit Sub
TipoErrs:
                 MsgBox Err.Description
                ' ControlErrores
                 Unload Me
                ' End If
                
                
   
End Sub

Private Sub TxtCodPostal_Change()
Salida = True
End Sub

Private Sub TxtCodPostal_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 CmbSexo.SetFocus
 Else
   Evaluar = False
  End If
End Sub


Private Sub TxtCodTipoNomina_Change()
On erro GoTo TipoErrs
 'Salida = True
 'PreparaSalida
 DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
     If DtaTipoNomina.Recordset("CodTipoNomina") = TxtCodTipoNomina.Text Then
        DBCTipoNomina.Text = DtaTipoNomina.Recordset("nomina")
        Exit Do
     Else
        'LimpiaEmpleado
     End If
       DtaTipoNomina.Recordset.MoveNext
   Loop
 Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub




Private Sub TxtComision_LostFocus()
If Not IsNumeric(TxtComision.Text) Then
    MsgBox "la Comisión es errónea"
    TxtComision.SetFocus
Else
    TxtComision.Text = Format((TxtComision.Text), "##,##0.00")
End If

End Sub

Private Sub TxtCredito_Change()
Salida = True
End Sub

Private Sub TxtCredito_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 Else
   Evaluar = False
  End If
End Sub

Private Sub TxtCuentaBanco_Change()
'Dim Nombres As String
'
' Nombres = Me.TxtNombre1.Text & " " & Me.TxtNombre2.Text & " " & Me.TxtApellido1.Text & " " & Me.TxtApellido2.Text
'  '////////////////////////VALIDO EL CAMBIO DE CUENTA DE BANCO /////////////////////////////////////
'  If Len(Me.TxtCuentaBanco.Text) >= 8 Then
'    Me.DtaConsulta.RecordSource = "SELECT CuentaBanco, CodEmpleado1, Nombre1 + ' ' + Nombre2 + ' ' + Apellido1 + ' ' + Apellido2 AS Nombres, NumCedula, NumeroInss From Empleado WHERE (CuentaBanco = '" & Me.TxtCuentaBanco.Text & "')"
'    Me.DtaConsulta.Refresh
'    Do While Not Me.DtaConsulta.Recordset.EOF
'
'        If Me.DtaConsulta.Recordset("NumCedula") <> Me.TxtNumCedula.Text Then
'            res = Bitacora(Now, NombreUsuario, "Empleados", "Se Cambio el Numero de Cuenta de Banco con Cedula Distinta: " & Me.DBCodigoEmpleado.Text & " " & Me.TxtNombre1.Text)
'
'        End If
'
'        If Me.DtaConsulta.Recordset("NumeroInss") <> Me.TxtNInss.Text Then
'           res = Bitacora(Now, NombreUsuario, "Empleados", "Se Cambio el Numero de Cuenta de Banco con Numero Inss Distinto: " & Me.DBCodigoEmpleado.Text & " " & Me.TxtNombre1.Text)
'        End If
'
'        If Me.DtaConsulta.Recordset("Nombres") <> Nombres Then
'           res = Bitacora(Now, NombreUsuario, "Empleados", "Se Cambio el Numero de Cuenta de Banco con Nombre Distinto: " & Me.DBCodigoEmpleado.Text & " " & Me.TxtNombre1.Text)
'        End If
'
'       Me.DtaConsulta.Recordset.MoveNext
'    Loop
'
'
'  End If
End Sub

Private Sub TxtCuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then
  Me.CmdAfectuar.Value = True
End If
End Sub

Private Sub TxtDebito_Change()
Salida = True
End Sub

Private Sub TxtDebito_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtCredito.SetFocus
 Else
   Evaluar = False
  End If
End Sub



Private Sub TxtDiasDescuento_LostFocus()


On Error GoTo TipoErr
Dim Periodo As String

If Not IsNumeric(TxtDiasDescuento.Text) Then
   MsgBox "Los dias de descuento son erróneos"
   TxtDiasDescuento.SetFocus
   Exit Sub
End If

DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
  If DtaTipoNomina.Recordset("nomina") = DBCTipoNomina.Text Then
     Periodo = DtaTipoNomina.Recordset("Periodo")
     Exit Do
  End If
DtaTipoNomina.Recordset.MoveNext
Loop

Select Case Periodo

Case "Semanal Viernes"
     If val(TxtDiasDescuento.Text) > 7 Then
        MsgBox "Los dias de descuento no pueden ser mayor que 7, porque esta es una nómina Semanal"
        TxtDiasDescuento.SetFocus
        Exit Sub
     End If
Case "Semanal Sabado"
    If val(TxtDiasDescuento.Text) > 7 Then
        MsgBox "Los dias de descuento no pueden ser mayor que 7, porque esta es una nómina Semanal"
        TxtDiasDescuento.SetFocus
        Exit Sub
     End If
Case "Catorcenal los Viernes"
    If val(TxtDiasDescuento.Text) > 14 Then
        MsgBox "Los dias de descuento no pueden ser mayor que 14, porque esta es una nómina Catorcenal"
        TxtDiasDescuento.SetFocus
        Exit Sub
     End If
Case "Catorcenal los Sabados"
    If val(TxtDiasDescuento.Text) > 14 Then
        MsgBox "Los dias de descuento no pueden ser mayor que 14, porque esta es una nómina Catorcenal"
        TxtDiasDescuento.SetFocus
        Exit Sub
     End If
Case "Quincenal"
    If val(TxtDiasDescuento.Text) > 15 Then
        MsgBox "Los dias de descuento no pueden ser mayor que 15, porque esta es una nómina Quincenal"
        TxtDiasDescuento.SetFocus
        Exit Sub
     End If
Case "Mensual"
    If val(TxtDiasDescuento.Text) > 30 Then
        MsgBox "Los dias de descuento no pueden ser mayor que 30, porque esta es una nómina Mensual"
        TxtDiasDescuento.SetFocus
        Exit Sub
     End If
Case "Trimestral"
    If val(TxtDiasDescuento.Text) > 90 Then
        MsgBox "Los dias de descuento no pueden ser mayor que 90, porque esta es una nómina Trimestral"
        TxtDiasDescuento.SetFocus
        Exit Sub
     End If
Case "Semestral"
    If val(TxtDiasDescuento.Text) > 180 Then
        MsgBox "Los dias de descuento no pueden ser mayor que 180, porque esta es una nómina Semestral"
        TxtDiasDescuento.SetFocus
        Exit Sub
     End If
End Select

Exit Sub
TipoErr:
  ControlErrores
End Sub

Private Sub TxtDireccion_Change()
Salida = True
End Sub

Private Sub TxtDireccion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  TxtNacionalidad.SetFocus
 Else
   Evaluar = False
  End If
End Sub


Private Sub TxtMontoDeduccion_KeyPress(KeyAscii As Integer)
 If KeyAscii = "13" Then
  CmdAgregarDeduccion.Value = True
 End If
End Sub

Private Sub TxtMontoPrestamoUS_KeyPress(KeyAscii As Integer)

If KeyAscii = "13" Then
  TxtSaldo.Text = TxtMontoPrestamoUS.Text
  Me.CmdAfectuar.Value = True

End If
End Sub

Private Sub TxtMontoPrestamoUS_LostFocus()
TxtSaldo.Text = TxtMontoPrestamoUS.Text
End Sub

Private Sub TxtMotivoAumento_Change()
Salida = True
End Sub

Private Sub TxtMotivoAumento_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  MaskEdAumento.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub TxtMotivoBaja_Change()
Salida = True
End Sub

Private Sub TxtMotivoBaja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  TxtMotivoAumento.SetFocus
 Else
   Evaluar = False
  End If
End Sub


Private Sub TxtMotivoSuspencion_Change()
Salida = True
End Sub

Private Sub TxtMotivoSuspencion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 TxtDebito.SetFocus
 Else
   Evaluar = False
  End If
End Sub


Private Sub TxtNacionalidad_Change()
Salida = True
End Sub

Private Sub TxtNacionalidad_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtCodPostal.SetFocus
 Else
   Evaluar = False
  End If
End Sub




Private Sub TxtNInss_Change()
Salida = True
End Sub

Private Sub TxtNInss_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  DBCDepartamento.SetFocus
 Else
   Evaluar = False
  End If
End Sub


Private Sub TxtNombre1_Change()
Salida = True
'PreparaSalida
End Sub

Private Sub TxtNombre1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  TxtNombre2.SetFocus
  Else
  Salida = True
  End If
 End Sub




Private Sub TxtNombre2_Change()
Salida = True
End Sub

Private Sub TxtNombre2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
  TxtApellido1.SetFocus
 Else
   Evaluar = False
  End If
End Sub




Private Sub TxtNRuc_Change()
Salida = True
End Sub

Private Sub TxtNRuc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  TxtNInss.SetFocus
 Else
   Evaluar = False
  End If
End Sub



Private Sub TxtOtrosIngresos_LostFocus()
If Not IsNumeric(TxtOtrosIngresos.Text) Then
  MsgBox "Error al digitar los Otros Ingresos"
  TxtOtrosIngresos = "0.00"
Else
  TxtOtrosIngresos = Format(val(TxtOtrosIngresos.Text), "###,##0.00")
  
  res = Bitacora(Now, NombreUsuario, "Empleados", "Se modifico Otros Ingresos: " & DBCodigoEmpleado.Text & " " & Me.TxtNombre1.Text & "Por la Cantidad de: " & Me.TxtOtrosIngresos.Text)
End If
End Sub

Private Sub TxtSueldoActual_Change()
Salida = True
End Sub

Private Sub TxtSueldoActual_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtSueldoActual.Text = Format((TxtSueldoActual.Text), "##,##0.00")
  MaskEdBaja.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub TxtSueldoActual_LostFocus()
TxtSueldoActual.Text = Format((TxtSueldoActual.Text), "##,##0.00")
End Sub

Private Sub TxtSueldoAnterior_Change()
Salida = True
End Sub

Private Sub TxtSueldoAnterior_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 TxtSueldoAnterior.Text = Format((TxtSueldoAnterior.Text), "##,##0.00")
 TxtSueldoActual.Text = Format((TxtSueldoActual.Text), "##,##0.00")
 TxtSueldoActual.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub TxtSueldoAnterior_LostFocus()
 TxtSueldoAnterior.Text = Format((TxtSueldoAnterior.Text), "##,##0.00")
End Sub

Private Sub TxtSueldoInicial_Change()
Salida = True
End Sub

Private Sub TxtSueldoInicial_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtSueldoInicial = Format((TxtSueldoInicial), "##,##0.00")
  TxtSueldoAnterior = Format((TxtSueldoAnterior), "##,##0.00")
  TxtSueldoAnterior.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub TxtSueldoInicial_LostFocus()
 TxtSueldoInicial.Text = Format((TxtSueldoInicial.Text), "##,##0.0000")
End Sub

Private Sub TxtSueldoPeriodo_Change()
Salida = True

res = Bitacora(Now, NombreUsuario, "Empleados", "Se Agrego Mofico el salario: " & DBCodigoEmpleado.Text & " " & Me.TxtNombre1.Text)
End Sub

Private Sub TxtSueldoPeriodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  TxtSueldoPeriodo.Text = Format((TxtSueldoPeriodo.Text), "##,##0.0000")
  TxtTarifaHoraria.Text = Format((TxtTarifaHoraria.Text), "##,##0.000000")
 Else
   Evaluar = False
  End If
End Sub

Private Sub TxtSueldoPeriodo_LostFocus()
 TxtSueldoPeriodo.Text = Format((TxtSueldoPeriodo.Text), "##,##0.0000")
End Sub

Private Sub TxtTarifaHoraria_Change()
Salida = True
End Sub

Private Sub TxtTarifaHoraria_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtTarifaHoraria.Text = Format((TxtTarifaHoraria.Text), "##,##0.000000")
  'CmbMonedaSueldo.SetFocus
Else
   Evaluar = False
  End If
End Sub

Private Sub TxtTarifaHoraria_LostFocus()
 TxtTarifaHoraria.Text = Format((TxtTarifaHoraria.Text), "##,##0.000000")
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub

Private Sub TxtVecesDeduccion_KeyPress(KeyAscii As Integer)
 If KeyAscii = "13" Then
  CmdAgregarDeduccion.Value = True
 End If
End Sub
