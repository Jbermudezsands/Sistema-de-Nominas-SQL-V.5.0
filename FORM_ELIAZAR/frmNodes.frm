VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmNodes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Item"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5490
   Begin VB.TextBox txtSufijo 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   11
      Top             =   240
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   5175
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   4080
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5175
      Begin VB.OptionButton chkPago 
         Caption         =   "Empresa"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton chkPago 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   3645
      End
      Begin VB.TextBox txtId 
         Height          =   285
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   1005
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "frmNodes.frx":0000
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "frmNodes.frx":006A
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmNodes.frx":00DE
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private raiz As Integer
Private llave As Integer
Private sufijo As String
Private Codigo As String
Private actividad As String
Private cliente As Boolean
Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset
Private sql As String

Private Sub cmdAgregar_Click()
On Error GoTo erradd
If Trim(Me.txtId.Text) = "" Then
    MsgBox "Codigo requerido", vbInformation
    Me.txtId.SetFocus
    Exit Sub
ElseIf Trim(Me.txtDesc.Text) = "" Then
    MsgBox "Descripción requerida", vbInformation
    Me.txtDesc.SetFocus
    Exit Sub
End If

If llave = -1 Then
    sql = "INSERT INTO [dbo].[_Actividades]([IdSuperior],sufijo,  [Codigo], [Actividad], Pagacliente) " & _
            "Values (" & raiz & ", '" & Trim(Me.txtSufijo) & "', '" & Trim(Me.txtId) & "', '" & Trim(Me.txtDesc) & "'," & IIf(Me.chkPago(1).Value, 1, 0) & ")"
    cnx.Execute sql
    
    sql = "SELECT max([idactividad]) AS llave " & _
            "From [dbo].[_Actividades] "
            
    With rs
        .CursorLocation = adUseClient
        .Open sql, cnx, adOpenDynamic, adLockOptimistic
    End With
    
    With frmActividades.tvActividades
        If raiz = 1 Then
            .Nodes.Add , tvwLast, "A" & Trim(Str(rs!llave)), Trim(Me.txtSufijo) & Trim(Me.txtId) & ".- " & Trim(Me.txtDesc), 1, 2
        Else
            .Nodes.Add "A" & Trim(Str(raiz)), tvwChild, "A" & Trim(Str(rs!llave)), Trim(Me.txtSufijo) & Trim(Me.txtId) & ".- " & Trim(Me.txtDesc), 1, 2
        End If
        .Nodes.item(.Nodes.Count).Selected = True
    End With
Else
    sql = "UPDATE _Actividades SET sufijo = " & Trim(Me.txtSufijo) & ", Codigo = " & Trim(Me.txtId) & ", [Actividad] = '" & Trim(Me.txtDesc) & "', Pagacliente = " & IIf(Me.chkPago(1).Value, 1, 0) & _
            " WHERE idactividad = " & llave
    cnx.Execute sql
    
    With frmActividades.tvActividades.SelectedItem
        .Text = Trim(Me.txtSufijo) & Trim(Me.txtId) & ".- " & Trim(Me.txtDesc)
    End With
End If

Unload Me

Exit Sub
erradd:
    If Err.Number = -2147217873 Then
        MsgBox "Operación cancelada, El código ya existe", vbInformation
        Me.txtId.SetFocus
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub cmdCerrar_Click()
Unload Me

End Sub

Private Sub Form_Activate()

MDIPrimero.Skin1.ApplySkin hWnd
End Sub

Private Sub Form_Load()
On Error GoTo errload

Me.Top = (MDIPrimero.ScaleHeight / 2) - (Me.Height / 2)
Me.Left = (MDIPrimero.ScaleWidth / 2) - (Me.Width / 2)
Me.txtSufijo.Text = sufijo

If llave = -1 Then
    Me.chkPago(0).Value = True
Else
    Me.txtId = Codigo
    Me.txtDesc = actividad
    Me.chkPago(0).Value = Not cliente
    Me.chkPago(1).Value = cliente
End If

If cnx.State = adStateClosed Then
'    sql = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PRUEBA;Data Source=WEBMASTER\SQL2005"
    cnx.ConnectionString = Conexion
    cnx.Open
End If

Exit Sub
errload:
    MsgBox Err.Description
End Sub

Public Property Let prRaiz(ByVal val As Integer)
raiz = val
End Property

Public Property Let prLlave(ByVal val As Integer)
llave = val
End Property

Public Property Let prSufijo(ByVal val As String)
sufijo = val
End Property

Public Property Let prCodigo(ByVal val As String)
Codigo = val
End Property

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmActividades.prModal = False
Set cnx = Nothing
Set rs = Nothing
End Sub

Public Property Let prActividad(ByVal val As String)
actividad = val
End Property

Public Property Let prCliente(ByVal val As Boolean)
cliente = val
End Property
 
