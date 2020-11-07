VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form FrmPeriodoFiscal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Periodo  Fiscal"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   475
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   600
      Top             =   8160
      Width           =   3855
      _ExtentX        =   6800
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
   Begin VB.ListBox lstPlanilla 
      Columns         =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   6615
   End
   Begin MSAdodcLib.Adodc DtaTipoNomina 
      Height          =   375
      Left            =   600
      Top             =   7080
      Width           =   3735
      _ExtentX        =   6588
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
   Begin VB.Frame fraPlanilla 
      Caption         =   "Planilla"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   4200
         OleObjectBlob   =   "FrmPeriodoFiscal.frx":0000
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   2160
         OleObjectBlob   =   "FrmPeriodoFiscal.frx":0072
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmPeriodoFiscal.frx":00E6
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "FrmPeriodoFiscal.frx":0162
         TabIndex        =   11
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo DBTipoNomina 
         Bindings        =   "FrmPeriodoFiscal.frx":01DC
         DataSource      =   "DtaTipoNomina"
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nomina"
         Text            =   ""
      End
      Begin VB.OptionButton optPlanSemanal 
         Caption         =   "Planilla Semanal"
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   2520
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton optPlanQuincenal 
         Caption         =   "Planilla Quincenal"
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   2400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker mskFinal 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Format          =   48693249
         CurrentDate     =   38483
      End
      Begin MSComCtl2.DTPicker mskInicio 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   48693249
         CurrentDate     =   38483
      End
      Begin VB.TextBox txtAño 
         Height          =   285
         Left            =   4200
         MaxLength       =   5
         TabIndex        =   1
         Text            =   " "
         Top             =   480
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmPeriodoFiscal.frx":01F8
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc DtaFecha 
      Height          =   330
      Left            =   600
      Top             =   7560
      Width           =   3855
      _ExtentX        =   6800
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
      Caption         =   "DtaFecha"
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
   Begin MSAdodcLib.Adodc DtaAño 
      Height          =   375
      Left            =   600
      Top             =   7800
      Width           =   3855
      _ExtentX        =   6800
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
      Caption         =   "DtaAño"
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
   Begin SmartButtonProject.SmartButton cmdSalir 
      Height          =   975
      Left            =   5640
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      Caption         =   "Salir"
      Picture         =   "FrmPeriodoFiscal.frx":026A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SmartButtonProject.SmartButton cmdGenerar 
      Height          =   975
      Left            =   2880
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      Caption         =   "Generar"
      Picture         =   "FrmPeriodoFiscal.frx":108C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SmartButtonProject.SmartButton cmdGuardar 
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      Caption         =   "Guardar"
      Picture         =   "FrmPeriodoFiscal.frx":2AE2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmPeriodoFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iAño As Integer, TipoNomina As String

Public Function DepurarFecha(TxtFecha As Control) As Boolean

Dim iCont As Integer
Dim sCad As String
Dim sCad2 As String

  Do While iCont < 3
  
  
  Select Case iCont
  
   Case 0:
     sCad = Mid$(TxtFecha, 1, 2)
     sCad2 = Mid$(TxtFecha, 4, 2)
     sCad = Trim(sCad)
     If Not IsNumeric(sCad) Then
         DepurarFecha = True
         Exit Function
     ElseIf CInt(sCad2) = 1 Or CInt(sCad2) = 3 Or _
            CInt(sCad2) = 5 Or CInt(sCad2) = 7 Or _
            CInt(sCad2) = 8 Or CInt(sCad2) = 10 Or _
            CInt(sCad2) = 12 Then
            If Not (CInt(sCad) >= 1 And CInt(sCad) <= 31) Then
               DepurarFecha = True
               Exit Function
            End If
        ElseIf CInt(sCad2) = 2 Then
               If Not (CInt(sCad) >= 1 And CInt(sCad) <= 28) Then
                    DepurarFecha = True
                    Exit Function
                End If
                
         ElseIf Not (CInt(sCad) >= 1 And CInt(sCad) <= 30) Then
                    DepurarFecha = True
                    Exit Function
                    
         ElseIf Not (CInt(sCad2) >= 1 And CInt(sCad2) <= 12) Then
                    DepurarFecha = True
                    Exit Function
              
             
      End If
      
   Case 1:
     sCad = Mid$(TxtFecha, 4, 2)
     sCad = Trim(sCad)
   Case 2:
     sCad = Mid$(TxtFecha, 7, 4)
     sCad = Trim(sCad)
     
     If Not IsNumeric(sCad) Then
        If CInt(sCad) <= 1995 Then
          DepurarFecha = True
          Exit Function
        End If
        
     End If
     
  End Select
  
     If Not IsNumeric(sCad) Then
         DepurarFecha = True
         Exit Function
     End If
      
            
     
   iCont = iCont + 1
   
  
  Loop

DepurarFecha = False

End Function



Private Sub CmdGenerar_Click()
Dim iPeriodo As Integer
Dim saMes As Variant, Lineas As Integer
Dim saMes2 As Variant, Mes As Integer
Dim iCont As Integer, MesLetra As String
Dim bContMes As Byte
Dim iMes As Integer
Dim iSem As Integer
Dim FechaIni As Date
Dim FechaFin As Date
Dim iAño As Integer
Dim sIntervalo As String
Dim bBisiesto As Boolean
Dim Fechas As Date

iPeriodo = 1
iSem = 1
iCont = 0

'On Error GoTo ManipularError

Select Case TipoNomina
 Case "Semana Sabado"

       If DepurarFecha(Me.mskInicio) Then
          MsgBox "La Fecha Inicial no es válida, corrija", vbInformation
          mskInicio.SetFocus
          Exit Sub
       ElseIf DepurarFecha(Me.mskFinal) Then
          MsgBox "La Fecha Final no es válida, corrija", vbInformation
          mskFinal.SetFocus
          Exit Sub
       End If



        lstPlanilla.Clear
        FechaIni = Me.mskInicio.Value
        FechaFin = FechaIni + 6
        saMes = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    
    If Month(Me.mskInicio.Value) = 1 Then
     saMes2 = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    ElseIf Month(Me.mskInicio.Value) = 7 Then
     saMes2 = Array("Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio")
    End If

        Me.DtaAño.RecordSource = "SELECT CodTipoNomina,Año, Actual, FecIni, FecFin From Año_Fiscal Where (Actual = 1) AND (CodTipoNomina='" & CodTipoNomina & "')"
        Me.DtaAño.Refresh
        If Not DtaAño.Recordset.EOF Then

            'DtaAño.Recordset("Actual") = 0
            'DtaAño.Recordset.Update
        End If
   
        If IsNumeric(Me.TxtAño.Text) Then
            Me.DtaAño.RecordSource = "SELECT Año, Actual, FecIni, FecFin, CodTipoNomina From Año_Fiscal Where (Año Like '" & Me.TxtAño.Text & "')"
            Me.DtaAño.Refresh
            If Me.DtaAño.Recordset.EOF Then
                Me.DtaAño.Recordset.AddNew
                DtaAño.Recordset("Año") = CInt(TxtAño.Text)
                DtaAño.Recordset("FecIni") = Me.mskInicio.Value
                DtaAño.Recordset("FecFin") = Me.mskFinal.Value
                DtaAño.Recordset("CodTipoNomina") = CodTipoNomina
                DtaAño.Recordset("Actual") = 1
                Me.DtaAño.Recordset.Update
            Else
                'Me.'DtaAño.Recordset.Edit
                DtaAño.Recordset("Actual") = 1
                Me.DtaAño.Recordset.Update
            End If
        Else
            MsgBox "El año actual de planilla no es válido", vbInformation, "Error!!!"
            Exit Sub
        End If
  
        Semanas
  'Hay que setear el año actual de planilla OJO
  'Para que no de problemas
        Lineas = 1
        If FechaFin < Me.mskFinal.Value Then
            Do While FechaFin <= Me.mskFinal.Value  'Hay que setear el año actual de planilla OJO
                'Me.dtaSem.Recordset.FindFirst "[Mes] like '" & saMes(iCont) & "' AND [Año] like " & Me.DtaAño.Recordset("Año") & ""
                'iMes = 'Me.dtaSem.Recordset.Fields(2)
                Fechas = "01/" & saMes(iCont) & "/" & Me.TxtAño.Text
                NumFecha2 = Fechas
                iMes = SabadosMes(NumFecha2)
                lstPlanilla.AddItem "        " & saMes2(iCont)
  
  
               Do While iMes >= iSem 'And iPeriodo <= 52
                    'GuardarPeriodo
                    lstPlanilla.AddItem " " & CStr(iPeriodo) & "   " & CStr(FechaIni) & "  al  " & CStr(FechaFin)
                    FechaIni = FechaFin + 1
                    FechaFin = FechaIni + 6
                    iSem = iSem + 1
                    iPeriodo = iPeriodo + 1
                    Lineas = Lineas + 1
                Loop
  
                iSem = 1
                iCont = iCont + 1
                lstPlanilla.AddItem "      "
                If Lineas = 14 Then
                    lstPlanilla.AddItem "      "
                End If
            Loop

        Else
            MsgBox "El intervalo de la fecha inicial y final no es válido", vbInformation, "corrija"
        End If
 Case "Semanal Viernes"

       If DepurarFecha(Me.mskInicio) Then
          MsgBox "La Fecha Inicial no es válida, corrija", vbInformation
          mskInicio.SetFocus
          Exit Sub
       ElseIf DepurarFecha(Me.mskFinal) Then
          MsgBox "La Fecha Final no es válida, corrija", vbInformation
          mskFinal.SetFocus
          Exit Sub
       End If



        lstPlanilla.Clear
        FechaIni = Me.mskInicio.Value
        FechaFin = FechaIni + 6
        saMes = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    
    If Month(Me.mskInicio.Value) = 1 Then
     saMes2 = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    ElseIf Month(Me.mskInicio.Value) = 7 Then
     saMes2 = Array("Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio")
    End If

        Me.DtaAño.RecordSource = "SELECT CodTipoNomina,Año, Actual, FecIni, FecFin From Año_Fiscal Where (Actual = 1) AND (CodTipoNomina='" & CodTipoNomina & "')"
        Me.DtaAño.Refresh
        If Not DtaAño.Recordset.EOF Then

            'DtaAño.Recordset("Actual") = 0
            'DtaAño.Recordset.Update
        End If
   
        If IsNumeric(Me.TxtAño.Text) Then
            Me.DtaAño.RecordSource = "SELECT Año, Actual, FecIni, FecFin, CodTipoNomina From Año_Fiscal Where (Año Like '" & Me.TxtAño.Text & "')"
            Me.DtaAño.Refresh
            If Me.DtaAño.Recordset.EOF Then
                Me.DtaAño.Recordset.AddNew
                DtaAño.Recordset("Año") = CInt(TxtAño.Text)
                DtaAño.Recordset("FecIni") = Me.mskInicio.Value
                DtaAño.Recordset("FecFin") = Me.mskFinal.Value
                DtaAño.Recordset("CodTipoNomina") = CodTipoNomina
                DtaAño.Recordset("Actual") = 1
                Me.DtaAño.Recordset.Update
            Else
                'Me.'DtaAño.Recordset.Edit
                DtaAño.Recordset("Actual") = 1
                Me.DtaAño.Recordset.Update
            End If
        Else
            MsgBox "El año actual de planilla no es válido", vbInformation, "Error!!!"
            Exit Sub
        End If
  
        Semanas
  'Hay que setear el año actual de planilla OJO
  'Para que no de problemas
        Lineas = 1
        If FechaFin < Me.mskFinal.Value Then
            Do While FechaFin <= Me.mskFinal.Value  'Hay que setear el año actual de planilla OJO
                'Me.dtaSem.Recordset.FindFirst "[Mes] like '" & saMes(iCont) & "' AND [Año] like " & Me.DtaAño.Recordset("Año") & ""
                'iMes = 'Me.dtaSem.Recordset.Fields(2)
                Fechas = "01/" & saMes(iCont) & "/" & Me.TxtAño.Text
                NumFecha2 = Fechas
                iMes = ViernesMes(NumFecha2)
                lstPlanilla.AddItem "        " & saMes2(iCont)
  
  
               Do While iMes >= iSem 'And iPeriodo <= 52
                    'GuardarPeriodo
                    lstPlanilla.AddItem " " & CStr(iPeriodo) & "   " & CStr(FechaIni) & "  al  " & CStr(FechaFin)
                    FechaIni = FechaFin + 1
                    FechaFin = FechaIni + 6
                    iSem = iSem + 1
                    iPeriodo = iPeriodo + 1
                    Lineas = Lineas + 1
                Loop
  
                iSem = 1
                iCont = iCont + 1
                lstPlanilla.AddItem "      "
                If Lineas = 14 Then
                    lstPlanilla.AddItem "      "
                End If
            Loop

        Else
            MsgBox "El intervalo de la fecha inicial y final no es válido", vbInformation, "corrija"
        End If

Case "Quincenal"
 

        Me.DtaAño.RecordSource = "SELECT CodTipoNomina,Año, Actual, FecIni, FecFin From Año_Fiscal Where (Actual = 1) AND (CodTipoNomina='" & CodTipoNomina & "')"
        Me.DtaAño.Refresh
        If Not DtaAño.Recordset.EOF Then

            'DtaAño.Recordset("Actual") = 0
            'DtaAño.Recordset.Update
        End If
    Me.DtaFecha.RecordSource = "PeriodoFiscal"


    Me.DtaAño.Refresh
    Me.DtaFecha.Refresh
    'Me.dtaSem.Refresh

    lstPlanilla.Clear
    FechaIni = CDate(Me.mskInicio.Value) - 1
    iAño = Me.TxtAño.Text
    saMes = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    
    If Month(Me.mskInicio.Value) = 1 Then
     saMes2 = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    ElseIf Month(Me.mskInicio.Value) = 7 Then
     saMes2 = Array("Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio")
    End If
  
   
   
    sIntervalo = CDbl(CSng(Year(Me.mskFinal.Value)) / 4)
     
    If InStr(1, sIntervalo, ".", vbTextCompare) = 0 Then
         bBisiesto = True
    End If
    Lineas = 1
    If FechaFin < Me.mskFinal.Value Then
        Do While FechaFin <= Me.mskFinal.Value  'Hay que setear el año actual de planilla OJO
      
            lstPlanilla.AddItem "        " & saMes2(iCont)
            bContMes = 0
            FechaIni = FechaIni + 1
            FechaFin = FechaIni + 14
            Do While iMes < 24 And bContMes < 2
                'GuardarPeriodo
                iMes = iMes + 1
                lstPlanilla.AddItem " " & CStr(iMes) & "   " & CStr(FechaIni) & "  al  " & CStr(FechaFin)
                bContMes = bContMes + 1
                FechaIni = FechaFin
                FechaFin = FechaFin + FechaQuincenal(bBisiesto, FechaFin)
                Lineas = Lineas + 1
                If bContMes = 1 Then
                  FechaIni = FechaIni + 1
                End If
            Loop
              
            iCont = iCont + 1
            'FechaIni = FechaIni + 1
            lstPlanilla.AddItem "      "
            If Lineas = 9 Then
                lstPlanilla.AddItem "      "
                lstPlanilla.AddItem "      "
            End If
  
        Loop
    
    End If
    
    
Case "Mensual"
  
    'Me.DtaAño.DatabaseName = App.Path + "\PlanQuince.mdb"
    'Me.DtaFecha.DatabaseName = App.Path + "\PlanQuince.mdb"
    'Me.dtaSem.DatabaseName = App.Path + "\PlanQuince.mdb"

        Me.DtaAño.RecordSource = "SELECT CodTipoNomina,Año, Actual, FecIni, FecFin From Año_Fiscal Where (Actual = 1) AND (CodTipoNomina='" & CodTipoNomina & "')"
        Me.DtaAño.Refresh
        If Not DtaAño.Recordset.EOF Then

            'DtaAño.Recordset("Actual") = 0
            'DtaAño.Recordset.Update
        End If
    Me.DtaFecha.RecordSource = "PeriodoFiscal"
    'Me.dtaSem.RecordSource = "SabadosMes"

    Me.DtaAño.Refresh
    Me.DtaFecha.Refresh
    'Me.dtaSem.Refresh

    lstPlanilla.Clear
    FechaIni = CDate(Me.mskInicio.Value) - 1
    iAño = Me.TxtAño.Text
    saMes = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    
    If Month(Me.mskInicio.Value) = 1 Then
     saMes2 = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    ElseIf Month(Me.mskInicio.Value) = 7 Then
     saMes2 = Array("Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio")
    End If
   
    sIntervalo = CDbl(CSng(Year(Me.mskFinal.Value)) / 1)
     
    If InStr(1, sIntervalo, ".", vbTextCompare) = 0 Then
         bBisiesto = True
    End If
    Lineas = 1
    If FechaFin < Me.mskFinal.Value Then
        Do While FechaFin < Me.mskFinal.Value  'Hay que setear el año actual de planilla OJO
      
            lstPlanilla.AddItem "        " & saMes2(iCont)
            bContMes = 0
 
 
               
'            Do While iMes < 12 'And bContMes < 2
                'GuardarPeriodo
                 iMes = iMes + 1
                If Mes = 0 Then
                 FechaIni = CDate(Me.mskInicio.Value)
                 Mes = Month(FechaIni)
                Else
                 FechaIni = FechaFin + 1
                 Mes = Month(FechaIni)
                 iAño = Year(FechaIni)
                End If
'                 Fechaini = DateSerial(iAño, iMes, 1)
                 FechaFin = DateSerial(iAño, Mes + 1, 0)
                lstPlanilla.AddItem " " & CStr(iMes) & "   " & CStr(FechaIni) & "  al  " & CStr(FechaFin)
'                bContMes = bContMes + 1
'                Fechaini = Fechafin
'                Fechafin = Fechafin + FechaQuincenal(bBisiesto, Fechafin)
                Lineas = Lineas + 1
                If bContMes = 1 Then
                  FechaIni = FechaIni + 1
                End If
'            Loop
              
            iCont = iCont + 1
            'FechaIni = FechaIni + 1
            lstPlanilla.AddItem "      "
            If Lineas = 9 Then
                lstPlanilla.AddItem "      "
                lstPlanilla.AddItem "      "
            End If
  
        Loop
    
    End If
  
End Select
  



Exit Sub

ManipularError:
 MsgBox Err.Description


End Sub

Private Sub cmdGuardar_Click()
'On Error GoTo TipoErrs
Dim iPeriodo As Integer
Dim saMes As Variant
Dim iCont As Integer
Dim bContMes As Byte
Dim iMes As Integer
Dim iSem As Integer
Dim FechaIni As String
Dim FechaFin As Date
Dim bEncontrado As Boolean
Dim iAño As Integer
Dim sIntervalo As String
Dim bBisiesto As Boolean
Dim Numfecha As Long, NFecha1 As Long, NFecha2 As Long
Dim Fechas As String


iPeriodo = 1
iSem = 1
iCont = 0



 If DepurarFecha(Me.mskInicio) Then
    MsgBox "La Fecha Inicial no es válida, corrija", vbInformation
    mskInicio.SetFocus
    Exit Sub
 ElseIf DepurarFecha(Me.mskFinal) Then
    MsgBox "La Fecha Final no es válida, corrija", vbInformation
    mskFinal.SetFocus
    Exit Sub
 End If



  
 Me.DtaAño.RecordSource = "SELECT Año, Actual, CodTipoNomina, FecIni, FecFin From Año_Fiscal WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
 Me.DtaAño.Refresh
    'DtaAño.Recordset.Edit
   If Not Me.DtaAño.Recordset.EOF Then
    DtaAño.Recordset("Actual") = 0
    DtaAño.Recordset.Update
   End If
   
   
 Select Case TipoNomina
 Case "Semanal Sabado"

        If IsNumeric(Me.TxtAño.Text) Then
            Me.DtaAño.RecordSource = "SELECT CodTipoNomina,Año, Actual, FecIni, FecFin From Año_Fiscal Where (Año Like '" & Me.TxtAño.Text & "')AND (CodTipoNomina = '" & CodTipoNomina & "')"
            Me.DtaAño.Refresh
            If Me.DtaAño.Recordset.EOF Then
                Me.DtaAño.Recordset.AddNew
                DtaAño.Recordset("Año") = CInt(TxtAño.Text)
                DtaAño.Recordset("Actual") = 1
                DtaAño.Recordset.Fields("FecIni") = Me.mskInicio.Value
                DtaAño.Recordset.Fields("FecFin") = Me.mskFinal.Value
                DtaAño.Recordset("CodTipoNomina") = CodTipoNomina
                Me.DtaAño.Recordset.Update
            Else
                'Me.'DtaAño.Recordset.Edit
                DtaAño.Recordset("Actual") = 1
                DtaAño.Recordset.Fields("FecIni") = Me.mskInicio.Value
                DtaAño.Recordset.Fields("FecFin") = Me.mskFinal.Value
                Me.DtaAño.Recordset.Update
            End If
        Else
            MsgBox "El año actual de planilla no es válido", vbInformation, "Error!!!"
            Exit Sub
        End If

        lstPlanilla.Clear
        FechaIni = Me.mskInicio.Value
        FechaFin = CDate(FechaIni) + 6
        saMes = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    
    If Month(Me.mskInicio.Value) = 1 Then
     saMes2 = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    ElseIf Month(Me.mskInicio.Value) = 7 Then
     saMes2 = Array("Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio")
    End If

        
        Semanas

        Fechas = Format(FechaIni, "yyyy/mm/dd")
        Me.DtaFecha.RecordSource = "SELECT CodTipoNomina, Periodo, año, mes, Inicio, Final, Actual From PeriodoFiscal Where (Inicio = '" & Fechas & "') AND (CodTipoNomina = '" & CodTipoNomina & "')"
        Me.DtaFecha.Refresh


        If Me.DtaFecha.Recordset.EOF Then
            bEncontrado = False
        Else
            bEncontrado = True
        End If


        If FechaFin < Me.mskFinal.Value Then
            Do While FechaFin <= Me.mskFinal.Value
                    Fechas = "01/" & saMes(iCont) & "/" & Me.TxtAño.Text
                    NumFecha2 = CDate(Fechas)
                    iMes = SabadosMes(NumFecha2)
     
                    lstPlanilla.AddItem "        " & saMes2(iCont)
  
  
                    Do While iMes >= iSem 'And iPeriodo <= 52
                            If bEncontrado Then
                                GoSub EditarPeriodo
                            Else
                                GoSub GuardarPeriodo
                            End If
       
                            lstPlanilla.AddItem " " & CStr(iPeriodo) & "   " & CStr(FechaIni) & " al  " & CStr(FechaFin)
                            FechaIni = FechaFin + 1
                            FechaFin = CDate(FechaIni) + 6
                            iSem = iSem + 1
                            iPeriodo = iPeriodo + 1
                    Loop
  
                    iSem = 1
                    iCont = iCont + 1
                    lstPlanilla.AddItem "      "
  
            Loop
 
        Else
            MsgBox "El intervalo de la fecha inicial y final no es válido", vbInformation, "corrija"

        End If
 Case "Semanal Viernes"

        If IsNumeric(Me.TxtAño.Text) Then
            Me.DtaAño.RecordSource = "SELECT CodTipoNomina,Año, Actual, FecIni, FecFin From Año_Fiscal Where (Año Like '" & Me.TxtAño.Text & "')AND (CodTipoNomina = '" & CodTipoNomina & "')"
            Me.DtaAño.Refresh
            If Me.DtaAño.Recordset.EOF Then
                Me.DtaAño.Recordset.AddNew
                DtaAño.Recordset("Año") = CInt(TxtAño.Text)
                DtaAño.Recordset("Actual") = 1
                DtaAño.Recordset.Fields("FecIni") = Me.mskInicio.Value
                DtaAño.Recordset.Fields("FecFin") = Me.mskFinal.Value
                DtaAño.Recordset("CodTipoNomina") = CodTipoNomina
                Me.DtaAño.Recordset.Update
            Else
                'Me.'DtaAño.Recordset.Edit
                DtaAño.Recordset("Actual") = 1
                DtaAño.Recordset.Fields("FecIni") = Me.mskInicio.Value
                DtaAño.Recordset.Fields("FecFin") = Me.mskFinal.Value
                Me.DtaAño.Recordset.Update
            End If
        Else
            MsgBox "El año actual de planilla no es válido", vbInformation, "Error!!!"
            Exit Sub
        End If

        lstPlanilla.Clear
        FechaIni = Me.mskInicio.Value
        FechaFin = CDate(FechaIni) + 6
        saMes = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    
    If Month(Me.mskInicio.Value) = 1 Then
     saMes2 = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    ElseIf Month(Me.mskInicio.Value) = 7 Then
     saMes2 = Array("Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio")
    End If

        
        Semanas

        Fechas = Format(FechaIni, "yyyy/mm/dd")
        Me.DtaFecha.RecordSource = "SELECT CodTipoNomina, Periodo, año, mes, Inicio, Final, Actual From PeriodoFiscal Where (Inicio = '" & Fechas & "') AND (CodTipoNomina = '" & CodTipoNomina & "')"
        Me.DtaFecha.Refresh


        If Me.DtaFecha.Recordset.EOF Then
            bEncontrado = False
        Else
            bEncontrado = True
        End If


        If FechaFin < Me.mskFinal.Value Then
            Do While FechaFin <= Me.mskFinal.Value
                    Fechas = "01/" & saMes(iCont) & "/" & Me.TxtAño.Text
                    NumFecha2 = CDate(Fechas)
                    iMes = ViernesMes(NumFecha2)
     
                    lstPlanilla.AddItem "        " & saMes2(iCont)
  
  
                    Do While iMes >= iSem 'And iPeriodo <= 52
                            If bEncontrado Then
                                GoSub EditarPeriodo
                            Else
                                GoSub GuardarPeriodo
                            End If
       
                            lstPlanilla.AddItem " " & CStr(iPeriodo) & "   " & CStr(FechaIni) & " al  " & CStr(FechaFin)
                            FechaIni = FechaFin + 1
                            FechaFin = CDate(FechaIni) + 6
                            iSem = iSem + 1
                            iPeriodo = iPeriodo + 1
                    Loop
  
                    iSem = 1
                    iCont = iCont + 1
                    lstPlanilla.AddItem "      "
  
            Loop
 
        Else
            MsgBox "El intervalo de la fecha inicial y final no es válido", vbInformation, "corrija"

        End If
 
Case "Quincenal" '/////////////Aqui comiensa el proceso para la PLANILLA QUINCENAL///////////////////////////
        
        If IsNumeric(Me.TxtAño.Text) Then
            Me.DtaAño.RecordSource = "SELECT CodTipoNomina,Año, Actual, FecIni, FecFin From Año_Fiscal Where (Año Like '" & Me.TxtAño.Text & "')AND (CodTipoNomina = '" & CodTipoNomina & "')"
            Me.DtaAño.Refresh
            If Me.DtaAño.Recordset.EOF Then
                Me.DtaAño.Recordset.AddNew
                DtaAño.Recordset("CodTipoNomina") = CodTipoNomina
                DtaAño.Recordset("Año") = CInt(TxtAño.Text)
                DtaAño.Recordset("Actual") = 1
                DtaAño.Recordset.Fields("FecIni") = Me.mskInicio.Value
                DtaAño.Recordset.Fields("FecFin") = Me.mskFinal.Value
                Me.DtaAño.Recordset.Update
            Else
                DtaAño.Recordset("Actual") = 1
                DtaAño.Recordset.Fields("FecIni") = Me.mskInicio.Value
                DtaAño.Recordset.Fields("FecFin") = Me.mskFinal.Value
                Me.DtaAño.Recordset.Update
            End If
        Else
            MsgBox "El año actual de planilla no es válido", vbInformation, "Error!!!"
            Exit Sub
        End If
 
        Me.DtaAño.Refresh
        Me.DtaFecha.Refresh


        lstPlanilla.Clear
        FechaIni = CDate(Me.mskInicio.Value) - 1
        iAño = Me.TxtAño.Text
        saMes = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    
    If Month(Me.mskInicio.Value) = 1 Then
     saMes2 = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    ElseIf Month(Me.mskInicio.Value) = 7 Then
     saMes2 = Array("Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio")
    End If

   
    sIntervalo = CDbl(CSng(Year(Me.mskFinal.Value)) / 4)
     
        If InStr(1, sIntervalo, ".", vbTextCompare) = 0 Then
            bBisiesto = True
        End If
   

        FechaIni = CDate(FechaIni) + 1
        'Numfecha = Fechaini
        'Me.DtaFecha.Recordset.FindFirst "[Inicio] = DateValue('" & FechaIni + 1 & "')"
        Me.DtaFecha.RecordSource = "SELECT CodTipoNomina, Periodo, año, mes, Inicio, Final, Actual From PeriodoFiscal Where (Inicio = '" & CDate(FechaIni) & "') AND (CodTipoNomina = '" & CodTipoNomina & "')"
        Me.DtaFecha.Refresh
  
          FechaIni = CDate(FechaIni) - 1
        If Me.DtaFecha.Recordset.EOF Then
            bEncontrado = False
        Else
            bEncontrado = True
        End If
   
   
        If FechaFin < Me.mskFinal.Value Then
            Do While FechaFin <= Me.mskFinal.Value  'Hay que setear el año actual de planilla OJO
      
                    lstPlanilla.AddItem "        " & saMes2(iCont)
                    bContMes = 0
                    FechaIni = CDate(FechaIni) + 1
                    FechaFin = CDate(FechaIni) + 14
       
                    Do While iMes < 24 And bContMes < 2
                        If bEncontrado Then
                            GoSub EditarPeriodo
                        Else
                            GoSub GuardarPeriodo
                        End If
        
                        iMes = iMes + 1
                        lstPlanilla.AddItem " " & CStr(FechaIni) & "  al  " & CStr(FechaFin)
                        bContMes = bContMes + 1
                        FechaIni = FechaFin
                        FechaFin = FechaFin + FechaQuincenal(bBisiesto, FechaFin)
                        iPeriodo = iPeriodo + 1
                        If bContMes = 1 Then
                            FechaIni = CDate(FechaIni) + 1
                        End If
                
                    Loop
              
                    iCont = iCont + 1
                    lstPlanilla.AddItem "      "
  
            Loop
        End If

Case "Mensual"

        If IsNumeric(Me.TxtAño.Text) Then
            Me.DtaAño.RecordSource = "SELECT CodTipoNomina,Año, Actual, FecIni, FecFin From Año_Fiscal Where (Año Like '" & Me.TxtAño.Text & "')AND (CodTipoNomina = '" & CodTipoNomina & "')"
            Me.DtaAño.Refresh
            If Me.DtaAño.Recordset.EOF Then
                Me.DtaAño.Recordset.AddNew
                DtaAño.Recordset("CodTipoNomina") = CodTipoNomina
                DtaAño.Recordset("Año") = CInt(TxtAño.Text)
                DtaAño.Recordset("Actual") = 1
                DtaAño.Recordset.Fields("FecIni") = Me.mskInicio.Value
                DtaAño.Recordset.Fields("FecFin") = Me.mskFinal.Value
                Me.DtaAño.Recordset.Update
            Else
                DtaAño.Recordset("Actual") = 1
                DtaAño.Recordset.Fields("FecIni") = Me.mskInicio.Value
                DtaAño.Recordset.Fields("FecFin") = Me.mskFinal.Value
                Me.DtaAño.Recordset.Update
            End If
        Else
            MsgBox "El año actual de planilla no es válido", vbInformation, "Error!!!"
            Exit Sub
        End If
  
   
    
    Me.DtaFecha.RecordSource = "PeriodoFiscal"


    Me.DtaAño.Refresh
    Me.DtaFecha.Refresh
    'Me.dtaSem.Refresh

    lstPlanilla.Clear
    FechaIni = CDate(Me.mskInicio.Value) - 1
    iAño = Me.TxtAño.Text
    saMes = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
   
    If Month(Me.mskInicio.Value) = 1 Then
     saMes2 = Array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    ElseIf Month(Me.mskInicio.Value) = 7 Then
     saMes2 = Array("Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio")
    End If

    sIntervalo = CDbl(CSng(Year(Me.mskFinal.Value)) / 1)
'    sIntervalo = CDbl(CSng(Me.txtAño.Text) / 1)
     
    If InStr(1, sIntervalo, ".", vbTextCompare) = 0 Then
         bBisiesto = True
    End If
    Lineas = 1
    
        
        Me.DtaFecha.RecordSource = "SELECT CodTipoNomina, Periodo, año, mes, Inicio, Final, Actual From PeriodoFiscal Where (Inicio = '" & Me.mskInicio.Value & "') AND (CodTipoNomina = '" & CodTipoNomina & "')"
        Me.DtaFecha.Refresh
  
        If Me.DtaFecha.Recordset.EOF Then
            bEncontrado = False
        Else
            bEncontrado = True
        End If
    
    iPeriodo = 0
    If FechaFin < Me.mskFinal.Value Then
        Do While FechaFin < Me.mskFinal.Value  'Hay que setear el año actual de planilla OJO
      
            lstPlanilla.AddItem "        " & saMes2(iCont)
            bContMes = 0
 
 
        
               
'            Do While iMes < 12 'And bContMes < 2
                'GuardarPeriodo
'                iMes = iMes + 1
'                 Fechaini = DateSerial(iAño, iMes, 1)
                                
               iMes = iMes + 1
                If Mes = 0 Then
                 FechaIni = CDate(Me.mskInicio.Value)
                 Mes = Month(FechaIni)
                Else
                 FechaIni = FechaFin + 1
                 Mes = Month(FechaIni)
                 iAño = Year(FechaIni)
                End If
                
                 FechaFin = DateSerial(iAño, Mes + 1, 0)
                 iPeriodo = iPeriodo + 1
                 If bEncontrado Then
                     GoSub EditarPeriodo
                 Else
                     GoSub GuardarPeriodo
                 End If
                 
                lstPlanilla.AddItem " " & CStr(iMes) & "   " & CStr(FechaIni) & "  al  " & CStr(FechaFin)
'                bContMes = bContMes + 1
'                Fechaini = Fechafin
'                Fechafin = Fechafin + FechaQuincenal(bBisiesto, Fechafin)
                Lineas = Lineas + 1
                If bContMes = 1 Then
                  FechaIni = FechaIni + 1
                End If
'            Loop
              
            iCont = iCont + 1
            'FechaIni = FechaIni + 1
            lstPlanilla.AddItem "      "
            If Lineas = 9 Then
                lstPlanilla.AddItem "      "
                lstPlanilla.AddItem "      "
            End If
  
        Loop
    
    End If



End Select
 
 
 
   Exit Sub
 
GuardarPeriodo:
 
Select Case TipoNomina
 Case "Semanal Sabado"
    Me.DtaAño.RecordSource = "SELECT Año, Actual, CodTipoNomina, FecIni, FecFin From Año_Fiscal WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
    Me.DtaAño.Refresh
    Me.DtaAño.Refresh

     Me.DtaFecha.Recordset.AddNew
     Me.DtaFecha.Recordset.Fields("CodTipoNomina") = CodTipoNomina
     Me.DtaFecha.Recordset.Fields("Periodo") = iPeriodo
     Me.DtaFecha.Recordset.Fields("año") = Me.DtaAño.Recordset("Año")
     Me.DtaFecha.Recordset.Fields("mes") = saMes(iCont)
     Me.DtaFecha.Recordset.Fields("Inicio") = CStr(FechaIni)
     Me.DtaFecha.Recordset.Fields("Final") = CStr(FechaFin)
     If iPeriodo = 1 Then
       Me.DtaFecha.Recordset.Fields("Actual") = 1
     Else
       Me.DtaFecha.Recordset.Fields("Actual") = 0
     End If
     Me.DtaFecha.Recordset.Update
 Case "Semanal Viernes"
    Me.DtaAño.RecordSource = "SELECT Año, Actual, CodTipoNomina, FecIni, FecFin From Año_Fiscal WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
    Me.DtaAño.Refresh
    Me.DtaAño.Refresh

     Me.DtaFecha.Recordset.AddNew
     Me.DtaFecha.Recordset.Fields("CodTipoNomina") = CodTipoNomina
     Me.DtaFecha.Recordset.Fields("Periodo") = iPeriodo
     Me.DtaFecha.Recordset.Fields("año") = Me.DtaAño.Recordset("Año")
     Me.DtaFecha.Recordset.Fields("mes") = saMes(iCont)
     Me.DtaFecha.Recordset.Fields("Inicio") = CStr(FechaIni)
     Me.DtaFecha.Recordset.Fields("Final") = CStr(FechaFin)
     If iPeriodo = 1 Then
       Me.DtaFecha.Recordset.Fields("Actual") = 1
     Else
       Me.DtaFecha.Recordset.Fields("Actual") = 0
     End If
     Me.DtaFecha.Recordset.Update
     
Case "Quincenal"
  
     Me.DtaFecha.Recordset.AddNew
     Me.DtaFecha.Recordset.Fields("CodTipoNomina") = CodTipoNomina
     Me.DtaFecha.Recordset.Fields("Periodo") = iPeriodo
     Me.DtaFecha.Recordset.Fields("año") = iAño
     Me.DtaFecha.Recordset.Fields("mes") = saMes(iCont)
     Me.DtaFecha.Recordset.Fields("Inicio") = FechaIni
     Me.DtaFecha.Recordset.Fields("Final") = FechaFin
     If iPeriodo = 1 Then
      Me.DtaFecha.Recordset.Fields("Actual") = 1
     Else
      Me.DtaFecha.Recordset.Fields("Actual") = 0
     End If
     Me.DtaFecha.Recordset.Update

 Case "Mensual"
    Me.DtaAño.RecordSource = "SELECT Año, Actual, CodTipoNomina, FecIni, FecFin From Año_Fiscal WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
    Me.DtaAño.Refresh
    Me.DtaAño.Refresh

     Me.DtaFecha.Recordset.AddNew
     Me.DtaFecha.Recordset.Fields("CodTipoNomina") = CodTipoNomina
     Me.DtaFecha.Recordset.Fields("Periodo") = iPeriodo
     Me.DtaFecha.Recordset.Fields("año") = iAño
     Me.DtaFecha.Recordset.Fields("mes") = saMes(iCont)
     Me.DtaFecha.Recordset.Fields("Inicio") = CStr(FechaIni)
     Me.DtaFecha.Recordset.Fields("Final") = CStr(FechaFin)
     If iPeriodo = 1 Then
       Me.DtaFecha.Recordset.Fields("Actual") = 1
     Else
       Me.DtaFecha.Recordset.Fields("Actual") = 0
     End If
     Me.DtaFecha.Recordset.Update


End Select
  
 Return
 
EditarPeriodo:

Select Case TipoNomina
 Case "Semanal Sabado"
       Fechas = Format(FechaIni, "YYYY/MM/DD")
        NFecha1 = CDate(FechaIni)
        NFecha2 = FechaFin
        Me.DtaFecha.RecordSource = "SELECT CodTipoNomina, Periodo, año, mes, Inicio, Final, Actual From PeriodoFiscal Where (Inicio = '" & Fechas & "') AND (CodTipoNomina = '" & CodTipoNomina & "')"
        Me.DtaFecha.Refresh
     
        Me.DtaFecha.Recordset.Fields("CodTipoNomina") = CodTipoNomina
        Me.DtaFecha.Recordset.Fields("Inicio") = CStr(FechaIni)
        Me.DtaFecha.Recordset.Fields("Final") = CStr(FechaFin)
        Me.DtaFecha.Recordset.Update
 Case "Semanal Viernes"
       Fechas = Format(FechaIni, "YYYY/MM/DD")
        NFecha1 = CDate(FechaIni)
        NFecha2 = FechaFin
        Me.DtaFecha.RecordSource = "SELECT CodTipoNomina, Periodo, año, mes, Inicio, Final, Actual From PeriodoFiscal Where (Inicio = '" & Fechas & "') AND (CodTipoNomina = '" & CodTipoNomina & "')"
        Me.DtaFecha.Refresh
     
        Me.DtaFecha.Recordset.Fields("CodTipoNomina") = CodTipoNomina
        Me.DtaFecha.Recordset.Fields("Inicio") = CStr(FechaIni)
        Me.DtaFecha.Recordset.Fields("Final") = CStr(FechaFin)
        Me.DtaFecha.Recordset.Update
  
Case "Quincenal"
        NFecha1 = CDate(FechaIni)
        
       Fechas = Format(FechaIni, "YYYY/MM/DD")
        NFecha2 = CDate(FechaFin)
        Me.DtaFecha.RecordSource = "SELECT CodTipoNomina, Periodo, año, mes, Inicio, Final, Actual From PeriodoFiscal WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Inicio = CONVERT(DATETIME, '" & Fechas & "', 102)) "
'        Me.DtaFecha.RecordSource = "SELECT CodTipoNomina, Periodo, año, mes, Inicio, Final, Actual From PeriodoFiscal Where (Inicio = '" & Fechaini & "') AND (CodTipoNomina = '" & CodTipoNomina & "')"
        Me.DtaFecha.Refresh
        
        If Not Me.DtaFecha.Recordset.EOF Then
            Me.DtaFecha.Recordset.Fields("CodTipoNomina") = CodTipoNomina
            Me.DtaFecha.Recordset.Fields("Inicio") = FechaIni
            Me.DtaFecha.Recordset.Fields("Final") = FechaFin
            Me.DtaFecha.Recordset.Update
        End If
        
Case "Mensual"
       Fechas = Format(FechaIni, "YYYY/MM/DD")
        NFecha1 = CDate(FechaIni)
        NFecha2 = CDate(FechaFin)
        Me.DtaFecha.RecordSource = "SELECT CodTipoNomina, Periodo, año, mes, Inicio, Final, Actual From PeriodoFiscal Where (Inicio = '" & Fechas & "') AND (CodTipoNomina = '" & CodTipoNomina & "')"
        Me.DtaFecha.Refresh
        If Not Me.DtaFecha.Recordset.EOF Then
         Me.DtaFecha.Recordset.Fields("Periodo") = iPeriodo
         Me.DtaFecha.Recordset.Fields("CodTipoNomina") = CodTipoNomina
         Me.DtaFecha.Recordset.Fields("Inicio") = FechaIni
         Me.DtaFecha.Recordset.Fields("Final") = FechaFin
         Me.DtaFecha.Recordset.Update
        End If

End Select
  
   Return
 
 
 
Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub cmdSalir_Click()

  Unload Me
  
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub DataCombo1_Change()


End Sub

Private Sub DBTipoNomina_Change()
Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, UltFecha, TipoPago, Moneda, MantValor From TipoNomina WHERE  (Nomina = '" & Me.DbTipoNomina.Text & "')"
Me.DtaConsulta.Refresh
If Me.DtaConsulta.Recordset.EOF Then
  Exit Sub
Else
 TipoNomina = Me.DtaConsulta.Recordset("Periodo")
 CodTipoNomina = Me.DtaConsulta.Recordset("CodTipoNomina")
End If
Semanas
Me.DtaAño.RecordSource = "SELECT CodTipoNomina,Año, Actual, FecIni, FecFin From Año_Fiscal Where (Actual = 1) AND (CodTipoNomina='" & CodTipoNomina & "')"
Me.DtaAño.Refresh
If DtaAño.Recordset.EOF Then
 Me.mskInicio.Value = CDate("01/07/" & Year(Now))
 Me.mskFinal.Value = CDate("30/06/" & Year(Now) + 1)
 Me.TxtAño.Text = Year(Now)

Else
 Me.mskInicio.Value = CDate(Me.DtaAño.Recordset.Fields("FecIni"))
 Me.mskFinal.Value = CDate(Me.DtaAño.Recordset.Fields("FecFin"))
 Me.TxtAño.Text = Year(Me.DtaAño.Recordset("FecIni"))
End If
CmdGenerar_Click
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
Me.CmdGenerar.BackColor = RGB(219, 226, 242)
Me.cmdGuardar.BackColor = RGB(219, 226, 242)
Me.CmdSalir.BackColor = RGB(219, 226, 242)
With Me.DtaAño
  .ConnectionString = Conexion
  
End With

With Me.DtaFecha
  .ConnectionString = Conexion

End With

With Me.DtaConsulta
  .ConnectionString = Conexion

End With

With Me.DtaTipoNomina
  .ConnectionString = Conexion
  .RecordSource = "TipoNomina"
  .Refresh
End With

'Me.DtaAño.Refresh
'Me.DtaFecha.Refresh
'Me.dtaSem.Refresh





End Sub

Public Sub Semanas()

Dim sCad As String
Dim Fecha As Date
Dim FechaFin As Date
Dim iCont As Integer
Dim iValor As Integer
Dim iMes As Integer



sCad = "01/01/"

Me.DtaAño.RecordSource = "SELECT Año, Actual, FecIni, FecFin From Año_Fiscal Where (Actual = 1)"
Me.DtaAño.Refresh

If Not Me.DtaAño.Recordset.EOF Then
   sCad = sCad + CStr(Me.DtaAño.Recordset("Año"))
   iAño = Me.DtaAño.Recordset("Año")
'   Fecha = sCad
   Fecha = Me.mskInicio.Value
   sCad = "31/12/"
   sCad = sCad + CStr(Me.DtaAño.Recordset("Año"))
'   Fechafin = sCad
    FechaFin = Me.mskFinal.Value
End If

iValor = CInt(Mid$(Fecha, 4, 2))

'Me.dtaSem.Recordset.FindFirst "[Año] like " & iAño & ""

'If Me.dtaSem.Recordset.NoMatch Then

  iMes = 1

  Do While Fecha <= FechaFin + 6
   
    If Format(Fecha, "dddd") = "Sábado" And iMes = iValor Then
       iCont = iCont + 1
       Fecha = Fecha + 7
     
    ElseIf iMes <> iValor Then
       GoSub GuardarMes
       iMes = iValor
       iCont = 0
    Else
       Fecha = Fecha + 1
    End If
  
       iValor = CInt(Mid$(Fecha, 4, 2))
  
  Loop

'End If

  Exit Sub


GuardarMes:
         
  Select Case iMes
  
   Case 1:
      sCad = "Enero"
   Case 2:
      sCad = "Febrero"
   Case 3:
      sCad = "Marzo"
   Case 4:
      sCad = "Abril"
   Case 5:
      sCad = "Mayo"
   Case 6:
      sCad = "Junio"
   Case 7:
      sCad = "Julio"
   Case 8:
      sCad = "Agosto"
   Case 9:
      sCad = "Septiembre"
   Case 10:
      sCad = "Octubre"
   Case 11:
      sCad = "Noviembre"
   Case 12:
      sCad = "Diciembre"
      Fecha = Fecha + 7
 End Select
 
 'Me.dtaSem.Recordset.AddNew
    'Me.dtaSem.Recordset.Fields(0) = iAño
    'Me.dtaSem.Recordset.Fields(1) = sCad
    'Me.dtaSem.Recordset.Fields(2) = iCont
 'Me.dtaSem.UpdateRecord
   
 Return
 
 
 
 
End Sub


Private Sub SmartButton1_Click()

End Sub

Private Sub mskFinal_Change()
 Me.TxtAño.Text = Year(Me.mskInicio.Value)
End Sub

Private Sub mskInicio_Change()
Me.TxtAño.Text = Year(Me.mskInicio.Value)
End Sub

Private Sub txtAño_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   Me.CmdGenerar.SetFocus
End If

End Sub

Public Function FechaQuincenal(bBisiesto As Boolean, FechaFin As Date) As Byte
 
Dim bMes As Byte
Dim bDias As Byte

bMes = Mid$(CStr(FechaFin), 4, 2)
bDias = Mid$(CStr(FechaFin), 1, 2)
 
     If (bMes = 1 Or bMes = 3 Or _
            bMes = 5 Or bMes = 7 Or _
            bMes = 8 Or bMes = 10 Or _
            bMes = 12) And bDias = 15 Then
            FechaQuincenal = 16
            
     ElseIf bMes = 1 Or bMes = 3 Or _
            bMes = 5 Or bMes = 7 Or _
            bMes = 8 Or bMes = 10 Or _
            bMes = 12 Then
            
            FechaQuincenal = 15
            
     ElseIf bMes = 2 And bDias = 15 Then
        If bBisiesto Then
           FechaQuincenal = 14
        Else
           FechaQuincenal = 13
        End If
        
     ElseIf bMes = 2 Then
        FechaQuincenal = 15
     
     Else
       FechaQuincenal = 15
     
     End If
                
     

End Function

Private Sub xpcmdbutton1_Click()

End Sub
