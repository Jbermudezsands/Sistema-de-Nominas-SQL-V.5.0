Attribute VB_Name = "Funciones"


Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst _
    As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 _
    As Long, ByVal un2 As Long) As Long

 Public Function Bitacora(FechaHora As Date, Usuario As String, Modulo As String, Accion As String) As Double
  Dim cn As New ADODB.Connection
  Dim rs As New ADODB.Recordset
  
  If RegistrarBitacora = True Then
'      MDIPrimero.AdoConsulta.ConnectionString = Conexion
'      MDIPrimero.AdoConsulta.RecordSource = "Bitacora"
'      MDIPrimero.AdoConsulta.Refresh
      
      If Usuario = "" Then
        Usuario = "Desconocido"
      End If
      
      rs.Open "INSERT INTO Bitacora ([FechaHora],[Usuario],[Modulo],[Accion]) Values ('" & Format(FechaHora, "dd/mm/yyyy HH:mm:ss") & "','" & Usuario & "','" & Modulo & "','" & Accion & "')", Conexion
    
'      MDIPrimero.AdoConsulta.Recordset.AddNew
'      MDIPrimero.AdoConsulta.Recordset("FechaHora") = FechaHora
'      MDIPrimero.AdoConsulta.Recordset("Usuario") = Usuario
'      MDIPrimero.AdoConsulta.Recordset("Modulo") = Modulo
'      MDIPrimero.AdoConsulta.Recordset("Accion") = Accion
'      MDIPrimero.AdoConsulta.Recordset.Update
      
      Bitacora = 1
 End If
End Function
    
    
    
Public Function RestarDiasViaticos(FechaIniNomina As Date, FechaFinNomina As Date, CodigoEmpleado As String) As Double
  Dim FechaIni As Date, FechaFin As Date, Dias As Double, DiasDisfrutados As Double
  
  
  '///////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////////BUSCO SI EXISTEN SOLICITUD DE SUBSIDIOS //////////////////
  '////////////////////////////////////////////////////////////////////////////////////////////////
  MDIPrimero.AdoConsulta.ConnectionString = Conexion
  MDIPrimero.AdoConsulta.RecordSource = "SELECT SolicitudVacaciones.* From SolicitudVacaciones " & _
                                        "WHERE (FechaInicio <= CONVERT(DATETIME, '" & Format(FechaFinNomina, "yyyy-mm-dd") & "', 102)) AND (FechaFin >= CONVERT(DATETIME, '" & Format(FechaIniNomina, "yyyy-mm-dd") & "', 102)) AND (CodigoEmpleado = '" & CodigoEmpleado & "')"
  MDIPrimero.AdoConsulta.Refresh
  Do While Not MDIPrimero.AdoConsulta.Recordset.EOF
    FechaIni = Format(MDIPrimero.AdoConsulta.Recordset("FechaInicio"), "dd/mm/yyyy")
    FechaFin = Format(MDIPrimero.AdoConsulta.Recordset("FechaFin"), "dd/mm/yyyy")
    DiasDisfrutados = MDIPrimero.AdoConsulta.Recordset("DiasDisfrutados")
     
    If DiasDisfrutados >= 1 Then
        If FechaIni >= FechaIniNomina Then
          If FechaFin >= FechaFinNomina Then
            Dias = Dias + DateDiff("d", FechaIni, FechaFinNomina) + 1
          Else
            Dias = Dias + DateDiff("d", FechaIni, FechaFin) + 1
          End If
        Else
          If FechaFin <= FechaFinNomina Then
            Dias = Dias + DateDiff("d", FechaIniNomina, FechaFin) + 1
          End If
        End If
    End If
   
   
  
   MDIPrimero.AdoConsulta.Recordset.MoveNext
  Loop
  
  
  RestarDiasViaticos = Dias


End Function
Public Function ConsecutivoUser(CodUser As String) As Double
 Dim Numero As Double
   MDIPrimero.AdoConsulta.ConnectionString = Conexion
   MDIPrimero.AdoConsulta.RecordSource = "SELECT Userid, REPLACE(STR(Userid), ' ', '0') AS Orden From Userinfo ORDER BY Orden"
   MDIPrimero.AdoConsulta.Refresh
   If MDIPrimero.AdoConsulta.Recordset.EOF Then
    Numero = 1
   Else
     MDIPrimero.AdoConsulta.Recordset.MoveLast
     Numero = MDIPrimero.AdoConsulta.Recordset("Userid")
     Numero = Numero + 1
   End If

ConsecutivoUser = Numero
 

End Function
 Function BuscaTarifaHoraria(CodEmpleado As String) As Double
        Dim Horas As Double, CodTipoNomina As String, TarifaHoraria As Double, Salario As Double
        Dim TipoPago As String, SueldoPeriodo As Double
        
        '////////////////////////////////////////////////////////////////////////
        '//////////////////////////BUSCO LOS DATOS DEL EMPLEADO ///////
        '//////////////////////////////////////////////////////////////////////////
        CodTipoNomina = 0
        MDIPrimero.DtaConsulta.RecordSource = "SELECT  Empleado.* From Empleado WHERE (CodEmpleado = " & CodEmpleado & ") AND (Activo = 1)"
        MDIPrimero.DtaConsulta.Refresh
        If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
          CodTipoNomina = MDIPrimero.DtaConsulta.Recordset("CodTipoNomina")
          TarifaHoraria = MDIPrimero.DtaConsulta.Recordset("TarifaHoraria")
          Salario = MDIPrimero.DtaConsulta.Recordset("SueldoPeriodo")
          SueldoPeriodo = MDIPrimero.DtaConsulta.Recordset("SueldoPeriodo")
        End If
        
        '/////////////////////////////////////////////////////////////////////////////
        '//////////////////////////BUSCO LA CONFIGURACION DEL TIPO DE NOMINA /////////
        '/////////////////////////////////////////////////////////////////////////////
        MDIPrimero.DtaConsulta.RecordSource = "SELECT  * From TipoNomina WHERE (CodTipoNomina = '" & CodTipoNomina & "')"
        MDIPrimero.DtaConsulta.Refresh
        If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
          Horas = MDIPrimero.DtaConsulta.Recordset("Horas")
          TipoPeriodo = MDIPrimero.DtaConsulta.Recordset("Periodo")
          TipoPago = MDIPrimero.DtaConsulta.Recordset("TipoPago")
        End If
        
        
      
         '/////////////////////////////////////////////////////////////////////////////
        '//////////////////////////BUSCO LA CONFIGURACION EN CONTROLES /////////
        '/////////////////////////////////////////////////////////////////////////////
        MDIPrimero.DtaConsulta.RecordSource = "SELECT  DiasMes, DiasSemana, VerificarTasa, SalarioPromedioReal, AntiguedadMenor, CalcularRedondeado From Controles "
        MDIPrimero.DtaConsulta.Refresh
        If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
          DiasMes = MDIPrimero.DtaConsulta.Recordset("DiasMes")
        End If
        
   
        
        
        
        Select Case TipoPeriodo
        
        Case "Semanal Viernes"
            Dim SalarioMes As Double, SalarioDia As Double, SalarioHora As Double
            SalarioMes = Salario * 4.33 '(52 / 12)  Factor se obtiene dividiendo
            SalarioDia = SalarioMes / DiasMes
            SalarioHora = SalarioDia / Horas
            BuscaTarifaHoraria = Format(SalarioHora, "###,##0.000000")
        Case "Semanal Sabado"
            BuscaTarifaHoraria = Format(TarifaHoraria, "###,##0.000000")
        Case "Catorcenal los Viernes"
           If TipoPago = "Salario Fijo" Then
             BuscaTarifaHoraria = Format(SueldoPeriodo / 14 / Horas, "###,##0.0000")
           Else
            BuscaTarifaHoraria = Format(TarifaHoraria, "###,##0.0000")
           End If
        Case "Catorcenal los Sabados"
           If TipoPago = "Salario Fijo" Then
             BuscaTarifaHoraria = Format(SueldoPeriodo / 14 / Horas, "###,##0.0000")
           Else
            BuscaTarifaHoraria = Format(TarifaHoraria, "###,##0.0000")
           End If
        Case "Quincenal"
            BuscaTarifaHoraria = Format((Salario / 15) / Horas, "###,##0.000000")
            'BuscaMontoHora = Format(((Salario) * 2) / (DiasMes * Horas), "###,##0.000000")
'            MontoHora = Format(DtaEmpleados.Recordset("SueldoPeriodo") / ((DiasMes * 8) / 2), "###,##0.000000")
        Case "Mensual"
            BuscaTarifaHoraria = Format(SueldoPeriodo / (DiasMes * Horas), "###,##0.00")
        Case "Trimestral"
            BuscaTarifaHoraria = Format(SueldoPeriodo / (DiasMes * Horas * 3), "###,##0.00")
        Case "Semestral"
            BuscaTarifaHoraria = Format(SueldoPeriodo / (DiasMes * Horas * 6), "###,##0.00")
        End Select


End Function



Public Function Exportar_ADO_Excel(Cadena As String, sql As String, sOutputPathXLS As String) As Boolean
      
    On Error GoTo errSub
      
    Dim cn          As New ADODB.Connection
    Dim rec         As New ADODB.Recordset
    Dim Excel       As Object
    Dim Libro       As Object
    Dim Hoja        As Object
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
      
          
   ' -- Abrir la base
'    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPathDB & ";"
    cn.Open Cadena
          
    ' -- Abrir el Recordset pasándole la cadena sql
    rec.Open sql, cn
      
    ' -- Crear los objetos para utilizar el Excel
    Set Excel = CreateObject("Excel.Application")
    Set Libro = Excel.Workbooks.Add
      
    ' -- Hacer referencia a la hoja
    Set Hoja = Libro.Worksheets(1)
      
    Excel.Visible = True: Excel.UserControl = True
    iCol = rec.Fields.Count
    For iCol = 1 To rec.Fields.Count
        Hoja.Cells(1, iCol).Value = rec.Fields(iCol - 1).Name
    Next
      
    If val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
        Hoja.Cells(2, 1).CopyFromRecordset rec
    Else
  
        arrData = rec.GetRows
  
        iRec = UBound(arrData, 2) + 1
          
        For iCol = 0 To rec.Fields.Count - 1
            For iRow = 0 To iRec - 1
  
                If IsDate(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))
  
                ElseIf IsArray(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = "Array Field"
                End If
            Next iRow
        Next iCol
              
        ' -- Traspasa los datos a la hoja de Excel
        Hoja.Cells(2, 1).Resize(iRec, rec.Fields.Count).Value = GetData(arrData)
    End If
  
    Excel.Selection.CurrentRegion.Columns.AutoFit
'    Excel.Selection.CurrentRegion.Rows.AutoFit
  
    ' -- Cierra el recordset y la base de datos y los objetos ADO
    rec.Close
    cn.Close
      
    Set rec = Nothing
    Set cn = Nothing
    ' -- guardar el libro
    Libro.SaveAs sOutputPathXLS
'    Libro.Close
    ' -- Elimina las referencias Xls
    Set Hoja = Nothing
    Set Libro = Nothing
    Excel.Quit
    Set Excel = Nothing
      
    Exportar_ADO_Excel = True

    Exit Function
errSub:
    MsgBox Err.Description, vbCritical, "Error"
    Exportar_ADO_Excel = False

End Function
  
Public Function GetData(vValue As Variant) As Variant
    Dim X As Long, Y As Long, xMax As Long, yMax As Long, T As Variant
      
    xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
      
    ReDim T(xMax, yMax)
    For X = 0 To xMax
        For Y = 0 To yMax
            T(X, Y) = vValue(Y, X)
        Next Y
    Next X
      
    GetData = T
End Function



Public Function CalcularMontoINSS(CodTipoNomina As String, TotalDevengado As Double, SalarioMensual As Double)
Dim MontoInss As Double, MontoInssPatronal As Double, MontoInssMensual As Double, MontoInssPatronalMensual As Double
Dim MontoInssPatronalAnterior As Double, TasaInss As Double, TasaInssPatronal As Double



'//////////////Verifico si el calculo del Inss es Porcentual//////////////
'CodTipoNomina = FrmCalcularNomina.DtaTipoNomina.Recordset("CodTipoNomina")
FrmCalcularNomina.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, PorcientoInss, TasaInssPatronal, TasaInss, PorcientoIr, TasaIr From TipoNomina WHERE (PorcientoInss = 1) AND (CodTipoNomina = '" & CodTipoNomina & "' )"
FrmCalcularNomina.DtaConsulta.Refresh
If FrmCalcularNomina.DtaConsulta.Recordset.EOF Then
 If FrmCalcularNomina.DtaEmpleados.Recordset("ExentoInss") = 0 Then
     
         FrmCalcularNomina.DtaInss.Refresh
         Do While Not FrmCalcularNomina.DtaInss.Recordset.EOF
          
    
                If FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Semanal Viernes" Then
                   If FrmCalcularNomina.DtaInss.Recordset("desde") < (SalarioMensual) And FrmCalcularNomina.DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      MontoInss = FrmCalcularNomina.DtaInss.Recordset("montolaboral1")
                      MontoInssPatronal = FrmCalcularNomina.DtaInss.Recordset("montopatronal1")
                      Exit Do
                   End If
                   
                ElseIf FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
                   If FrmCalcularNomina.DtaInss.Recordset("desde") < (SalarioMensual) And FrmCalcularNomina.DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      MontoInss = FrmCalcularNomina.DtaInss.Recordset("montolaboral1")
                      MontoInssPatronal = FrmCalcularNomina.DtaInss.Recordset("montopatronal1")
                      Exit Do
                   End If
                ElseIf FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
                
                   If FrmCalcularNomina.DtaInss.Recordset("desde") < (SalarioMensual) And FrmCalcularNomina.DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      If CantSabados = 4 Then
                        If DiaFin < 28 Then
                        MontoInss = (FrmCalcularNomina.DtaInss.Recordset("montolaboral4") / 2)
                        MontoInssPatronal = (FrmCalcularNomina.DtaInss.Recordset("montopatronal4") / 2)
                        Exit Do
                       Else
                        MontoInssMensual = FrmCalcularNomina.DtaInss.Recordset("montolaboral4")
                        MontoInssPatronalMensual = FrmCalcularNomina.DtaInss.Recordset("montopatronal4")
                        MontoInss = MontoInssMensual - MontoInssAnterior
                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
                       End If
                      Else
              '/////////Calcula para Cinco Semanas////////
                     If DiaFin < 28 Then
                        MontoInss = (FrmCalcularNomina.DtaInss.Recordset("montolaboral5") / 2)
                        MontoInssPatronal = (FrmCalcularNomina.DtaInss.Recordset("montopatronal5") / 2)
                        Exit Do
                     Else
                        MontoInssMensual = FrmCalcularNomina.DtaInss.Recordset("montolaboral5")
                        MontoInssPatronalMensual = FrmCalcularNomina.DtaInss.Recordset("montopatronal5")
                        MontoInss = MontoInssMensual - MontoInssAnterior
                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
                     End If
                      
                      End If
                   End If
                ElseIf FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
                
                   If FrmCalcularNomina.DtaInss.Recordset("desde") < (SalarioMensual) And FrmCalcularNomina.DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      If CantSabados = 4 Then
               '/////////////Calcula para cuatro semanas////
                       If DiaFin < 28 Then
                        MontoInss = (FrmCalcularNomina.DtaInss.Recordset("montolaboral4") / 2)
                        MontoInssPatronal = (FrmCalcularNomina.DtaInss.Recordset("montopatronal4") / 2)
                        Exit Do
                       Else
                        MontoInssMensual = FrmCalcularNomina.DtaInss.Recordset("montolaboral4")
                        MontoInssPatronalMensual = FrmCalcularNomina.DtaInss.Recordset("montopatronal4")
                        MontoInss = MontoInssMensual - MontoInssAnterior
                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
                       End If
                      Else
              '/////////Calcula para Cinco Semanas////////
                     If DiaFin < 28 Then
                        MontoInss = (FrmCalcularNomina.DtaInss.Recordset("montolaboral5") / 2)
                        MontoInssPatronal = (FrmCalcularNomina.DtaInss.Recordset("montopatronal5") / 2)
                        Exit Do
                     Else
                        MontoInssMensual = FrmCalcularNomina.DtaInss.Recordset("montolaboral5")
                        MontoInssPatronalMensual = FrmCalcularNomina.DtaInss.Recordset("montopatronal5")
                        MontoInss = MontoInssMensual - MontoInssAnterior
                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
                     End If
                      End If
                   End If
                ElseIf FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
                               
                   If FrmCalcularNomina.DtaInss.Recordset("desde") < (SalarioMensual) And FrmCalcularNomina.DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      If CantSabados = 4 Then
                       '///////Calculo para 4 Semanas///////////
                       If DiaFin < 28 Then
                        MontoInss = (FrmCalcularNomina.DtaInss.Recordset("montolaboral4") / 2)
                        MontoInssPatronal = (FrmCalcularNomina.DtaInss.Recordset("montopatronal4") / 2)
                        Exit Do
                       Else
                         MontoInssMensual = FrmCalcularNomina.DtaInss.Recordset("montolaboral4")
                         MontoInssPatronalMensual = FrmCalcularNomina.DtaInss.Recordset("montopatronal4")
                         MontoInss = MontoInssMensual - MontoInssAnterior
                         MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
                         Exit Do
                       End If
                      Else
                      '///Calculo para 5 Semansas//////////
                       If DiaFin < 28 Then
                        MontoInss = (FrmCalcularNomina.DtaInss.Recordset("montolaboral5") / 2)
                        MontoInssPatronal = (FrmCalcularNomina.DtaInss.Recordset("montopatronal5") / 2)
                        
                        Exit Do
                       Else
                        MontoInssMensual = FrmCalcularNomina.DtaInss.Recordset("montolaboral5")
                        MontoInssPatronalMensual = FrmCalcularNomina.DtaInss.Recordset("montopatronal5")
                        MontoInss = MontoInssMensual - MontoInssAnterior
                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
                       End If
                      End If
                   End If
                
                ElseIf FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Mensual" Then

                   If FrmCalcularNomina.DtaInss.Recordset("desde") < (SalarioMensual) And FrmCalcularNomina.DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      If CantSabados = 4 Then
                        MontoInss = FrmCalcularNomina.DtaInss.Recordset("montolaboral4")
                        MontoInssPatronal = FrmCalcularNomina.DtaInss.Recordset("montopatronal4")
                        Exit Do
                      Else
                        MontoInss = FrmCalcularNomina.DtaInss.Recordset("montolaboral5")
                        MontoInssPatronal = FrmCalcularNomina.DtaInss.Recordset("montopatronal5")
                        Exit Do
                      
                      End If
                   End If
                
                ElseIf FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
                
                   If FrmCalcularNomina.DtaInss.Recordset("desde") < (SalarioMensual) And FrmCalcularNomina.DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      If CantSabados = 4 Then
                        MontoInss = FrmCalcularNomina.DtaInss.Recordset("montolaboral4") * 3
                        MontoInssPatronal = FrmCalcularNomina.DtaInss.Recordset("montopatronal4") * 3
                        Exit Do
                      Else
                        MontoInss = FrmCalcularNomina.DtaInss.Recordset("montolaboral5") * 3
                        MontoInssPatronal = FrmCalcularNomina.DtaInss.Recordset("montopatronal5") * 3
                        Exit Do
                      
                      End If
                   End If
                
                
                
                ElseIf FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
                
                
                   If FrmCalcularNomina.DtaInss.Recordset("desde") < (SalarioMensual) And FrmCalcularNomina.DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      If CantSabados = 4 Then
                        MontoInss = FrmCalcularNomina.DtaInss.Recordset("montolaboral4") * 6
                        MontoInssPatronal = FrmCalcularNomina.DtaInss.Recordset("montopatronal4") * 6
                        Exit Do
                      Else
                        MontoInss = FrmCalcularNomina.DtaInss.Recordset("montolaboral5") * 6
                        MontoInssPatronal = FrmCalcularNomina.DtaInss.Recordset("montopatronal5") * 6
                        Exit Do
                      End If
                   End If
                
                End If
                
         FrmCalcularNomina.DtaInss.Recordset.MoveNext
         Loop
 End If 'del if que pregunta si el empleado es excento de INSS
Else



 If FrmCalcularNomina.DtaEmpleados.Recordset("ExentoInss") = 0 Then
  TasaInss = FrmCalcularNomina.DtaConsulta.Recordset("TasaInss")
  TasaInssPatronal = FrmCalcularNomina.DtaConsulta.Recordset("TasaInssPatronal")
  
  Select Case FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo")
    Case "Quincenal"
        If DiaFin < 28 Then
           If (TotalDevengado) > 88005.78 Then
              MontoInss = (88005.78 * (TasaInss / 100)) / 2
              MontoInssPatronal = (88005.78 * (TasaInssPatronal / 100)) / 2
              MontoInssMensual = MontoInssAnterior + MontoInss
              MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
           Else
             MontoInss = (TotalDevengado) * (TasaInss / 100)
'             MontoInss = (TotalDevengado + MontoVacaciones) * (TasaInss / 100)
             MontoInssPatronal = (TotalDevengado) * (TasaInssPatronal / 100)
             MontoInssMensual = MontoInssAnterior + MontoInss
             MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
           End If
        Else
           MontoInssMensual = ((TotalDevengado) * (TasaInss / 100)) + MontoInssAnterior
           If MontoInssMensual > (88005.78 * (TasaInss / 100)) Then
              MontoInss = (88005.78 * (TasaInss / 100)) - MontoInssAnterior
              MontoInssMensual = MontoInss + MontoInssAnterior
              MontoInssPatronal = (88005.78 * (TasaInssPatronal / 100)) - MontoInssPatronalAnterior
              MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
           Else
             MontoInss = (TotalDevengado) * (TasaInss / 100)
'             MontoInss = (TotalDevengado + MontoVacaciones) * (TasaInss / 100)
             MontoInssPatronal = (TotalDevengado) * (TasaInssPatronal / 100)
             MontoInssMensual = MontoInssAnterior + MontoInss
             MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
           End If
        End If
    Case "Mensual"
           If (TotalDevengado) > 88005.78 Then
              MontoInss = 88005.78 * 0.0625
              MontoInssPatronal = 88005.78 * 0.019
              MontoInssMensual = MontoInssAnterior + MontoInss
              MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
           Else
             MontoInss = (TotalDevengado) * (TasaInss / 100)
'             MontoInss = (TotalDevengado + MontoVacaciones) * (TasaInss / 100)
             MontoInssPatronal = (TotalDevengado) * (TasaInssPatronal / 100)
             MontoInssMensual = MontoInssAnterior + MontoInss
             MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
           End If

    Case Else
             '+ MontoOtrosIngresos
             MontoInss = (TotalDevengado) * (TasaInss / 100)
             MontoInssPatronal = (TotalDevengado) * (TasaInssPatronal / 100)
             MontoInssMensual = MontoInssAnterior + MontoInss
             MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
  
  End Select
  
 Else
    MontoInss = 0
    MontoInssPatronal = 0
    MontoInssMensual = 0
    MontoInssPatronalMensual = 0
 End If
 
 MontoInssRegistros(0) = MontoInss
 MontoInssRegistros(1) = MontoInssPatronal
 MontoInssRegistros(2) = MontoInssMensual
 MontoInssRegistros(3) = MontoInssPatronalMensual

End If 'del if del porciento
End Function

    
    
    
Public Function BuscaTasaCambio(Fecha As Date) As Double
        MDIPrimero.DtaConsulta.RecordSource = "SELECT FechaDia, MontoDia From CambioMoneda WHERE (FechaDia = '" & Format(Fecha, "yyyymmdd") & "')"
        MDIPrimero.DtaConsulta.Refresh
        If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
           BuscaTasaCambio = MDIPrimero.DtaConsulta.Recordset("MontoDia")
        Else
           BuscaTasaCambio = 1
        End If

End Function
Public Function MontoDeduccion(NumeroNomina As Double, CodTipoNomina As String, CodigoEmpleado As String) As Double
  Dim SqlString As String
  
  SqlString = "SELECT SUM(DetalleDeduccion.Valor) AS Valor, TipoDeduccion.CodTipoDeduccion, TipoDeduccion.Deduccion, DetalleDeduccion.NumNomina, Empleado.CodEmpleado1 FROM DetalleDeduccion INNER JOIN Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion INNER JOIN TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion INNER JOIN Empleado ON Deduccion.CodEmpleado = Empleado.CodEmpleado GROUP BY TipoDeduccion.CodTipoDeduccion, TipoDeduccion.Deduccion, DetalleDeduccion.NumNomina, Empleado.CodEmpleado1 HAVING (DetalleDeduccion.NumNomina = " & NumeroNomina & ") AND (TipoDeduccion.CodTipoDeduccion = '" & CodTipoNomina & "') AND (Empleado.CodEmpleado1 = '" & CodigoEmpleado & "')"
  MDIPrimero.DtaConsulta.RecordSource = SqlString
  MDIPrimero.DtaConsulta.Refresh
  If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
    MontoDeduccion = MDIPrimero.DtaConsulta.Recordset("Valor")
  End If

End Function

Public Function Redondear(Valor As Double) As Double
  Dim ValReturn As Double
      SqlString = "Select CalcularRedondeado from Controles"
  MDIPrimero.DtaConsulta.RecordSource = SqlString
  MDIPrimero.DtaConsulta.Refresh
  If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
    If MDIPrimero.DtaConsulta.Recordset("CalcularRedondeado") = True Then
     Dim CadenaValor As String
        CadenaValor = CStr(Valor)
        CadenaValor = Valor - CInt(Valor)
        
        If CDbl(CadenaValor) > 0.5 Then
            Redondear = CInt(Valor)
        Else
            Redondear = CInt(Valor) + 1
        End If
        
    Else
        
        Redondear = Valor
        
    End If
  End If
End Function


Public Function MontoDeduccionTotal(NumeroNomina As Double, CodTipoNomina As String) As Double
  Dim SqlString As String
  
'  SqlString = "SELECT SUM(DetalleDeduccion.Valor) AS Valor, TipoDeduccion.CodTipoDeduccion, TipoDeduccion.Deduccion, DetalleDeduccion.NumNomina, Empleado.CodEmpleado1 FROM DetalleDeduccion INNER JOIN Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion INNER JOIN TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion INNER JOIN Empleado ON Deduccion.CodEmpleado = Empleado.CodEmpleado GROUP BY TipoDeduccion.CodTipoDeduccion, TipoDeduccion.Deduccion, DetalleDeduccion.NumNomina, Empleado.CodEmpleado1 HAVING (DetalleDeduccion.NumNomina = " & NumeroNomina & ") AND (TipoDeduccion.CodTipoDeduccion = '" & CodTipoNomina & "') "
  SqlString = "SELECT SUM(DetalleDeduccion.Valor) AS Valor, TipoDeduccion.CodTipoDeduccion, TipoDeduccion.Deduccion, DetalleDeduccion.NumNomina FROM DetalleDeduccion INNER JOIN Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion INNER JOIN TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion INNER JOIN Empleado ON Deduccion.CodEmpleado = Empleado.CodEmpleado GROUP BY TipoDeduccion.CodTipoDeduccion, TipoDeduccion.Deduccion, DetalleDeduccion.NumNomina HAVING  (TipoDeduccion.CodTipoDeduccion = '" & CodTipoNomina & "') AND (DetalleDeduccion.NumNomina = " & NumeroNomina & ") "
  MDIPrimero.DtaConsulta.RecordSource = SqlString
  MDIPrimero.DtaConsulta.Refresh
  If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
    MontoDeduccionTotal = MDIPrimero.DtaConsulta.Recordset("Valor")
  End If

End Function
Public Function MontoSalarioPorciento(CodigoEmpleado As String) As Double
 Dim SalarioPorciento As Double, SueldoPeriodo As Double
  
             'PORCIENTO DEL SALARIO
             MDIPrimero.DtaConsulta.RecordSource = "SELECT  * From Empleado Where (CodEmpleado = " & CodigoEmpleado & ")"
             MDIPrimero.DtaConsulta.Refresh
       If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
              If Not IsNull(MDIPrimero.DtaConsulta.Recordset("salporcentaje")) Then
                 SalarioPorciento = MDIPrimero.DtaConsulta.Recordset("salporcentaje")
              Else
                 SalarioPorciento = 0
              End If
              
              If Not IsNull(MDIPrimero.DtaConsulta.Recordset("SueldoPeriodo")) Then
                 SueldoPeriodo = MDIPrimero.DtaConsulta.Recordset("SueldoPeriodo")
              Else
                 SueldoPeriodo = 0
              End If
              
           MontoSalarioPorciento = SueldoPeriodo * (SalarioPorciento / 100)
              
              
       End If



End Function


Public Function BuscaCodigoInterno(CodigoEmpleado As String) As Double
  
  Dim SqlString As String

  SqlString = "SELECT  * From Empleado WHERE (CodEmpleado1 = '" & CodigoEmpleado & "')"
  MDIPrimero.DtaConsulta.RecordSource = SqlString
  MDIPrimero.DtaConsulta.Refresh
  If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
    BuscaCodigoInterno = MDIPrimero.DtaConsulta.Recordset("CodEmpleado")
  Else
   
   BuscaCodigoInterno = -1
  End If
  
End Function

Public Function GrabaDescuentoDias(FechaDescuento As Date, CodigoEmpleado As String, TipoDescuento As String, NumeroSolicitud As String, CantDias As Double) As Double
   
   Dim SqlString As String, Fecha As String, Color As String
   Color = BuscaColor(TipoDescuento)
'  Fecha = Format(FechaDescuento, "yyyy-mm-dd")
  SqlString = "SELECT  * From DescuentoDiasVacaciones WHERE (FechaDescuento = '" & Format(FechaDescuento, "yyyy") & Format(FechaDescuento, "MM") & Format(FechaDescuento, "dd") & "') AND (TipoDescuento = '" & TipoDescuento & "') AND (CodigoEmpleado = '" & CodigoEmpleado & "')"
  MDIPrimero.DtaConsulta.RecordSource = SqlString
  MDIPrimero.DtaConsulta.Refresh
  If MDIPrimero.DtaConsulta.Recordset.EOF Then
    MDIPrimero.DtaConsulta.Recordset.AddNew
    MDIPrimero.DtaConsulta.Recordset("FechaDescuento") = FechaDescuento
    MDIPrimero.DtaConsulta.Recordset("CodigoEmpleado") = CodigoEmpleado
    MDIPrimero.DtaConsulta.Recordset("TipoDescuento") = TipoDescuento
    MDIPrimero.DtaConsulta.Recordset("NumeroSolicitud") = NumeroSolicitud
    MDIPrimero.DtaConsulta.Recordset("CantDias") = CantDias
    MDIPrimero.DtaConsulta.Recordset("Color") = Color
    MDIPrimero.DtaConsulta.Recordset.Update
  Else
    MDIPrimero.DtaConsulta.Recordset("CantDias") = CantDias
    MDIPrimero.DtaConsulta.Recordset.Update
  End If

GrabaDescuentoDias = 1

End Function





Public Function BuscaColor(TipoVacaciones As String) As String
  Select Case TipoVacaciones
    Case "Vacaciones Pagadas": BuscaColor = "&H0000C000&"
    Case "Vacaciones": BuscaColor = "&H00FFC0C0&"
    Case "Subsidio": BuscaColor = "&H0080C0FF&"
    Case "Ausente": BuscaColor = "&H00FF80FF&"
    Case "Feriado": BuscaColor = "&H008888FB&"
    Case "Vacaciones Programadas": BuscaColor = "&H0080FFFF&"
    Case "Permiso Programado": BuscaColor = "&H00FFFF00&"
  End Select

End Function



Public Function ExisteEmpleado(CodigoEmpleado As String) As Boolean
  Dim SqlString As String
   SqlString = "SELECT  * From Empleado WHERE (CodEmpleado1 = '" & CodigoEmpleado & "') AND (Activo = 1)"
  MDIPrimero.DtaConsulta.RecordSource = SqlString
  MDIPrimero.DtaConsulta.Refresh
  If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
    ExisteEmpleado = True
  Else
   ExisteEmpleado = False
  End If


End Function
Public Function ExisteEmpleado2(CodigoEmpleado As String) As Boolean
  Dim SqlString As String
   SqlString = "SELECT  * From Empleado WHERE (CodEmpleado= '" & CodigoEmpleado & "') AND (Activo = 1)"
  MDIPrimero.DtaConsulta.RecordSource = SqlString
  MDIPrimero.DtaConsulta.Refresh
  If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
    ExisteEmpleado2 = True
  Else
   ExisteEmpleado2 = False
  End If

End Function

Public Function ConsecutivoSolicitud() As Double
 Dim SqlString As String, NumeroSolicitud As String
 
  ConsecutivoSolicitud = 0
  SqlString = "SELECT  * From Consecutivos"
  MDIPrimero.DtaConsulta.RecordSource = SqlString
  MDIPrimero.DtaConsulta.Refresh
  If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
    If Not IsNull(MDIPrimero.DtaConsulta.Recordset("Solicitud")) Then
        ConsecutivoSolicitud = MDIPrimero.DtaConsulta.Recordset("Solicitud") + 1
        Else
        ConsecutivoSolicitud = 1
    End If
    
  Else
    ConsecutivoSolicitud = 0
  End If
  
  
  '////////////////////BUSCO SI EXISTE EN LOS REGISTROS DE SOLICITUD /////////////////
   MDIPrimero.AdoConsulta.ConnectionString = Conexion
   MDIPrimero.AdoConsulta.RecordSource = "SELECT  * From SolicitudVacaciones ORDER BY NumeroSolicitud DESC"
   MDIPrimero.AdoConsulta.Refresh
   If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
     ConsecutivoSolicitud = MDIPrimero.AdoConsulta.Recordset("NumeroSolicitud") + 1
   Else
     ConsecutivoSolicitud = 1
   End If
 

End Function


Public Function BuscaUltimaSemana(CantSabados As Double, NumeroNomina As Double, Mes As String, Año As Double) As Boolean
  Dim SqlString As String, i As Double
  
  SqlString = "SELECT * From Fecha_Planilla WHERE (año = " & Año & ") AND (mes = '" & Mes & "')"
  FrmCalcularNomina.DtaConsulta.RecordSource = SqlString
  FrmCalcularNomina.DtaConsulta.Refresh
  i = 1
  BuscaUltimaSemana = False
  Do While Not FrmCalcularNomina.DtaConsulta.Recordset.EOF
    If FrmCalcularNomina.DtaConsulta.Recordset("NumNomina") = NumeroNomina Then
       If CantSabados = i Then
         BuscaUltimaSemana = True
       End If
    End If
    
    i = i + 1
    FrmCalcularNomina.DtaConsulta.Recordset.MoveNext
  Loop



End Function


Public Function DiasVacaDescuentos(CodigoEmpleado As String, Fecha1 As Date, Fecha2 As Date, TipoVacaciones) As Double
  Dim SqlString As String, FechaInicio As String, FechaFin As String
  
  FechaInicio = Format(Fecha1, "yyyy-mm-dd")
  FechaFin = Format(Fecha2, "yyyy-mm-dd")
  
  SqlString = "SELECT SUM(CantDias) AS CantDias From DescuentoDiasVacaciones WHERE  (TipoDescuento = '" & TipoVacaciones & "') AND (FechaDescuento BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME,'" & FechaFin & "', 102)) AND (CodigoEmpleado = '" & CodigoEmpleado & "')"
  MDIPrimero.DtaConsulta.RecordSource = SqlString
  MDIPrimero.DtaConsulta.Refresh
  If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
   If Not IsNull(MDIPrimero.DtaConsulta.Recordset("CantDias")) Then
     DiasVacaDescuentos = MDIPrimero.DtaConsulta.Recordset("CantDias")
   Else
     DiasVacaDescuentos = 0
   End If
  Else
   DiasVacaDescuentos = 0
  End If

End Function
Public Function DiasVacaDesAcumulados(CodigoEmpleado As String, Fecha1 As Date) As Double
  Dim SqlString As String, FechaInicio As String, FechaFin As String
  
  FechaInicio = Format(Fecha1, "yyyy-mm-dd")
  FechaFin = Format(Fecha2, "yyyy-mm-dd")
  
  SqlString = "SELECT  SUM(CantDias) AS CantDias From DescuentoDiasVacaciones WHERE (FechaDescuento <= CONVERT(DATETIME, '" & FechaInicio & "', 102)) AND (CodigoEmpleado = '" & CodigoEmpleado & "') AND (TipoDescuento <> 'Permiso Programado' AND TipoDescuento <> 'Vacaciones Programadas'AND TipoDescuento <> 'Subsidio')"
  MDIPrimero.DtaConsulta.RecordSource = SqlString
  MDIPrimero.DtaConsulta.Refresh
  If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
   If Not IsNull(MDIPrimero.DtaConsulta.Recordset("CantDias")) Then
    DiasVacaDesAcumulados = MDIPrimero.DtaConsulta.Recordset("CantDias")
   Else
    DiasVacaDesAcumulados = 0
   End If
  Else
   DiasVacaDesAcumulados = 0
  End If

End Function

Public Function DiasDifrutados(CodigoEmpleado As String, Fecha1 As Date, Fecha2 As Date) As Double
  Dim SqlString As String, FechaInicio As String, FechaFin As String
  
  FechaInicio = Format(Fecha1, "yyyy-mm-dd")
  FechaFin = Format(Fecha2, "yyyy-mm-dd")
  
  SqlString = "SELECT SUM(CantDias) AS CantDias From DescuentoDiasVacaciones WHERE  (FechaDescuento BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (CodigoEmpleado = '" & CodigoEmpleado & "') AND (TipoDescuento <> 'Permiso Programado') AND (TipoDescuento <> 'Vacaciones Programadas') AND (TipoDescuento <> 'Subsidio') "
  MDIPrimero.DtaConsulta.RecordSource = SqlString
  MDIPrimero.DtaConsulta.Refresh
  If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
   If Not IsNull(MDIPrimero.DtaConsulta.Recordset("CantDias")) Then
    DiasDifrutados = MDIPrimero.DtaConsulta.Recordset("CantDias")
   Else
    DiasDifrutados = 0
   End If
  Else
   DiasDifrutados = 0
  End If

End Function

Public Function CalcularQuincenaletras(CodTipoNomina As String, NumeroNomina As Double) As String
 Dim SqlString As String, Mes As String, Año As String, Periodo As Double
 '////////////////////////////////////BUSQUE EL MES //////////////////////////////////////
 SqlString = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (NumNomina = " & NumeroNomina & ")"
 MDIPrimero.DtaConsulta.RecordSource = SqlString
 MDIPrimero.DtaConsulta.Refresh
 If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
   Mes = MDIPrimero.DtaConsulta.Recordset("mes")
   Año = MDIPrimero.DtaConsulta.Recordset("año")
   Periodo = MDIPrimero.DtaConsulta.Recordset("Periodo")
 End If
 

 
 
End Function
Public Function BuscaCodigo(Descripcion As String, Tabla As String, Campo As String, CampoWhere As String) As String
 Dim SqlString As String
 
 SqlString = "SELECT  * From " & Tabla & " WHERE (" & CampoWhere & "  = '" & Descripcion & "')"
 MDIPrimero.DtaConsulta.RecordSource = SqlString
 MDIPrimero.DtaConsulta.Refresh
 If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
  BuscaCodigo = MDIPrimero.DtaConsulta.Recordset(Campo)
 Else
  BuscaCodigo = "00"
  UltimoCodigo = ""

   '/////////////////BUSCO EL ULTIMO CODIGO DEPARTAMENTO ////////////
   MDIPrimero.DtaConsulta.RecordSource = "SELECT * From " & Tabla & " ORDER BY " & Campo & " DESC"
   MDIPrimero.DtaConsulta.Refresh
   If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
     UltimoCodigo = MDIPrimero.DtaConsulta.Recordset(Campo)
     UltimoCodigo = Format(CDbl(UltimoCodigo) + 1, "00#")
   Else
     UltimoCodigo = "001"
   End If
 End If



End Function





Public Function PeriodoNominaLetras(CodTipoNomina As String, Mes As String, Año As Double, FechaInicio As Date, FechaFin As Date) As String
 Dim SqlSring As String, Mes2 As String, i As Double
 
 Mes2 = Format(Mes, "0#")
 SqlString = "SELECT  * From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (mes = '" & Mes2 & "') AND (año = " & Año & ")"
 MDIPrimero.DtaConsulta.RecordSource = SqlString
 MDIPrimero.DtaConsulta.Refresh
 i = 1
 Do While Not MDIPrimero.DtaConsulta.Recordset.EOF
   If FechaInicio = MDIPrimero.DtaConsulta.Recordset("Inicio") Then
     If FechaFin = MDIPrimero.DtaConsulta.Recordset("Final") Then
       Exit Do
     End If
   End If
   
   i = i + 1
   MDIPrimero.DtaConsulta.Recordset.MoveNext
 Loop
 
 Select Case i
   Case 1: PeriodoNominaLetras = "Primera"
   Case 2: PeriodoNominaLetras = "Segunda"
   Case 3: PeriodoNominaLetras = "Tercera"
   Case 4: PeriodoNominaLetras = "Cuarta"
   Case 5: PeriodoNominaLetras = "Quinta"
 End Select
 
End Function

Public Function CalcularDiasVaca(FechaInicio As Date, FechaFin As Date)
  Dim Dias As Double, Mes As Double, FechaFin2 As Date, FechaInicio2 As Date, Dias2 As Double, Dias1 As Double, FechaInicio1 As Date, FechaFin1 As Date
  Dim Fecha As String
  '/////////////////////////////////////////////////////////BUSCO LA CONFIGURACION DE CALCULO ///////////////
   MDIPrimero.DtaControles.Refresh
   If MDIPrimero.DtaControles.Recordset("DiasMes") = 30 Then
     '//////////////////////BUSCO LOS DIAS DEL PRIMER MES////////////////////////////////
      If Day(FechaInicio) > 1 Then
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////SI LA FECHA ES MAYOR, SIGINIFICA QUE TENGO QUE CALCULAR LA FRACCION//////////
        '//////////////////////////////////////////////////////////////////////////////////////////////////////
''        FechaInicio1 = DateSerial(Year(FechaInicio), Month(FechaInicio), 1)
'        FechaFin1 = DateSerial(Year(FechaInicio), Month(FechaInicio) + 1, 0)
         'FechaFin1 = 30 & "/" & Month(FechaInicio) & "/" & Year(FechaInicio)
        dechafin1 = DateSerial(Year(FechaInicio), Month(FechaInicio) + 1, -1) + 1
         Dias1 = DateDiff("d", CDbl(FechaInicio), CDbl(FechaFin)) + 1
      Else
        FechaInicio1 = DateSerial(Year(FechaInicio), Month(FechaInicio), 1)
        FechaFin1 = DateSerial(Year(FechaInicio), Month(FechaInicio) + 1, 0)
        Dias1 = 30
      End If
      
     '///////////////////BUSCO LA FECHA DEL MES ANTERIOR //////////////////
     FechaFin2 = DateSerial(Year(FechaFin), Month(FechaFin), 0)
     If FechaFin2 > (FechaFin1 + 1) Then '//////VALIDO SI LA FECHA FINAL YA CALCULADA ES MENOR QUE LA FECHA FIN //////
        Mes = DateDiff("m", CDbl(FechaFin1 + 1), CDbl(FechaFin2)) + 1

            FechaInicio2 = DateSerial(Year(FechaFin), Month(FechaFin), 1)
            Dias2 = DateDiff("d", CDbl(FechaInicio2), CDbl(FechaFin)) + 1
            If Dias2 > 30 Then
              Dias2 = 30
            End If
            '///////////VERFICIO SI LA FECHA FIN ES FEBRERO ///////////////////////////////////
            If Month(FechaFin) = 2 Then
              If Dias2 >= 28 Then
                Dias2 = 30
              End If
            End If
            
            Dias = Dias1 + Mes * 30 + Dias2
        
     Else
        If FechaFin > (FechaFin1 + 1) Then '///////////////SI EL RANGO ES PARA DOS MESES ///////////////
          Mes = DateDiff("m", CDbl(FechaFin1 + 1), CDbl(FechaFin)) + 1
          Dias2 = DateDiff("d", CDbl(FechaFin1 + 1), CDbl(FechaFin)) + 1
            If Month(FechaFin) = 2 Then
              If Dias2 >= 28 Then
                Dias2 = 30
              End If
            Else '///////////////////SI NO ES FEBRERO, VALIDO EL MES ////////////////
              If Dias2 >= 30 Then
                Dias2 = 30
              End If
            End If
          Dias = Dias1 + Dias2
          
        Else
          Dias = DateDiff("d", CDbl(FechaInicio), CDbl(FechaFin))
        End If
        
     End If
   
   Else
     Dias = DateDiff("d", CDbl(FechaInicio), CDbl(FechaFin))
   End If
  
  CalcularDiasVaca = Dias
    
End Function

Public Function CalculoDiasVacaciones(CodEmpleado As String, FechaFin As Date) As Double
    
Dim fs As Boolean
fs = False
'//////////////////////////////// Saco datos generales del empleado ///////////////////////
MDIPrimero.DtaConsulta.RecordSource = "SELECT     TOP (1) Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombre, Historico.FechaContratoVac, Empleado.CodEmpleado, DATEADD(month,    (YEAR(Historico.FechaContratoVac) - 1900) * 12 + MONTH(Historico.FechaContratoVac), - 1) AS UdMes  FROM         Empleado INNER JOIN   Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (Empleado.CodEmpleado1 = '" & CodEmpleado & "') and Empleado.Activo = 'True'"
MDIPrimero.DtaConsulta.Refresh

'///////// Inicializo parametros generales ////////////
Dim VacacionesAcumuladas, VacacionesSolicitadas, SaldoActual, tempVacacionSolicitada, tempVacacionAcumulada As Double
Dim NombreCompleto As String
    NombreCompleto = MDIPrimero.DtaConsulta.Recordset("Nombre")
Dim TotalAcumuladas As Double, TotalSolicitadas As Double
Dim CodEmpleado1A As String
CodEmpleado1A = MDIPrimero.DtaConsulta.Recordset("CodEmpleado")

Dim Inicio As Date
Dim tempInicio As Date
Dim Fin As Date, tempFin As Date

Inicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
Fin = FechaFin

'/////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////


tempInicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")

tempFin = MDIPrimero.DtaConsulta.Recordset("udMes")


'/////////////////////////////// Inicializo el reporte //////////////////////////////////

         SaldoActual = 0
         tempVacacionSolicitada = 0
         tempVacacionAcumulada = 0
         
         
         Dim Tipo As String
         MDIPrimero.DtaControles.Refresh
         Tipo = MDIPrimero.DtaControles.Recordset("DiasMes")
         
    Do While (Inicio < Fin)
     
          If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
          End If
          

            
            Dim DiasMes As Double
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
        If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
               
               
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 3) * 0.0833, "####0.00")   '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 2) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
                
               'if tipo = 31
               Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 4) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 3) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
               End If
               
                
            Else
                If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00") '/ 12
                    Else
                        tempVacacionAcumulada = 2.58
                    End If
                End If
            End If
          
             
             
             
             '////////     /////////     ////////        /////////       /////////       ////////        //////
             '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
             '/////////////////////////////////////////////////////////////////////////////////////////////////
           If DiasMes = 31 Then
                 MDIPrimero.DtaConsulta.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'  and  (CodigoEmpleado = '" & CodEmpleado & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
           Else
                 MDIPrimero.DtaConsulta.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado' and  (CodigoEmpleado = '" & CodEmpleado & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
           End If
           
             MDIPrimero.DtaConsulta.Refresh
             If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                If Not IsNull(MDIPrimero.DtaConsulta.Recordset("VacacionesSolicitadas")) Then
                    tempVacacionSolicitada = MDIPrimero.DtaConsulta.Recordset("VacacionesSolicitadas")
                Else
                    tempVacacionSolicitada = 0
                End If
             Else
                 tempVacacionSolicitada = 0
             End If
             
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
             

            TotalAcumuladas = TotalAcumuladas + tempVacacionAcumulada
            TotalSolicitadas = TotalSolicitadas + tempVacacionSolicitada


             If DateAdd("d", 2, tempFin) >= Fin Then
                 tempFin = DateAdd("m", -2, tempFin)
                 tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                 tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                 'ponerle temp inicio para que los dias  no varien
                 Inicio = Fin
             Else
                tempFin = DateAdd("d", 2, tempFin)
                tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                'ponerle temp inicio para que los dias  no varien
                Inicio = tempInicio
             End If
             

            
           ' ////////////////
            
            
        If DateSerial(Year(tempFin), Month(tempFin) + 1, 0) >= Fin Then
                tempFin = Fin
    
          If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
          End If
            
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
            If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
                 
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00") '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00") '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
                
               'if tipo = 31
               Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
               End If
            Else
                If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.58
                    End If
                End If
            End If
          
            
             
             '////////     /////////     ////////        /////////       /////////       ////////        //////
             '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
             '/////////////////////////////////////////////////////////////////////////////////////////////////
           If DiasMes = 31 Then
               MDIPrimero.DtaConsulta.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM SolicitudVacaciones WHERE  not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'  and    (CodigoEmpleado = '" & CodEmpleado & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
           Else
               MDIPrimero.DtaConsulta.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM SolicitudVacaciones WHERE  not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'  and    (CodigoEmpleado = '" & CodEmpleado & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
           End If
           
             MDIPrimero.DtaConsulta.Refresh
             If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                If Not IsNull(MDIPrimero.DtaConsulta.Recordset("VacacionesSolicitadas")) Then
                    tempVacacionSolicitada = MDIPrimero.DtaConsulta.Recordset("VacacionesSolicitadas")
                Else
                    tempVacacionSolicitada = 0
                End If
             Else
                 tempVacacionSolicitada = 0
             End If
                    
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
             
            TotalAcumuladas = TotalAcumuladas + tempVacacionAcumulada
            TotalSolicitadas = TotalSolicitadas + tempVacacionSolicitada
             
             tempFin = DateAdd("d", 2, tempFin)
             
             tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1) ' Inicio
             tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin
                         'ponerle temp inicio para que los dias  no varien
            Inicio = tempInicio
                    Inicio = DateAdd("m", 1, Inicio)
            End If
            

    Loop
    
    
            CalculoDiasVacaciones = SaldoActual
End Function


Public Function CalculoDiasVacaSFicha(CodEmpleado As String, FechaFin As Date) As Double
    
Dim fs As Boolean
fs = False
'//////////////////////////////// Saco datos generales del empleado ///////////////////////
MDIPrimero.DtaConsulta.RecordSource = "SELECT     TOP (1) Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombre, Historico.FechaContratoVac, Empleado.CodEmpleado, DATEADD(month,    (YEAR(Historico.FechaContratoVac) - 1900) * 12 + MONTH(Historico.FechaContratoVac), - 1) AS UdMes  FROM         Empleado INNER JOIN   Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (Empleado.CodEmpleado1 = '" & CodEmpleado & "') and Empleado.Activo = 'True'"
MDIPrimero.DtaConsulta.Refresh

'///////// Inicializo parametros generales ////////////
Dim VacacionesAcumuladas, VacacionesSolicitadas, SaldoActual, tempVacacionSolicitada, tempVacacionAcumulada As Double
Dim NombreCompleto As String
    NombreCompleto = MDIPrimero.DtaConsulta.Recordset("Nombre")
Dim TotalAcumuladas As Double, TotalSolicitadas As Double
Dim CodEmpleado1A As String
CodEmpleado1A = MDIPrimero.DtaConsulta.Recordset("CodEmpleado")

Dim Inicio As Date
Dim tempInicio As Date
Dim Fin As Date, tempFin As Date

Inicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
Fin = FechaFin

'/////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////


tempInicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")

tempFin = MDIPrimero.DtaConsulta.Recordset("udMes")


'/////////////////////////////// Inicializo el reporte //////////////////////////////////

         SaldoActual = 0
         tempVacacionSolicitada = 0
         tempVacacionAcumulada = 0
         
         
         Dim Tipo As String
         MDIPrimero.DtaControles.Refresh
         Tipo = MDIPrimero.DtaControles.Recordset("DiasMes")
         
    Do While (Inicio < Fin)
     
          If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
          End If
          

            
            Dim DiasMes As Double
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
        If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
               
               
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 3) * 0.0833, "####0.00")   '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 2) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
                
               'if tipo = 31
               Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 4) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 3) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
               End If
               
                
            Else
                If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00") '/ 12
                    Else
                        tempVacacionAcumulada = 2.58
                    End If
                End If
            End If
          
             
      tempVacacionSolicitada = 0
             
             '////////     /////////     ////////        /////////       /////////       ////////        //////
             '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
             '/////////////////////////////////////////////////////////////////////////////////////////////////
           If DiasMes = 31 Then
                 MDIPrimero.DtaConsulta.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'  and  (CodigoEmpleado = '" & CodEmpleado & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
           Else
                 MDIPrimero.DtaConsulta.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado' and  (CodigoEmpleado = '" & CodEmpleado & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
           End If
           
             MDIPrimero.DtaConsulta.Refresh
             If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                If Not IsNull(MDIPrimero.DtaConsulta.Recordset("VacacionesSolicitadas")) Then
'                    tempVacacionSolicitada = MDIPrimero.DtaConsulta.Recordset("VacacionesSolicitadas")
                Else
                    tempVacacionSolicitada = 0
                End If
             Else
                 tempVacacionSolicitada = 0
             End If
             
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
             

            TotalAcumuladas = TotalAcumuladas + tempVacacionAcumulada
             TotalSolicitadas = TotalSolicitadas + tempVacacionSolicitada


             If DateAdd("d", 2, tempFin) >= Fin Then
                 tempFin = DateAdd("m", -2, tempFin)
                 tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                 tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                 'ponerle temp inicio para que los dias  no varien
                 Inicio = Fin
             Else
                tempFin = DateAdd("d", 2, tempFin)
                tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                'ponerle temp inicio para que los dias  no varien
                Inicio = tempInicio
             End If
             

            
           ' ////////////////
            
            
        If DateSerial(Year(tempFin), Month(tempFin) + 1, 0) >= Fin Then
                tempFin = Fin
    
          If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
          End If
            
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
            If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
                 
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00") '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00") '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
                
               'if tipo = 31
               Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
               End If
            Else
                If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.58
                    End If
                End If
            End If
          
             tempVacacionSolicitada = 0
             
             '////////     /////////     ////////        /////////       /////////       ////////        //////
             '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
             '/////////////////////////////////////////////////////////////////////////////////////////////////
           If DiasMes = 31 Then
               MDIPrimero.DtaConsulta.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM SolicitudVacaciones WHERE  not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'  and    (CodigoEmpleado = '" & CodEmpleado & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
           Else
               MDIPrimero.DtaConsulta.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM SolicitudVacaciones WHERE  not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'  and    (CodigoEmpleado = '" & CodEmpleado & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
           End If
           
             MDIPrimero.DtaConsulta.Refresh
             If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                If Not IsNull(MDIPrimero.DtaConsulta.Recordset("VacacionesSolicitadas")) Then
'                    tempVacacionSolicitada = MDIPrimero.DtaConsulta.Recordset("VacacionesSolicitadas")
                Else
                    tempVacacionSolicitada = 0
                End If
             Else
                 tempVacacionSolicitada = 0
             End If
                    
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
             
            TotalAcumuladas = TotalAcumuladas + tempVacacionAcumulada
            TotalSolicitadas = TotalSolicitadas + tempVacacionSolicitada
             
             tempFin = DateAdd("d", 2, tempFin)
             
             tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1) ' Inicio
             tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin
                         'ponerle temp inicio para que los dias  no varien
            Inicio = tempInicio
                    Inicio = DateAdd("m", 1, Inicio)
            End If
            

    Loop
    
    
            CalculoDiasVacaSFicha = SaldoActual
End Function


Public Function CalcularDiasAntiguedad(FechaInicio, FechaFin) As Double

Dim fs As Boolean
fs = False

'///////// Inicializo parametros generales ////////////
Dim VacacionesAcumuladas, VacacionesSolicitadas, SaldoActual, tempVacacionSolicitada, tempVacacionAcumulada As Double
Dim TotalAcumuladas As Double, TotalSolicitadas As Double, TemporalDias As Double

Dim Inicio As Date
Dim tempInicio As Date
Dim Fin As Date, tempFin As Date

Inicio = FechaInicio
Fin = FechaFin
tempInicio = FechaInicio
'tempFin = DateSerial(Year(FechaInicio), Month(FechaInicio) + 1, 0)
''tempFin = FechaFin

    If DateDiff("d", Inicio, Fin) > 30 Then
      tempFin = DateSerial(Year(FechaInicio), Month(FechaInicio) + 1, 0)
    Else
      tempFin = FechaFin
    End If


         SaldoActual = 0
         tempVacacionSolicitada = 0
         tempVacacionAcumulada = 0
         tempDias = 0
         TemporalDias = 0
         
         
         Dim Tipo As String
         MDIPrimero.DtaControles.Refresh
         Tipo = MDIPrimero.DtaControles.Recordset("DiasMes")
         
    Do While (Inicio < Fin)
 
          If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
          End If
          

            
            Dim DiasMes As Double
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
            If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
                 
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 3) * 0.0833, "####0.00")  '/ 12
                        TemporalDias = DateDiff("d", tempInicio, tempFin) + 1
                    Else
                        tempVacacionAcumulada = 2.5
                        TemporalDias = 28
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 2) * 0.0833, "####0.00")  ' / 12
                        TemporalDias = DateDiff("d", tempInicio, tempFin) + 1
                    Else
                        tempVacacionAcumulada = 2.5
                        TemporalDias = 29
                    End If
                End If
                
               'if tipo = 31
              Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 4) * 0.0833, "####0.00") '/ 12
                        TemporalDias = DateDiff("d", tempInicio, tempFin) + 1
                    Else
                        tempVacacionAcumulada = 2.5
                         TemporalDias = 28
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 3) * 0.0833, "####0.00")  '/ 12
                        TemporalDias = DateDiff("d", tempInicio, tempFin) + 1
                    Else
                        tempVacacionAcumulada = 2.5
                         TemporalDias = 29
                    End If
                End If
               End If
            Else
                If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                         TemporalDias = DateDiff("d", tempInicio, tempFin) + 1
                    Else
                        tempVacacionAcumulada = 2.5
                         TemporalDias = 30
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                         TemporalDias = DateDiff("d", tempInicio, tempFin) + 1
                    Else
                        tempVacacionAcumulada = 2.58
                        TemporalDias = 30
                    End If
                End If
            End If

                 tempVacacionSolicitada = 0
            
             
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
             TotalAcumuladas = TotalAcumuladas + tempVacacionAcumulada
             TotalSolicitadas = TotalSolicitadas + tempVacacionSolicitada
             tempDias = tempDias + TemporalDias
                         

             If DateAdd("d", 2, tempFin) >= Fin Then
             
                 If DateDiff("d", tempFin, Fin) <= 1 Then
                   tempInicio = Fin
                   tempFin = Fin
                 Else
'                    tempFin = DateAdd("d", -2, tempFin)
                    tempFin = DateAdd("d", 2, tempFin)
                    tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                    tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                 End If

                 'ponerle temp inicio para que los dias  no varien
                 Inicio = Fin
             Else
                tempFin = DateAdd("d", 2, tempFin)
                tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                'ponerle temp inicio para que los dias  no varien
                Inicio = tempInicio
             End If
             

            
           ' ////////////////
            
            
             If DateSerial(Year(tempFin), Month(tempFin) + 1, 0) >= Fin Then
                 tempFin = Fin


          If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
          End If
            
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
            If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
               
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                         TemporalDias = DateDiff("d", tempInicio, tempFin) + 1
                    Else
                        tempVacacionAcumulada = 2.5
                         TemporalDias = 28
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                        TemporalDias = DateDiff("d", tempInicio, tempFin) + 1
                    Else
                        tempVacacionAcumulada = 2.5
                         TemporalDias = 29
                    End If
                End If
                
               'if tipo = 31
               Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                         TemporalDias = DateDiff("d", tempInicio, tempFin) + 1
                    Else
                        tempVacacionAcumulada = 2.5
                         TemporalDias = 28
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                         TemporalDias = DateDiff("d", tempInicio, tempFin) + 1
                    Else
                        tempVacacionAcumulada = 2.5
                         TemporalDias = 29
                    End If
                End If
               End If
            Else
                If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                         TemporalDias = DateDiff("d", tempInicio, tempFin) + 1
                    Else
                        tempVacacionAcumulada = 2.5
                        TemporalDias = 30
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.033, "####0.00")  '/ 12
                         TemporalDias = DateDiff("d", tempInicio, tempFin) + 1
                    Else
                        tempVacacionAcumulada = 2.58
                         TemporalDias = 30
                    End If
                End If
            End If
 
             tempVacacionSolicitada = 0
             
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
             TotalAcumuladas = TotalAcumuladas + tempVacacionAcumulada
             TotalSolicitadas = TotalSolicitadas + tempVacacionSolicitada
             tempDias = tempDias + TemporalDias

             tempFin = DateAdd("d", 2, tempFin)
             
             tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1) ' Inicio
             tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin
                         'ponerle temp inicio para que los dias  no varien
            Inicio = tempInicio
            Inicio = DateAdd("m", 1, Inicio)
            End If

    Loop
    
    CalcularDiasAntiguedad = TotalAcumuladas
End Function

Public Function CalcularDiasAguinaldo(CodEmpleado As String, FechaInicio As Date, FechaFin As Date) As Double

Dim fs As Boolean
fs = False


'///////// Inicializo parametros generales ////////////
Dim VacacionesAcumuladas, VacacionesSolicitadas, SaldoActual, tempVacacionSolicitada, tempVacacionAcumulada As Double
Dim TotalAcumuladas As Double, TotalSolicitadas As Double

Dim Inicio As Date
Dim tempInicio As Date
Dim Fin As Date, tempFin As Date

Inicio = FechaInicio
Fin = FechaFin
tempInicio = FechaInicio

If DateDiff("d", Inicio, Fin) > 30 Then
  tempFin = DateSerial(Year(FechaInicio), Month(FechaInicio) + 1, 0)
Else
  tempFin = FechaFin
End If
'


         SaldoActual = 0
         tempVacacionSolicitada = 0
         tempVacacionAcumulada = 0
         
         
         Dim Tipo As String
         MDIPrimero.DtaControles.Refresh
         Tipo = MDIPrimero.DtaControles.Recordset("DiasMes")
         
    Do While (Inicio < Fin)
 
          If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
          End If
          

            
            Dim DiasMes As Double
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
        If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
                 
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                       If Day(tempFin) >= 28 And UCase(Format(tempFin, "MMMM")) = "FEBRERO" Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 3) * 0.0833, "####0.00")  '/ 12
                       Else
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                       End If
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                       If Day(tempFin) >= 28 And UCase(Format(tempFin, "MMMM")) = "FEBRERO" Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 2) * 0.0833, "####0.00")  '/ 12
                       Else
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                       End If
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
                
               'if tipo = 31
               Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                      If Day(tempFin) >= 28 And UCase(Format(tempFin, "MMMM")) = "FEBRERO" Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 4) * 0.0833, "####0.00")  '/ 12
                       Else
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                       End If
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                       If Day(tempFin) >= 28 And UCase(Format(tempFin, "MMMM")) = "FEBRERO" Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 3) * 0.0833, "####0.00")  '/ 12
                       Else
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                       End If
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
               End If
        Else
                If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.58
                    End If
                End If
            End If

             '////////     /////////     ////////        /////////       /////////       ////////        //////
             '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
             '/////////////////////////////////////////////////////////////////////////////////////////////////
           If DiasMes = 31 Then
                MDIPrimero.DtaConsulta.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM SolicitudVacaciones WHERE  TipoSolicitud = 'Suspension' and    (CodigoEmpleado = '" & CodEmpleado & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
           Else
                MDIPrimero.DtaConsulta.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE TipoSolicitud = 'Suspension' and  (CodigoEmpleado = '" & CodEmpleado & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
           End If
           
             MDIPrimero.DtaConsulta.Refresh
             If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                If Not IsNull(MDIPrimero.DtaConsulta.Recordset("VacacionesSolicitadas")) Then
                    tempVacacionSolicitada = CDbl(MDIPrimero.DtaConsulta.Recordset("VacacionesSolicitadas")) * 0.0833
                Else
                    tempVacacionSolicitada = 0
                End If
             Else
                 tempVacacionSolicitada = 0
             End If
             
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
             TotalAcumuladas = TotalAcumuladas + tempVacacionAcumulada
             TotalSolicitadas = TotalSolicitadas + tempVacacionSolicitada
                         

             If DateAdd("d", 2, tempFin) >= Fin Then
             
             
'                 tempFin = DateAdd("d", -2, tempFin)
                 If DateDiff("d", tempFin, Fin) = 1 Then
                   tempInicio = Fin
                   tempFin = Fin
                 Else
                     tempFin = DateAdd("d", 2, tempFin)
                     tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                     tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                 End If
                 

                 'ponerle temp inicio para que los dias  no varien
                 Inicio = Fin
             Else
                tempFin = DateAdd("d", 2, tempFin)
                tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                'ponerle temp inicio para que los dias  no varien
                Inicio = tempInicio
             End If
             

            
           ' ////////////////
            
            
             If DateSerial(Year(tempFin), Month(tempFin) + 1, 0) >= Fin Then
                 tempFin = Fin


          If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
          End If
            
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
            If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
                 
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                      If Day(tempFin) >= 28 And UCase(Format(tempFin, "MMMM")) = "FEBRERO" Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 3) * 0.0833, "####0.00")  '/ 12
                       Else
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                       End If
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                       If Day(tempFin) >= 28 And UCase(Format(tempFin, "MMMM")) = "FEBRERO" Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 2) * 0.0833, "####0.00")  '/ 12
                       Else
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                       End If
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
                
               'if tipo = 31
               Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                       If Day(tempFin) >= 28 And UCase(Format(tempFin, "MMMM")) = "FEBRERO" Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 3) * 0.0833, "####0.00")  '/ 12
                       Else
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                       End If
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                       If Day(tempFin) >= 28 And UCase(Format(tempFin, "MMMM")) = "FEBRERO" Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 2) * 0.0833, "####0.00")  '/ 12
                       Else
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                       End If
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
               End If
            Else
                If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = Format((DateDiff("d", tempInicio, tempFin) + 1) * 0.0833, "####0.00")  '/ 12
                    Else
                        tempVacacionAcumulada = 2.58
                    End If
                End If
            End If
          

             '////////     /////////     ////////        /////////       /////////       ////////        //////
             '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
             '/////////////////////////////////////////////////////////////////////////////////////////////////
           If DiasMes = 31 Then
                MDIPrimero.DtaConsulta.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM SolicitudVacaciones WHERE  TipoSolicitud = 'Suspension' and    (CodigoEmpleado = '" & CodEmpleado & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
           Else
                MDIPrimero.DtaConsulta.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM SolicitudVacaciones WHERE  TipoSolicitud = 'Suspension' and    (CodigoEmpleado = '" & CodEmpleado & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
           End If
             MDIPrimero.DtaConsulta.Refresh
             If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                If Not IsNull(MDIPrimero.DtaConsulta.Recordset("VacacionesSolicitadas")) Then
                    tempVacacionSolicitada = CDbl(MDIPrimero.DtaConsulta.Recordset("VacacionesSolicitadas")) * 0.0833333
                Else
                    tempVacacionSolicitada = 0
                End If
             Else
                 tempVacacionSolicitada = 0
             End If
             
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
             TotalAcumuladas = TotalAcumuladas + tempVacacionAcumulada
             TotalSolicitadas = TotalSolicitadas + tempVacacionSolicitada

             tempFin = DateAdd("d", 2, tempFin)
             
             tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1) ' Inicio
             tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin
                         'ponerle temp inicio para que los dias  no varien
            Inicio = tempInicio
            Inicio = DateAdd("m", 1, Inicio)
            End If

    Loop
    
    
    
    CalcularDiasAguinaldo = SaldoActual
End Function

Public Sub KillProcess(ByVal processName As String)
On Error GoTo errHandler
Dim oWMI
Dim ret
Dim sService
Dim oWMIServices
Dim oWMIService
Dim oServices
Dim oService
Dim servicename
Set oWMI = GetObject("winmgmts:")
Set oServices = oWMI.InstancesOf("win32_process")
For Each oService In oServices

servicename = LCase(Trim(CStr(oService.Name) & ""))

If InStr(1, servicename, LCase(processName), vbTextCompare) > 0 Then
ret = oService.Terminate
End If

Next

Set oServices = Nothing
Set oWMI = Nothing

errHandler:
Err.Clear
End Sub
Public Function ConsecutivoSubsidio(Tabla As String) As Double
  MDIPrimero.DtaConsulta.RecordSource = "SELECT  * From Subsidio ORDER BY NumSubsidio"
  MDIPrimero.DtaConsulta.Refresh
  If MDIPrimero.DtaConsulta.Recordset.EOF Then
    ConsecutivoSubsidio = 1
  Else
    MDIPrimero.DtaConsulta.Recordset.MoveLast
    ConsecutivoSubsidio = MDIPrimero.DtaConsulta.Recordset("NumSubsidio") + 1
  End If
  
End Function
Public Function BuscaIncioPeriodo(Año As String, Mes As String, CodTipoNomina As String) As Date
    MDIPrimero.DtaConsulta2.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año & ") AND (mes = '" & Mes & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
    MDIPrimero.DtaConsulta2.Refresh
     If Not MDIPrimero.DtaConsulta2.Recordset.EOF Then
       BuscaIncioPeriodo = MDIPrimero.DtaConsulta2.Recordset("Inicio")
     End If
     
End Function
Public Function BuscaFinPeriodo(Año As String, Mes As String, CodTipoNomina As String) As Date
     
    MDIPrimero.DtaConsulta2.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año & ") AND (mes = '" & Mes & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
    MDIPrimero.DtaConsulta2.Refresh
     If Not MDIPrimero.DtaConsulta2.Recordset.EOF Then
       MDIPrimero.DtaConsulta2.Recordset.MoveLast
       BuscaFinPeriodo = MDIPrimero.DtaConsulta2.Recordset("Final")
     End If
End Function
Public Function CalcularIr(MontoBrutoMensual As Double, PeriodoNomina As String) As Double
  Dim MinIR As Double, CantSabados As Double
  Dim MontoIr As Double, MontoIRPatronal As Double
  Dim TipoCalculoIr As String
  
   MDIPrimero.DtaEmpresa.Refresh
    If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
      TipoCalculoIr = MDIPrimero.DtaEmpresa.Recordset("TipoCalculoIR")
    End If

        CantSabados = 4
        MDIPrimero.DtaIR.RecordSource = "SELECT * From IR"
        MDIPrimero.DtaIR.Refresh
        MDIPrimero.DtaIR.Recordset.MoveNext
        MinIR = MDIPrimero.DtaIR.Recordset("desde")
        MinIR = MinIR - 1
        MinIR = (MinIR / 12)
     '   MsgBox MinIR
        Do While Not MDIPrimero.DtaIR.Recordset.EOF
        
           'ubicar la linea
         If PeriodoNomina = "Semanal Viernes" Then
            If (MontoBrutoMensual) >= MinIR Then
            If MDIPrimero.DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And MDIPrimero.DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - MDIPrimero.DtaIR.Recordset("SobreExceso")) * (MDIPrimero.DtaIR.Recordset("PorcientoImpuesto") / 100) + MDIPrimero.DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / 12, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do
            End If
            End If
            
         ElseIf FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
            If (MontoBrutoMensual) >= MinIR Then
            If MDIPrimero.DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And MDIPrimero.DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - MDIPrimero.DtaIR.Recordset("SobreExceso")) * (MDIPrimero.DtaIR.Recordset("PorcientoImpuesto") / 100) + MDIPrimero.DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / CantSabados / 12, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do
                       
            End If
            End If
            
        ElseIf FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
            If (MontoBrutoMensual) >= MinIR Then
            If MDIPrimero.DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And MDIPrimero.DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - MDIPrimero.DtaIR.Recordset("SobreExceso")) * (MDIPrimero.DtaIR.Recordset("PorcientoImpuesto") / 100) + MDIPrimero.DtaIR.Recordset("ImpuestoBase")
  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
               If DiaFin < 28 Then
                MontoIr = Format(MontoIr / 2 / 12, "###,##0.00")
                MontoIRPatronal = MontoIr
                Exit Do
               Else
                MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
                MontoIr = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
               End If
            End If
            Else
               MontoIrMensual = 0
               MontoIr = MontoIrMensual - MontoIrAnterior
               MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If
         ElseIf FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
            If (MontoBrutoMensual) >= MinIR Then
            If MDIPrimero.DtaIR.Recordset("desde") <= (MontoBruto * 26) And MDIPrimero.DtaIR.Recordset("Hasta") >= (MontoBruto * 26) Then
               MontoIr = ((MontoBruto * 26) - MDIPrimero.DtaIR.Recordset("SobreExceso")) * (MDIPrimero.DtaIR.Recordset("PorcientoImpuesto") / 100) + MDIPrimero.DtaIR.Recordset("ImpuestoBase")
  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
               If DiaFin < 28 Then
                MontoIr = Format(MontoIr / 26, "###,##0.00")
                MontoIRPatronal = MontoIr
                Exit Do
               Else
                MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
                MontoIr = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
               End If
            End If
            Else
               MontoIrMensual = 0
                MontoIr = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If
         ElseIf FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
'            If (MontoBrutoMensual) >= MinIR Then

              '//////////////////////////FORMULA  IR PROGRESIVA ////////////////////////////////////////
              
                'RentaGravable = ((IngresoAcumulado + IngresoMes)/NQuincenas)*24 (Quincenas al ano)
                'MontoIr= (RentaGravable - SobreExceso*%Impuesto + ImpuestoBase-IrAcumulado)/24-(NQuincenas-1)
                

             If MDIPrimero.DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And MDIPrimero.DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - MDIPrimero.DtaIR.Recordset("SobreExceso")) * (MDIPrimero.DtaIR.Recordset("PorcientoImpuesto") / 100) + MDIPrimero.DtaIR.Recordset("ImpuestoBase")
'///////Ve      rfico si el la Ultima Quincena para hacer ajustes////////////

                If TipoCalculoIr = "Calcular IR x 12" Then
                    If DiaFin < 28 Then
                        MontoIr = Format(MontoIr / 2 / 12, "###,##0.00")
                        MontoIRPatronal = MontoIr
                        Exit Do
                    Else
'                        MontoIrMensual = Format(MontoIR / 1 / 12, "###,##0.00")
'                        MontoIR = MontoIrMensual - MontoIrAcumulado
                        MontoIr = Format(MontoIr / 2 / 12, "###,##0.00")
                        MontoIRPatronal = MontoIr
                    End If
                Else
'                If Not NumeroPeriodo = 0 Then
'                  'NumeroPeriodo = 24-(NQuincenas-1)
'                 MontoIr = (MontoIr - MontoIrAcumulado) / NumeroPeriodo
'                Else
'                 MontoIr = 0
'                End If
                End If
                
                MontoIRPatronal = MontoIr - MontoIrPatronalAnterior
                Exit Do
'               End If
             End If
'            Else
'               MontoIrMensual = 0
               
'                MontoIR = MontoIrMensual - MontoIrAnterior
'                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'            End If

          
         
         ElseIf FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
'           If (MontoBrutoAnual) >= MinIR Then
            If MDIPrimero.DtaIR.Recordset("desde") <= (MontoBrutoAnual) And MDIPrimero.DtaIR.Recordset("Hasta") >= (MontoBrutoAnual) Then

               MontoIr = ((MontoBrutoAnual) - MDIPrimero.DtaIR.Recordset("SobreExceso")) * (MDIPrimero.DtaIR.Recordset("PorcientoImpuesto") / 100) + MDIPrimero.DtaIR.Recordset("ImpuestoBase")

                MontoIr = (MontoIr - MontoIrAcumulado) / NumeroPeriodo
                MontoIRPatronal = MontoIr - MontoIrPatronalAnterior
                Exit Do

               Exit Do
            End If
'         End If
         ElseIf FrmCalcularNomina.DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
           If (MontoBrutoMensual) >= MinIR Then
            If MDIPrimero.DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And MDIPrimero.DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - MDIPrimero.DtaIR.Recordset("SobreExceso")) * (MDIPrimero.DtaIR.Recordset("PorcientoImpuesto") / 100) + MDIPrimero.DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / 4, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do
            End If
           End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
             If (MontoBrutoMensual) >= MinIR Then
            If MDIPrimero.DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And MDIPrimero.DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - MDIPrimero.DtaIR.Recordset("SobreExceso")) * (MDIPrimero.DtaIR.Recordset("PorcientoImpuesto") / 100) + MDIPrimero.DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / 2, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do
            End If
            End If
         End If
  MDIPrimero.DtaIR.Recordset.MoveNext
  Loop
  
  CalcularIr = MontoIr
End Function




Public Function FechaVacaciones(FechaFin As Date) As Date
Dim SqlSalarios, FechaInicio As Date


                        SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos, SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Comisiones) As Comisiones FROM  DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY Nomina.Mes, Nomina.Ano  " & _
                                      "HAVING  (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.Comisiones) <> 0) AND (MAX(Nomina.FechaNomina) <= CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) ORDER BY Nomina.Ano, Nomina.Mes"
                        
                        MDIPrimero.DtaConsulta.RecordSource = SqlSalarios
                        MDIPrimero.DtaConsulta.Refresh
                        If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                         MDIPrimero.DtaConsulta.Recordset.MoveLast
                        Else
'                         FechaFin = Format(Now, "dd/mm/yyyy")
                         FechaInicio = Format(Now, "dd/mm/yyyy")
                        End If
                        i = 0
                        Do While Not MDIPrimero.DtaConsulta.Recordset.BOF
                          If i = 1 Then
                            FechaInicio = MDIPrimero.DtaConsulta.Recordset("FechaInicio")
                        
                          ElseIf i = 5 Then
                            FechaInicio = MDIPrimero.DtaConsulta.Recordset("FechaInicio")
                            Exit Do
                          ElseIf i = 0 Then
                            FechaInicio = MDIPrimero.DtaConsulta.Recordset("FechaInicio")
'                            FechaFin = frmempleado.DtaConsulta.Recordset("FechaFin")
                          Else
                            FechaInicio = MDIPrimero.DtaConsulta.Recordset("FechaInicio")
                          End If
                          i = i + 1
                        
                          MDIPrimero.DtaConsulta.Recordset.MovePrevious
                        Loop
                        
                        FechaVacaciones = FechaInicio

End Function
'-------------------------------------------------------------------------------------------
'-------------------------------------FUNCIONES DEL RELOJ --------------------------------
'-------------------------------------------------------------------------------------------

Function RestaAlmuerzo(CodHorario As String, Dia As Double) As Double
  Dim HoraInicio As String, HoraFin As String, HoraAlmuerzo As Double, RestarAlmuerzo As Boolean, ExcluirSabados As Boolean
  
   If CodHorario = "" Then
    RestaAlmuerzo = 0
    Exit Function
   End If
  
   MDIPrimero.AdoConsulta.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodHorario & "))"
   MDIPrimero.AdoConsulta.Refresh
   If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
     If Not IsNull(MDIPrimero.AdoConsulta.Recordset("EntradaAlmuerzo")) Then
       HoraInicio = MDIPrimero.AdoConsulta.Recordset("EntradaAlmuerzo")
     End If
     If Not IsNull(MDIPrimero.AdoConsulta.Recordset("SalidaAlmuerzo")) Then
       HoraFin = MDIPrimero.AdoConsulta.Recordset("SalidaAlmuerzo")
     End If
     
       RestarAlmuerzo = MDIPrimero.AdoConsulta.Recordset("RestarAlmuerzo")
       ExcluirSabados = MDIPrimero.AdoConsulta.Recordset("ExcluirSabado")
       
       If RestarAlmuerzo = True Then
          HoraAlmuerzo = DateDiff("n", HoraInicio, HoraFin) / 60
       Else
          HoraAlmuerzo = 0
       End If
       
       If ExcluirSabados = True Then
        Select Case Dia
           Case 6:  HoraAlmuerzo = 0
           Case 0:  HoraAlmuerzo = 0
        End Select
       End If
       
 
    RestaAlmuerzo = HoraAlmuerzo
    
  
   Else
     RestaAlmuerzo = 0
   End If
   
End Function

Function Buscar_Carpeta(Optional Titulo As String, _
                        Optional Path_Inicial As Variant) As String
  
On Local Error GoTo errFunction
      
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
      
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
      
    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder( _
                            0, _
                            Titulo, _
                            0, _
                            Path_Inicial)
      
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self
      
    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.Path
  
Exit Function
'Error
errFunction:
    MsgBox Err.Description, vbCritical
    Buscar_Carpeta = vbNullString
  
End Function


Public Function sumaHoras(H1 As String, H2 As String) As String
    Dim vh1 As Variant
    Dim vh2 As Variant
    Dim intContador As Integer
    Dim vh3(2) As Long
    Dim H3 As String
    
    'Convertir a arrays
    vh1 = Split(H1, ":")
    vh2 = Split(H2, ":")
    
    'Contemplar tambien los segundos
    For intContador = 0 To 2
    
    'Sumar las horas, minutos, segundos
    If intContador <= UBound(vh1) Then vh3(intContador) = val(vh1(intContador))
    If intContador <= UBound(vh2) Then vh3(intContador) = vh3(intContador) + val(vh2(intContador))
    Next intContador
    
    'Descontar las cantidades mayores de 60 en 1 y 2
    vh3(1) = vh3(1) + vh3(2) \ 60
    vh3(2) = vh3(2) Mod 60
    vh3(0) = vh3(0) + vh3(1) \ 60
    vh3(1) = vh3(1) Mod 60
    
    'Constuir la cadena a devolver
    sumaHoras = Format(vh3(0), "00") & ":" & Format(vh3(1), "00") & ":" & Format(vh3(2), "00")

End Function




Public Function DiaHorario(FechaIni As Date, FechaFin As Date, Ciclos As Double) As Double
 Dim i As Double, j As Double, Dias As Double, Diffechas As Double, DiaInicio As Double
 Dim n As Double
 
 Dias = Ciclos * 7

 
 DiaInicio = DiaSemana(Day(FechaIni), Month(FechaIni), Year(FechaIni))
 
 '*****************************************************************************************
 '//////////////////CALCULO EL NUMERO DE DIAS ENTRE LAS DOS FECHAS /////////////////////////
 '******************************************************************************************
 Diffechas = DateDiff("d", FechaIni, FechaFin) + 1
 
 i = DiaInicio
 j = 1
 Do While j <= Diffechas
    If i > (Dias - 1) Then
      i = 0
    End If
    
    n = i
    
    i = i + 1
    j = j + 1
  Loop
 
 DiaHorario = n  '+ (Ciclos - 1) * 7

End Function

Public Function ConvertirSegundos(Segundos As Double, Dia As Double) As String
Dim Horas As Double
Dim Minutos As Double
Dim Cadena As String
Dim RestarAlmuerzo As Double

If CodigoH <> "" Then
 RestarAlmuerzo = RestaAlmuerzo(CodigoH, Dia)
Else
 RestarAlmuerzo = 0
End If

If Segundos > 0 Then
    Horas = Segundos / 3600
    
    If Horas > 0 Then '/////RESTO EL ALMUERZO //////////////////
      Horas = (Segundos / 3600) - RestarAlmuerzo
      Segundos = Horas * 3600
    End If
    
    Horas = Int(Segundos / 3600)
    Minutos = Int((Segundos Mod 3600) / 60)
    Cadena = Horas & ":" & Minutos
Else
    Cadena = "00:00"
End If
ConvertirSegundos = Cadena

End Function
Public Function ConvertirS(Segundos As Double) As String
Dim Horas As Double
Dim Minutos As Double
Dim Cadena As String

On Error GoTo TipoErrs

If Segundos > 0 Then
    Horas = Int(Segundos / 3600)
    Minutos = Int((Segundos Mod 3600) / 60)
    
    Cadena = Horas & ":" & Minutos
Else
    Cadena = "00:00"
End If
ConvertirS = Cadena

TipoErrs:
 


End Function


Public Function ConvertirSegundosHoras(Segundos As Double) As Double
Dim Horas As Double
Dim Minutos As Double
Dim Cadena As String

If Segundos > 0 Then
    Horas = Int(Segundos / 3600)
   
    If Horas > 0 Then '/////RESTO EL ALMUERZO //////////////////
     Horas = Horas - 1
    End If

End If

ConvertirSegundosHoras = Horas

End Function
Public Function ConvertirSegundosMinutos(Segundos As Double) As Double
Dim Horas As Double
Dim Minutos As Double
Dim Cadena As String

On Error GoTo TipoErrs

If Segundos > 0 Then
    Minutos = Int((Segundos Mod 3600) / 60)
End If

Segundos = 0

ConvertirSegundosMinutos = Minutos

TipoErrs:
' MsgBox Err.Description
 

End Function

Public Function DiaSemana(Dias As Double, Mes As Double, Año As Double) As Double
 Dim A As Double, AñoQ As Double, DosD As Double
 Dim b As Double, c As Integer, D As Double, E As Double, F As Double, R As Double
 

 
 '//////////////////////////////////////////////////////////////////////////
 '////////////////////BUSCAMOS EL SIGLO, DEPENDIENDO DEL ANO ///////////////
 '//////////////////////////////////////////////////////////////////////////
' 1700…1799   1800…1899   1900…1999   2000…2099   2100…2199   2200…2299
'     +5         +3          +1           0          -2          -4

    If Año >= 1700 And Año <= 1799 Then
      A = 5
    ElseIf Año >= 1800 And Año <= 1899 Then
      A = 3
    ElseIf Año >= 1900 And Año <= 1999 Then
      A = 1
    ElseIf Año >= 2000 And Año <= 2099 Then
      A = 0
    ElseIf Año >= 2100 And Año <= 2199 Then
      A = -2
    ElseIf Año >= 2200 And Año <= 2299 Then
      A = -4
    End If
    
 '//////////////////////////////////////////////////////////////////////////////////////
 '//////////////////////////////CALCULO EL CUARTO DEL LOS ULTIMOS DIGITOS DEL ANO ////
 '//////////////////////////////////////////////////////////////////////////////////////
    DosD = Mid(Año, 3, 2)
    AñoQ = Int(DosD / 4)
    b = DosD + AñoQ
 
 '////////////////////////////////////////////////////////////////////////////
 '/////////////////////////CALCULO LOS Años BISIESTOS  ////////////////////////
 '////////////////////////////////////////////////////////////////////////////
 ' Años bisiestos: Éstos son los que cumplen que sus dos últimas cifras forman un múltiplo de 4
 '                 (por ejemplo, 1992 o 2004) excepto los terminados en 00. Entre estos últimos sólo son bisiestos los
 '                 múltiplos de cuatrocientos (por ejemplo 2000). Nuestro tercer coeficiente, C depende de ellos:
 '                 si el año es bisiesto, y el mes es enero o febrero el coeficiente será C = –1. En cualquier otro caso C = 0.
 '                 En nuestro ejemplo, como 2007 no es bisiesto tenemos que C = 0.
 
 '/////BUSCO SI SON MULTIPLOS DE CUATRO //////////////////////////////////////
 

    c = 0
    
   If DosD <> "00" Then
         If val(DosD / 4) - Int(val(DosD / 4)) = 0 Then
          '///////SI SON ENTEROS SON MULTIPLOS DE  4
          '///AHORA CONSULTO EL MES CORRESPONDIENTE ///////
          If Mes = 1 Or Mes = 2 Then
           c = 1 - 2
          End If
        End If
   Else
         If val(Año / 400) - Int(val(Año / 400)) = 0 Then
          '///////SI SON ENTEROS SON MULTIPLOS DE  400
          '///AHORA CONSULTO EL MES CORRESPONDIENTE ///////
          If Mes = 1 Or Mes = 2 Then
           c = 1 - 2
          End If
        End If
   End If
    
 '/////////////////////////////////////////////////////////////////////////////////
 '//////////////////////CALCULO EL FACTOR PARA EL MES //////////////////////////////
 '/////////////////////////////////////////////////////////////////////////////////
' Enero   Feb.    Marzo   Abril   Mayo    Junio   Julio   Agosto  Sept.   Oct.    Nov.    Dic.
'  6       2        2       5       0       3       5       1       4      6       2       4
 Select Case Mes
   Case 1: D = 6
   Case 2: D = 2
   Case 3: D = 2
   Case 4: D = 5
   Case 5: D = 0
   Case 6: D = 3
   Case 7: D = 5
   Case 8: D = 1
   Case 9: D = 4
   Case 10: D = 6
   Case 11: D = 2
   Case 12: D = 4
 End Select
 
 '/////////////////////////////////////////////////////////////////////////////////////////
 '/////////////////////CALCULO EL FACTOR DEL DIA ///////////////////////////////////////////
 '/////////////////////////////////////////////////////////////////////////////////////////
 E = Dias

  '/////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////CORREOMOS EL ALGORITMO PARA VER EL DIA /////////////////////////////
  '//////////////////////////////////////////////////////////////////////////////////////////
'  Lunes   Martes  Miércoles   Jueves  Viernes     Sábado  Domingo
'    1        2        3          4       5           6       0

  F = A + b + c + D + E
  R = F - 7
  
  Do While R > 6
   R = R - 7
  Loop

DiaSemana = R

End Function


Public Function Dia(Fecha As Date) As String
 Dim DiaSemana As Double
 
 DiaSemana = Weekday(Fecha)
 Select Case DiaSemana
    Case 1: Dia = "Domingo"
    Case 2: Dia = "Lunes"
    Case 3: Dia = "Martes"
    Case 4: Dia = "Miercoles"
    Case 5: Dia = "Jueves"
    Case 6: Dia = "Viernes"
    Case 7: Dia = "Sabado"
    
 End Select

End Function





