VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARBaja 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ARBaja.dsx":0000
End
Attribute VB_Name = "ARBaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Dim SqlReport As String

Me.LblMeses.Caption = FrmBajas.TxtMeses.Text

'DaoDtaBajas.DatabaseName = ruta
'DaoDtaBajas.ConnectionString = Conexion
SqlReport = "SELECT Bajas.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento,  Bajas.AnnosTrabajados, Bajas.MesesTrabajados, Bajas.DiasTrabajados, Bajas.MontoNomPropor, Bajas.FechaBaja, Bajas.MontoVaca, Bajas.Monto13Mes, Bajas.MontoAnosTrab, Bajas.MontoCargoConfianza, Bajas.MontoAntiguedad, Bajas.MotivoBaja, Bajas.TipoBaja, Bajas.Otro, Bajas.MontoOtro, Bajas.Prestamo, Bajas.Deducciones, [Bajas].[MontoVaca]+[Bajas].[Monto13Mes]+[Bajas].[MontoAnosTrab]+[Bajas].[MontoCargoConfianza]+[Bajas].[MontoAntiguedad]+[Bajas].[MontoOtro]-[Bajas].[Prestamo]-[Bajas].[Deducciones] AS TotalaPagar FROM (Departamento INNER JOIN (Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) ON Departamento.CodDepartamento = Empleado.CodDepartamento) INNER JOIN Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado Where Bajas.CodEmpleado = '" & CodEmpleado & "'"
'DaoDtaBajas.RecordSource = SqlReport
'DaoDtaBajas.Refresh
LblTitulo.Caption = Titulo
LblSubtitulo.Caption = SubTitulo
Me.LblReferencia.Caption = FrmBajas.Referencia
ImgLogo.Picture = LoadPicture(App.Path + "\fotos\Zw.bmp")


End Sub

Private Sub GroupHeader1_Format()
Dim sql As String, Nombres As String, Empresa As String, RutaLogo As String
Dim FechaBusqueda As Date, FechaHistorico As Date, Destino As String, SueldoActual As Double
Dim Contador As Double, TotalSalario As Double, Salario As Double, SalarioAlto As Double
Dim DiasTrabajados As Double
 '//////////////////////////////////////////////////////////////////////////////////////////
 '///////////////////////ASIGNO EL SUBREPORTE/////////////////////////////////////////////////////
 '////////////////////////////////////////////////////////////////////////////////////////
 
 DiasTrabajados = Me.FldDiasTrabajados.Text
 
 
'    SQlEmpleado = "SELECT Empleado.SalarioFijo, Empleado.SueldoPeriodo, Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion, Empleado.Direccion AS Expr1, Empleado.Sexo, Empleado.Activo, Empleado.TarifaHoraria, Empleado.SueldoActualBasico, Historico.SueldoActual FROM Departamento INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado Where (Empleado.CodEmpleado = " & TxtCodEmpleado.Text & ")"
'    DtaEmpleado.RecordSource = SQlEmpleado
   
 
'
  FechaBusqueda = FrmBajas.FechaBusqueda1
  FechaHistorico = FrmBajas.FechaHistorico1

  If FrmBajas.ChkSueldoActual.Value = xtpUnchecked Then
     If FrmBajas.ChkIncentivos.Value = xtpUnchecked Then
        sql = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Antiguedad) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(0) AS Incentivos,SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + DetalleNomina.Antiguedad)AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO,SUM(DetalleNomina.Reembolso) As Comisiones " & _
                      "FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & FrmBajas.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) Between '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') " & _
                      "ORDER BY Nomina.Ano, Nomina.Mes "
     Else
        sql = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Antiguedad) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + DetalleNomina.Incentivos + DetalleNomina.Antiguedad)AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO,SUM(DetalleNomina.Reembolso) As Comisiones " & _
                      "FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & FrmBajas.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) Between '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') " & _
                      "ORDER BY Nomina.Ano, Nomina.Mes "
     End If
   Else
      If FrmBajas.ChkIncentivos.Value = xtpUnchecked Then
              sql = "SELECT DISTINCT TOP (100) PERCENT DetalleNomina.CodEmpleado, AVG(Historico.SueldoActual) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia * 0) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(0) AS Incentivos,  SUM(DetalleNomina.Destajo + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso) + AVG(Historico.SueldoActual) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) As Comisiones FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING (DetalleNomina.CodEmpleado = '" & FrmBajas.TxtCodEmpleado.Text & "') AND (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') ORDER BY AÑO, Nomina.Mes"
      Else
              sql = "SELECT DISTINCT TOP (100) PERCENT DetalleNomina.CodEmpleado, AVG(Historico.SueldoActual) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia * 0) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,  SUM(DetalleNomina.Destajo + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + DetalleNomina.Incentivos) + AVG(Historico.SueldoActual) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) As Comisiones FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING (DetalleNomina.CodEmpleado = '" & FrmBajas.TxtCodEmpleado.Text & "') AND (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') ORDER BY AÑO, Nomina.Mes"
      
      End If
   End If

    Set Me.SrpBajas.object = New ArepBajasSubRpt

    Me.SrpBajas.object.DataControl1.ConnectionString = Conexion
    Me.SrpBajas.object.DataControl1.Source = sql
    

  If FrmBajas.ChkSueldoActual.Value = xtpUnchecked Then
    If FrmBajas.ChkIncentivos.Value = xtpUnchecked Then
      MDIPrimero.DtaConsulta.RecordSource = "SELECT DISTINCT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso  + DetalleNomina.HorasExtras + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) AS TotalIngresos, SUM(Nomina.TotalDestajo) AS Expr1 FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina  " & _
                                        "WHERE (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion <> 0) AND (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Format(FechaBusqueda, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaHistorico, "yyyymmdd") & "', 102)) GROUP BY DetalleNomina.CodEmpleado Having (DetalleNomina.CodEmpleado = '" & FrmBajas.TxtCodEmpleado.Text & "')"
    Else
      MDIPrimero.DtaConsulta.RecordSource = "SELECT DISTINCT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) AS TotalIngresos, SUM(Nomina.TotalDestajo) AS Expr1 FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina  " & _
                                        "WHERE (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion <> 0) AND (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Format(FechaBusqueda, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaHistorico, "yyyymmdd") & "', 102)) GROUP BY DetalleNomina.CodEmpleado Having (DetalleNomina.CodEmpleado = '" & FrmBajas.TxtCodEmpleado.Text & "')"
    End If
  Else
      If FrmBajas.ChkIncentivos.Value = xtpUnchecked Then
         MDIPrimero.DtaConsulta.RecordSource = "SELECT DISTINCT TOP (100) PERCENT DetalleNomina.CodEmpleado, AVG(Historico.SueldoActual) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia * 0) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(0) AS Incentivos,  SUM(DetalleNomina.Destajo + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso) + AVG(Historico.SueldoActual) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) As Reembolso FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING (DetalleNomina.CodEmpleado = '" & FrmBajas.TxtCodEmpleado.Text & "') AND (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') ORDER BY AÑO, Nomina.Mes"
      Else
         MDIPrimero.DtaConsulta.RecordSource = "SELECT DISTINCT TOP (100) PERCENT DetalleNomina.CodEmpleado, AVG(Historico.SueldoActual) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia * 0) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,  SUM(DetalleNomina.Destajo + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + DetalleNomina.Incentivos) + AVG(Historico.SueldoActual) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) As Reembolso FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING (DetalleNomina.CodEmpleado = '" & FrmBajas.TxtCodEmpleado.Text & "') AND (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') ORDER BY AÑO, Nomina.Mes"
      End If
  End If
  
  MDIPrimero.DtaConsulta.Refresh
  
    
    Contador = 0
    TotalSalario = 0
    Salario = 0
    SalarioAlto = 0
    Do While Not MDIPrimero.DtaConsulta.Recordset.EOF
    
      If Not IsNull(MDIPrimero.DtaConsulta.Recordset("TotalIngresos")) Then
        TotalSalario = TotalSalario + MDIPrimero.DtaConsulta.Recordset("TotalIngresos")
        Salario = MDIPrimero.DtaConsulta.Recordset("TotalIngresos")
      Else
        Salario = 0
      End If
 
        If Salario > SalarioAlto Then
            SalarioAlto = Salario
        End If
 
        Contador = Contador + 1
        MDIPrimero.DtaConsulta.Recordset.MoveNext
    Loop
  
   
   
     Me.LblTotalSalario.Caption = Format(TotalSalario, "##,##0.00")
 

'-------------------------------------------BUSCO INFORMACION ADICINAL DEL EMPLEADO --------------------
   MDIPrimero.DtaConsulta.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo,Dolarizado,CuentaBanco,SueldoActualBasico From Empleado WHERE (CodEmpleado = '" & FrmBajas.TxtCodEmpleado.Text & "') And (Activo = 1)"
   MDIPrimero.DtaConsulta.Refresh
   If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
     Me.LblNumCedula.Caption = MDIPrimero.DtaConsulta.Recordset("NumCedula")
     Me.LblNumINSS.Caption = MDIPrimero.DtaConsulta.Recordset("NumeroInss")
     Me.LblNumCedula2.Caption = MDIPrimero.DtaConsulta.Recordset("NumCedula")
   End If
   
   Nombres = Me.FldNombres.Text
   Empresa = Me.LblTitulo.Caption
   
   Me.LblNota.Caption = "Yo:" & Nombres & " hago constar que recibo a mi entera satisfacción de parte de " & Empresa & " mi liquidación de prestaciones laborales, no teniendo ningún reclamo anterior ni posterior que hacer por concepto de salario, horas extras, ni ningun otro beneficio que tenga por causa la relación laboral."

        
        If Dir(RutaFoto & FrmBajas.TxtCodEmpleado1.Text & ".jpg") <> "" Then
           Destino = RutaFoto & FrmBajas.TxtCodEmpleado1.Text & ".jpg"
        ElseIf Dir(RutaFoto & FrmBajas.TxtCodEmpleado1.Text & ".gif") <> "" Then
           Destino = RutaFoto & FrmBajas.TxtCodEmpleado1.Text & ".gif"
        ElseIf Dir(RutaFoto & FrmBajas.TxtCodEmpleado1.Text & ".bmp") <> "" Then
           Destino = RutaFoto & FrmBajas.TxtCodEmpleado1.Text & ".bmp"
        End If
        
        If (Dir(Destino) <> "") Then
         Me.ImgFoto.Picture = LoadPicture(Destino)
        Else
          Destino = App.Path + "\Zw.bmp"
'          Destino = RutaLogo
         Me.ImgFoto.Picture = LoadPicture(Destino)
        End If
        
        RutaLogo = MDIPrimero.DtaEmpresa.Recordset("RutaLogo")
         If (Dir(RutaLogo, vbDirectory) <> "") Then
         Me.ImgLogo.Picture = LoadPicture(RutaLogo)
        Else
          RutaLogo = App.Path + "\Zw.bmp"
          Me.ImgLogo.Picture = LoadPicture(RutaLogo)
        End If
        
        
 
        
        Me.LblDiasNeto.Caption = FrmBajas.TxtDias.Text

End Sub

