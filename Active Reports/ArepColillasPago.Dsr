VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepColillasPago 
   Caption         =   "Colillas de Pago"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19368
   SectionData     =   "ArepColillasPago.dsx":0000
End
Attribute VB_Name = "ArepColillasPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ActiveReport_ReportStart()

Me.LblTitulo.Caption = Titulo
Me.LblPeriodo.Caption = PeriodoReporte

' Dim SQlReportes As String
'
'                      SQlReportes = "SELECT Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.BonoProduccion, DetalleNomina.Prestamo, " & _
'                                    "DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre,DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres,DetalleNomina.HE, DetalleNomina.SalarioBasico DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno + DetalleNomina.Antiguedad AS TotalDevengado, DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir,  " & _
'                                    "(DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno + DetalleNomina.Antiguedad) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, Empleado.NumeroInss, Empleado.NumCedula, Departamento.Departamento, DetalleNomina.HorasTurno, DetalleNomina.HTurno, DetalleNomina.Antiguedad, DetalleNomina.AñoAntiguedad , departamento.CodDepartamento, Nomina.FechaNominaINI FROM  Nomina INNER JOIN Grupo INNER JOIN Cargo INNER JOIN  TipoNomina INNER JOIN  " & _
'                                    "Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  " & _
'                                    "WHERE (Nomina.NumNomina = " & NumNomina & ") AND (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno + DetalleNomina.Antiguedad <> 0) AND (Departamento.CodDepartamento = '" & CodDepartamento & "') ORDER BY Nomina.NumNomina, Empleado.Nombre1"
'
'                      MDIPrimero.AdoConsulta.RecordSource = SQlReportes
'                      MDIPrimero.AdoConsulta.Refresh
'                      If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
'
'                        ArepColillasPago.AdoColillas.Source = SQlReportes
''                        ArepColillasPago.LblPeriodo.Caption = MDIPrimero.AdoConsulta.Recordset("FechaNominaINI") & " al " & Me.AdoBusca.Recordset("FechaNomina")
'                        ArepColillasPago.LblTitulo.Caption = Titulo
'                     End If
End Sub

Private Sub Detail_Format()
  Dim NetoPagar As Double, NumNomina As Double, CodEmpleado As String

    CodEmpleado = Field14.Text
    If Field17.Text <> "" Then
      NumNomina = Field17.Text
      Me.LblTotalDeduccion.Caption = Format(CDbl(Me.FldDeducciones.Text) - CDbl(Me.FldDescuento.Text) - CDbl(Me.FldAdelanto.Text), "#,##0.00")
    End If
    
    Me.LblIncentivo.Caption = "0.00"
    
If Quien = "ListadoNominas" Then
   Me.LblPeriodo.Caption = FechaInicio & " al " & FechaFinal
   Me.Label22.Caption = " Colilla de Pago Reimpresion"
End If

'Me.LblTotalDeduccion.Caption = Format(CDbl(Me.FldAdelanto.Text) + CDbl(Me.FldDeducciones.Text) + CDbl(Me.FldDescuento.Text), "#,##0.00")



' '/////////////////////////////BUSCO TODOS LOS INCENTIVOS QUE NO SON  EXCENTO /////////////////////////////////////////////
' MDIPrimero.AdoConsulta.ConnectionString = Conexion
' MDIPrimero.AdoConsulta.RecordSource = "SELECT MAX(DetalleIncentivo.NumIncentivo) AS NumIncentivo, SUM(DetalleIncentivo.Valor) AS Valor FROM DetalleIncentivo INNER JOIN Incentivo ON DetalleIncentivo.NumIncentivo = Incentivo.NumIncentivo INNER JOIN Empleado ON Incentivo.CodEmpleado = Empleado.CodEmpleado  " & _
'                                       "WHERE (Incentivo.CodTipoIncentivo <> N'14') AND (Empleado.CodEmpleado1 = '" & CodEmpleado & "') AND (DetalleIncentivo.NumNomina = " & NumNomina & ")"
' MDIPrimero.AdoConsulta.Refresh
' If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
'    Me.LblIncentivo.Caption = Format(MDIPrimero.AdoConsulta.Recordset("Valor"), "##,##0.00")
' End If
'
'
'  '/////////////////////////////BUSCO LOS INCENTIVOS /////////////////////////////////////////////
' MDIPrimero.AdoConsulta.ConnectionString = Conexion
' MDIPrimero.AdoConsulta.RecordSource = "SELECT MAX(DetalleIncentivo.NumIncentivo) AS NumIncentivo, SUM(DetalleIncentivo.Valor) AS Valor FROM DetalleIncentivo INNER JOIN Incentivo ON DetalleIncentivo.NumIncentivo = Incentivo.NumIncentivo INNER JOIN Empleado ON Incentivo.CodEmpleado = Empleado.CodEmpleado  " & _
'                                       "WHERE (Incentivo.CodTipoIncentivo = '14') AND (Empleado.CodEmpleado1 = '" & CodEmpleado & "') AND (DetalleIncentivo.NumNomina = " & NumNomina & ")"
' MDIPrimero.AdoConsulta.Refresh
' If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
'  Me.LblViaticos.Caption = Format(MDIPrimero.AdoConsulta.Recordset("Valor"), "##,##0.00")
' End If


End Sub

