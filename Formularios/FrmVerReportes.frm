VERSION 5.00
Begin VB.Form FrmVerReportes 
   Caption         =   "Resportes Zeus Nóminas"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data DtaReportes 
      Caption         =   "DtaReportes"
      Connect         =   "Access"
      DatabaseName    =   "C:\Zeus Nominas\Nominas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox CRViewer1 
      Height          =   7005
      Left            =   0
      ScaleHeight     =   6945
      ScaleWidth      =   5745
      TabIndex        =   0
      Top             =   0
      Width           =   5805
   End
End
Attribute VB_Name = "FrmVerReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim Report1 As New CRPRNomina
'Dim Report2 As New CRNomSubsidio
'Dim Report3 As New CRNomina
'Dim Report4 As New CRLstEmpleados
'Dim Report5 As New CRLstCargos
'Dim Report6 As New CRLstDepartamentos
'Dim Report7 As New CRLstTipoSubsidio
'Dim Report8 As New CRTipoIncentivo
'Dim Report9 As New CRTipoDeducciones
'Dim Report10 As New CRNomVacaciones
'Dim Report11 As New CRVerprestamo
'Dim Report12 As New CRDetalleIncentivoEmpleado
'Dim Report13 As New CrHistoricoDeduccionEmpleado
'Dim Report14 As New CRDetalleNomSubsidioEmpleado
'Dim Report15 As New CRBaja
'Dim Report16 As New CRDetallePrestamo
'Dim Report17 As New CRDetalleIncentivos
'Dim Report18 As New CrDetalleDeducciones
'Dim Report19 As New CrDetalleSubsidios

Private Sub Form_Load()
Dim SQLReportes As String
Dim Espacio As String

With Me.DtaReportes
   .DatabaseName = Ruta
   .Connect = Conexion
End With


MousePointer = 11
Espacio = " " & Chr(34) & " " & Chr(34)
Select Case NumReport

Case 1:

        SQLReportes = "SELECT Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalotrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado, "
        SQLReportes = SQLReportes + "[empleado].[nombre1] " & "+" & Espacio
        SQLReportes = SQLReportes + "+[empleado].[nombre2] " & "+" & Espacio
        SQLReportes = SQLReportes + "+[empleado].[apellido1] " & "+" & Espacio
        'SQLReportes = SQLReportes + "+[empleado].[apellido2] AS Nombre,Cargo.CodCargo, Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio FROM Nomina INNER JOIN ((Cargo INNER JOIN (TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina) ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) ON (TipoNomina.CodTipoNomina = Nomina.CodTipoNomina) AND (Nomina.NumNomina = DetalleNomina.NumNomina) WHERE Nomina.NumNomina= " & NumNomina & ""
        SQLReportes = SQLReportes + "+[empleado].[apellido2] AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio FROM Nomina INNER JOIN ((Cargo INNER JOIN (TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina) ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) ON (TipoNomina.CodTipoNomina = Nomina.CodTipoNomina) AND (Nomina.NumNomina = DetalleNomina.NumNomina) WHERE Nomina.NumNomina= " & NumNomina & ""
        DtaReportes.RecordSource = SQLReportes
        DtaReportes.Refresh
        
        Report1.ParameterFields.Parent.Database.SetDataSource DtaReportes.Recordset
        
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report1
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report1 = Nothing

Case 2:

        'SQLReportes = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Subsidio.NumSubsidio, TipoSubsidio.Subsidio, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.NumNominaSubsidio FROM TipoSubsidio INNER JOIN ((Empleado INNER JOIN Subsidio ON Empleado.CodEmpleado = Subsidio.CodEmpleado) INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio) ON TipoSubsidio.CodTipoSubsidio = Subsidio.CodTipoSubsidio Where DetalleSubsidio.NumNominaSubsidio = " & NumNominaSubsidio & " ORDER BY Empleado.CodEmpleado"
        SQLReportes = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Subsidio.NumSubsidio, TipoSubsidio.Subsidio, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.Descripcion, DetalleSubsidio.NumNominaSubsidio FROM TipoSubsidio INNER JOIN ((Empleado INNER JOIN Subsidio ON Empleado.CodEmpleado = Subsidio.CodEmpleado) INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio) ON TipoSubsidio.CodTipoSubsidio = Subsidio.CodTipoSubsidio Where DetalleSubsidio.NumNominaSubsidio = " & NumNominaSubsidio & " ORDER BY Empleado.CodEmpleado"
        DtaReportes.RecordSource = SQLReportes
        DtaReportes.Refresh
        
        Report2.ParameterFields.Parent.Database.SetDataSource DtaReportes.Recordset
        
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report2
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report2 = Nothing

Case 3:
        SQLReportes = "SELECT Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Empleado.CodGrupo, Grupo.Grupo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalOtrosIngresos, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalINATEC, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.FechaNominaINI, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado,"
        SQLReportes = SQLReportes + "[empleado].[nombre1] " & "+" & Espacio
        SQLReportes = SQLReportes + "+[empleado].[nombre2] " & "+" & Espacio
        SQLReportes = SQLReportes + "+[empleado].[apellido1] " & "+" & Espacio
        'SQLReportes = SQLReportes + "+[empleado].[apellido2] AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.OtrosIngresos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.INSSPatronal , DetalleNomina.INATEC, DetalleNomina.IRPatronal, DetalleNomina.Mes13 FROM Nomina INNER JOIN (Grupo INNER JOIN ((Cargo INNER JOIN (TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina) ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) ON Grupo.CodGrupo = Empleado.CodGrupo) ON (TipoNomina.CodTipoNomina = Nomina.CodTipoNomina) AND (Nomina.NumNomina = DetalleNomina.NumNomina) WHERE Nomina.NumNomina= " & NumNomina & ""
        SQLReportes = SQLReportes + "+[empleado].[apellido2] AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.HE, DetalleNomina.OtrosIngresos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.INSSPatronal, DetalleNomina.INATEC, DetalleNomina.IRPatronal, DetalleNomina.Mes13 FROM Nomina INNER JOIN (Grupo INNER JOIN ((Cargo INNER JOIN (TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina) ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) ON Grupo.CodGrupo = Empleado.CodGrupo) ON (TipoNomina.CodTipoNomina = Nomina.CodTipoNomina) AND (Nomina.NumNomina = DetalleNomina.NumNomina) WHERE Nomina.NumNomina= " & NumNomina & ""
        DtaReportes.RecordSource = SQLReportes
        DtaReportes.Refresh
        Report3.ParameterFields.Parent.Database.SetDataSource DtaReportes.Recordset
        
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report3
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report3 = Nothing


Case 4:
      
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report4
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report4 = Nothing
        
Case 5:
      
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report5
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report5 = Nothing

Case 6:
      
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report6
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report6 = Nothing

Case 7:
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report7
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report7 = Nothing
Case 8:
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report8
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report8 = Nothing
Case 9:
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report9
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report9 = Nothing
        
Case 10:

        SQLReportes = "SELECT NomVaca.NumNomVaca,  NomVaca.Fechaini,  NomVaca.FechaFin, DetalleNomVaca.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, [DetalleNomVaca]![SalarioMensual]*([DetalleNomVaca]![DiasAPagar]-[DetalleNomVaca]![DiasDescuento])/15 AS MontoAPagar FROM NomVaca INNER JOIN (Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado) ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca WHERE NomVaca.NumNomVaca=" & NumNomVaca & ""
        DtaReportes.RecordSource = SQLReportes
        DtaReportes.Refresh
        
        Report10.ParameterFields.Parent.Database.SetDataSource DtaReportes.Recordset
        
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report10
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report10 = Nothing

Case 11:

        SQLReportes = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Prestamo.Saldo, MovPrestamo.NumPrestamo, MovPrestamo.NumCuota, MovPrestamo.CuotaIgual, MovPrestamo.NumNomina, MovPrestamo.Cancelado FROM (Empleado INNER JOIN Prestamo ON Empleado.CodEmpleado = Prestamo.CodEmpleado) INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo WHERE MovPrestamo.NumPrestamo = " & NumPrestamo & " AND Empleado.CodEmpleado= '" & CodEmpleado & "'"
        DtaReportes.RecordSource = SQLReportes
        DtaReportes.Refresh
        
        Report11.ParameterFields.Parent.Database.SetDataSource DtaReportes.Recordset
        
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report11
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report11 = Nothing
Case 12:

        SQLReportes = "SELECT Incentivo.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Incentivo.CodTipoIncentivo, TipoIncentivo.Incentivo, DetalleIncentivo.NumIncentivo, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado, DetalleIncentivo.NumNomina FROM TipoIncentivo INNER JOIN ((Empleado INNER JOIN Incentivo ON Empleado.CodEmpleado = Incentivo.CodEmpleado) INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo) ON TipoIncentivo.CodTipoIncentivo = Incentivo.CodTipoIncentivo WHERE DetalleIncentivo.Pagado=True and Incentivo.CodEmpleado= '" & CodEmpleado & "'"
        DtaReportes.RecordSource = SQLReportes
        DtaReportes.Refresh
        
        Report11.ParameterFields.Parent.Database.SetDataSource DtaReportes.Recordset
        
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report12
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report12 = Nothing

Case 13:
       'SQLReportes = "SELECT Deduccion.NumDeduccion, Deduccion.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Deduccion.CodTipoDeduccion, TipoDeduccion.Deduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado, DetalleDeduccion.NumNomina FROM Empleado INNER JOIN (TipoDeduccion INNER JOIN (Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion) ON (TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion) AND (TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion)) ON Empleado.CodEmpleado = Deduccion.CodEmpleado WHERE DetalleDeduccion.Pagado=True AND Deduccion.CodEmpleado='" & CodEmpleado & "'"
       ' DtaReportes.RecordSource = SQLReportes
       ' DtaReportes.Refresh
        
       ' Report13.ParameterFields.Parent.Database.SetDataSource DtaReportes.Recordset
        
        'Screen.MousePointer = vbHourglass
       ' CRViewer1.ReportSource = Report13
       ' CRViewer1.ViewReport
       ' Screen.MousePointer = vbDefault
        
       ' Set Report13 = Nothing

Case 14:
        SQLReportes = "SELECT DetalleSubsidio.NumSubsidio, Subsidio.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Subsidio.CodTipoSubsidio, TipoSubsidio.Subsidio, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.Pagado, DetalleSubsidio.NumNominaSubsidio FROM TipoSubsidio INNER JOIN ((Empleado INNER JOIN Subsidio ON Empleado.CodEmpleado = Subsidio.CodEmpleado) INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio) ON TipoSubsidio.CodTipoSubsidio = Subsidio.CodTipoSubsidio WHERE DetalleSubsidio.Pagado=True And Subsidio.CodEmpleado='" & CodEmpleado & "'"
        DtaReportes.RecordSource = SQLReportes
        DtaReportes.Refresh
        
        Report14.ParameterFields.Parent.Database.SetDataSource DtaReportes.Recordset
        
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report14
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report14 = Nothing
        
Case 15:
        SQLReportes = "SELECT Bajas.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento,  Bajas.AnnosTrabajados, Bajas.MesesTrabajados, Bajas.DiasTrabajados, Bajas.MontoNomPropor, Bajas.FechaBaja, Bajas.MontoVaca, Bajas.Monto13Mes, Bajas.MontoAnosTrab, Bajas.MontoCargoConfianza, Bajas.MontoAntiguedad, Bajas.MotivoBaja, Bajas.TipoBaja, Bajas.Otro, Bajas.MontoOtro, Bajas.Prestamo, Bajas.Deducciones, [Bajas]![MontoVaca]+[Bajas]![Monto13Mes]+[Bajas]![MontoAnosTrab]+[Bajas]![MontoCargoConfianza]+[Bajas]![MontoAntiguedad]+[Bajas]![MontoOtro]-[Bajas]![Prestamo]-[Bajas]![Deducciones] AS TotalaPagar FROM (Departamento INNER JOIN (Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) ON Departamento.CodDepartamento = Empleado.CodDepartamento) INNER JOIN Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado Where Bajas.CodEmpleado = '" & CodEmpleado & "'"
        'SQLReportes = "SELECT Bajas.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Bajas.FechaBaja, Bajas.MontoVaca, Bajas.Monto13Mes, Bajas.MontoAnosTrab, Bajas.MontoCargoConfianza, Bajas.MontoAntiguedad, Bajas.MotivoBaja, Bajas.TipoBaja, Bajas.Otro, Bajas.MontoOtro, Bajas.Prestamo, Bajas.Deducciones, [Bajas]![MontoVaca]+[Bajas]![Monto13Mes]+[Bajas]![MontoAnosTrab]+[Bajas]![MontoCargoConfianza]+[Bajas]![MontoAntiguedad]+[Bajas]![MontoOtro]-[Bajas]![Prestamo]-[Bajas]![Deducciones] AS TotalaPagar FROM (Departamento INNER JOIN (Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) ON Departamento.CodDepartamento = Empleado.CodDepartamento) INNER JOIN Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado Where Bajas.CodEmpleado = '" & CodEmpleado & "'"
        DtaReportes.RecordSource = SQLReportes
        DtaReportes.Refresh
        
        Report15.ParameterFields.Parent.Database.SetDataSource DtaReportes.Recordset
        
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report15
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report15 = Nothing
        
Case 16:
        SQLReportes = "SELECT MovPrestamo.NumPrestamo, Prestamo.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Prestamo.Monto AS Prestamo, Prestamo.Saldo, Prestamo.CuotasIguales, MovPrestamo.NumCuota, MovPrestamo.Monto, MovPrestamo.Interes, MovPrestamo.CuotaIgual, MovPrestamo.SaldoCuota, MovPrestamo.Cancelado, MovPrestamo.NumNomina FROM (Empleado INNER JOIN Prestamo ON Empleado.CodEmpleado = Prestamo.CodEmpleado) INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo WHERE MovPrestamo.Cancelado=True AND MovPrestamo.NumNomina= " & NumNomina & ""
        DtaReportes.RecordSource = SQLReportes
        DtaReportes.Refresh
        
        Report16.ParameterFields.Parent.Database.SetDataSource DtaReportes.Recordset
        
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report16
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report16 = Nothing

Case 17:
        SQLReportes = "SELECT DetalleIncentivo.NumIncentivo, Incentivo.CodTipoIncentivo, TipoIncentivo.Incentivo, Incentivo.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado, DetalleIncentivo.NumNomina FROM Empleado INNER JOIN (TipoIncentivo INNER JOIN (Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo) ON TipoIncentivo.CodTipoIncentivo = Incentivo.CodTipoIncentivo) ON Empleado.CodEmpleado = Incentivo.CodEmpleado WHERE DetalleIncentivo.Pagado=True AND DetalleIncentivo.NumNomina= " & NumNomina & ""
        DtaReportes.RecordSource = SQLReportes
        DtaReportes.Refresh
        
        Report17.ParameterFields.Parent.Database.SetDataSource DtaReportes.Recordset
        
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report17
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report17 = Nothing

Case 18:
        'SQLReportes = "SELECT DetalleIncentivo.NumIncentivo, Incentivo.CodTipoIncentivo, TipoIncentivo.Incentivo, Incentivo.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado, DetalleIncentivo.NumNomina FROM Empleado INNER JOIN (TipoIncentivo INNER JOIN (Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo) ON TipoIncentivo.CodTipoIncentivo = Incentivo.CodTipoIncentivo) ON Empleado.CodEmpleado = Incentivo.CodEmpleado WHERE DetalleIncentivo.Pagado=True AND DetalleIncentivo.NumNomina= " & NumNomina & ""
        SQLReportes = "SELECT DetalleDeduccion.NumDeduccion, TipoDeduccion.Deduccion, Deduccion.CodTipoDeduccion, Deduccion.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado, DetalleDeduccion.NumNomina FROM TipoDeduccion INNER JOIN (Empleado INNER JOIN (Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion) ON Empleado.CodEmpleado = Deduccion.CodEmpleado) ON (TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion) AND (TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion) WHERE DetalleDeduccion.Pagado=True AND DetalleDeduccion.NumNomina= " & NumNomina & ""
        DtaReportes.RecordSource = SQLReportes
        DtaReportes.Refresh
        
        Report18.ParameterFields.Parent.Database.SetDataSource DtaReportes.Recordset
        
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report18
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report18 = Nothing

Case 19:
        SQLReportes = "SELECT DetalleSubsidio.NumSubsidio, TipoSubsidio.Subsidio, Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.Pagado, DetalleSubsidio.NumNominaSubsidio FROM TipoSubsidio INNER JOIN ((Empleado INNER JOIN Subsidio ON Empleado.CodEmpleado = Subsidio.CodEmpleado) INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio) ON TipoSubsidio.CodTipoSubsidio = Subsidio.CodTipoSubsidio WHERE DetalleSubsidio.Pagado=True AND DetalleSubsidio.NumNominaSubsidio= " & NumNominaSubsidio & ""
        DtaReportes.RecordSource = SQLReportes
        DtaReportes.Refresh
        
        Report19.ParameterFields.Parent.Database.SetDataSource DtaReportes.Recordset
        
        Screen.MousePointer = vbHourglass
        CRViewer1.ReportSource = Report19
        CRViewer1.ViewReport
        Screen.MousePointer = vbDefault
        
        Set Report19 = Nothing

End Select
MousePointer = 1
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
