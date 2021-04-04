Attribute VB_Name = "Registro"
Public objExcel As Excel.Application  ', ObjExcelFormato As Excel.CellFormat
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Nmes As Integer

Public Sub FMes(Mes As String)
  Select Case Mes
      Case "Enero"
         Nmes = 1
      Case "Febrero"
         Nmes = 2
      Case "Marzo"
           Nmes = 3
      Case "Abril"
               Nmes = 4
      Case "Mayo"
               Nmes = 5
      
      Case "Junio"
               Nmes = 6
      Case "Julio"
               Nmes = 7
      Case "Agosto"
               Nmes = 8
      Case "Septiembre"
               Nmes = 9
      Case "Octubre"
               Nmes = 10
      Case "Noviembre"
               Nmes = 11
      Case "Diciembre"
               Nmes = 12

  End Select
End Sub
Public Function SemanasPeriodos(Ano As Double, Mes As String, CodTipoNomina As String) As Double

Dim Sabados As Double
Dim SQlConsulta As String
Dim MesIni As String
Mes = Format(Mes, "0#")

 SQlConsulta = "SELECT * From Fecha_Planilla WHERE (año = " & Ano & ") AND (mes = '" & Mes & "') AND (CodTipoNomina = '" & CodTipoNomina & "')"
                        
 MDIPrimero.DtaConsulta.RecordSource = SQlConsulta
 MDIPrimero.DtaConsulta.Refresh
 If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
  MDIPrimero.DtaConsulta.Recordset.MoveLast
  Sabados = MDIPrimero.DtaConsulta.Recordset.RecordCount
 End If

 SemanasPeriodos = Sabados

End Function


Public Function SabadosMes(FechaNomina As Long) As Integer

Dim MiDiaSemana As Byte
Dim DiasMes As Byte
Dim Sabados As Byte
Dim Fecha As String

Fecha = CDate(FechaNomina)

DiasMes = Day(DateSerial(Year(FechaNomina), Month(FechaNomina) + 1, 0))


Sabados = 0

For i = 1 To DiasMes
 'construyo la fecha
 Fecha = Str(i) + "/" + Str(Month(FechaNomina)) + "/" + Str(Year(FechaNomina))
 Fecha = CDate(Fecha)
 MiDiaSemana = Weekday(Fecha, vbMonday)
 If MiDiaSemana = 6 Then
    Sabados = Sabados + 1
 End If
Next

SabadosMes = Sabados

End Function
Public Function ViernesMes(FechaNomina As Long) As Integer

Dim MiDiaSemana As Byte
Dim DiasMes As Byte
Dim Viernes As Byte
Dim Fecha As String

Fecha = CDate(FechaNomina)

DiasMes = Day(DateSerial(Year(FechaNomina), Month(FechaNomina) + 1, 0))


Viernes = 0

For i = 1 To DiasMes
 'construyo la fecha
 Fecha = Str(i) + "/" + Str(Month(FechaNomina)) + "/" + Str(Year(FechaNomina))
 Fecha = CDate(Fecha)
 MiDiaSemana = Weekday(Fecha, vbMonday)
 If MiDiaSemana = 5 Then
    Viernes = Viernes + 1
 End If
Next

ViernesMes = Viernes

End Function


Public Sub Limpia()
        FrmAnotaciones.ComboFaltas.Text = ""
        FrmAnotaciones.TxtJustificaFalta.Text = ""
        FrmAnotaciones.TxtDatosRecord.Text = ""
        FrmAnotaciones.TxtIdiomas.Text = ""
        FrmAnotaciones.TxtTelEmergencia.Text = ""
        FrmAnotaciones.TxtCursos.Text = ""
        FrmAnotaciones.TxtRazones.Text = ""
        FrmAnotaciones.TxtTrabAnteriores.Text = ""
        FrmAnotaciones.TxtRecomendaciones.Text = ""
        FrmAnotaciones.TxtSalida.Text = ""
        FrmAnotaciones.MaskFechaContratacion.Value = "__/__/____"
        FrmAnotaciones.MaskFechaFinaliza.Value = "__/__/____"
End Sub
 
 Public Sub GuardaRegistro()
  On Error GoTo TipoErrs
  FrmAnotaciones.DtaCurriculum.Recordset("CodEmpleado") = FrmAnotaciones.DBEmpleado.Text
  FrmAnotaciones.DtaCurriculum.Recordset("Faltas") = FrmAnotaciones.ComboFaltas.Text
  FrmAnotaciones.DtaCurriculum.Recordset.Fields("JustificacionFaltas") = FrmAnotaciones.TxtJustificaFalta.Text
  FrmAnotaciones.DtaCurriculum.Recordset.Fields("DatosRecord") = FrmAnotaciones.TxtDatosRecord.Text
  FrmAnotaciones.DtaCurriculum.Recordset.Fields("Idiomas") = FrmAnotaciones.TxtIdiomas.Text
  FrmAnotaciones.DtaCurriculum.Recordset.Fields("TelefonoCasoEmergencia") = FrmAnotaciones.TxtTelEmergencia.Text
  FrmAnotaciones.DtaCurriculum.Recordset.Fields("Cursos") = FrmAnotaciones.TxtCursos.Text
  FrmAnotaciones.DtaCurriculum.Recordset.Fields("RazonesContratacion") = FrmAnotaciones.TxtRazones.Text
  FrmAnotaciones.DtaCurriculum.Recordset.Fields("TrabajoAnterior") = FrmAnotaciones.TxtTrabAnteriores.Text
  FrmAnotaciones.DtaCurriculum.Recordset.Fields("Recomendaciones") = FrmAnotaciones.TxtRecomendaciones
  FrmAnotaciones.DtaCurriculum.Recordset.Fields("CausaSalida") = FrmAnotaciones.TxtSalida.Text
  FrmAnotaciones.DtaCurriculum.Recordset.Fields("FechaContratacion") = FrmAnotaciones.MaskFechaContratacion.Value
  FrmAnotaciones.DtaCurriculum.Recordset.Fields("FechaFinalizacion") = FrmAnotaciones.MaskFechaFinaliza.Value
  FrmAnotaciones.DtaCurriculum.Recordset.Update
Exit Sub
TipoErrs:
  ControlErrores
 End Sub

Public Sub ControlErrores()
     Select Case Err
        Case 13
             MsgBox "El Formato no es Correcto", vbInformation, "Error 13:Sistema de Nominas"
             error = 1
             
        Case 484
             MsgBox "Controlador de la Impresora no Disponible WIN.INI", vbInformation, "Error de Impresion 484:Sistema de Nominas"
        Case 483
             MsgBox "El controlador de la Impresora no admite esta Propiedad", vbInformation, "Error de Impresion 483:Sistema de Nominas"
        Case 482
             MsgBox "Error de la Impresora", vbInformation, "Error de Impresion 482:Sistema de Nominas"
        Case 396
              MsgBox "Imposible establecer la Propiedad dentro de la Pag.", vbInformation, "Error de Impresion 396:Sistema de Nominas"
        Case 91
              MsgBox "Error grabe ocurrido con la Base de Datos", vbInformation, "Error 91:Sistema de Nominas"
        Case 424
             MsgBox "No se Ha encontrado el Objeto Asociado", vbInformation, "Error 424:Sistema de Nominas"
        Case 53
             If Prueba = 1 Then
              MsgBox "El Archivo no se Ha encontrado", vbInformation, "Error 53:Sistema de Nominas"
             Else
              MsgBox "La Imagen Zw No se Ha encontrado", vbInformation, "Error53:Sistema de Nominas"
              Prueba = 0
             End If
        Case 380
             MsgBox "El tipo de Dato No es Correcto", vbCritical, "Error de Registro 380: Sistema de Nominas"
        Case 3163
             MsgBox "Desvordamiento de Datos", vbCritical, "Error de Registro 3163: Sistema de Nominas"
        Case 3021
             MsgBox "No Existe Registro Activo", vbCritical, "Error de Registro 3021: Sistema de Nominas"
        Case 3200
             MsgBox "No se Puede Eliminar Tiene Datos Relacionados", vbCritical, "Error de Registro 3200: Sistema de Nominas"
        Case 3315
            MsgBox "Debe Existir una Clave Primaria", vbCritical, "Error de Registro 3315: Sistema de Nominas"
        Case 3421
            MsgBox "No se Puede Agregar este Registro.", vbInformation, "Error de Registro 3421: Sistema de Nominas"
            error = 1
        Case 3201
            MsgBox "No se Puede Modificar el Registro Desde Aqui", vbCritical, "Error de Registro 3201: Sistema de Nominas"
        Case 440
            MsgBox "No coincide la data, con la intruccion", vbCritical, "Error de Datos 440: Sistema de Nominas"
        Case 68
            MsgBox Prompt:="La unidad no está preparada. Inserte un disco en la unidad.", Buttons:=vbExclamation, Title:="Sistema de Nominas"
            ' Restablece la ruta a la unidad anterior.
'            Drive1.Drive = Dir1.Path
            Exit Sub
        Case 52
             MsgBox Prompt:="La unidad no está preparada. Inserte un disco en la unidad.", Buttons:=vbExclamation, Title:="Sistema de Nominas"
        Case 70
             MsgBox Prompt:="La Unidad Esta protegida.Contra Escritura", Buttons:=vbExclamation, Title:="Sistema de Nominas"
        Case 71
             MsgBox Prompt:="La unidad no está preparada. Inserte un disco en la unidad.", Buttons:=vbExclamation, Title:="Sistema de Nominas"
        Case 76
           ' MsgBox Prompt:="No se Ha encontrado la Ruta Indicada.", Buttons:=vbExclamation, Title:="Sistema de Nominas"
        Case -2147217873
        
            MsgBox "Existes Registros Relacionados en otras Tablas", vbCritical, "Error de Registro -2147217873: Sistema de Nominas"
        Case Else
            MsgBox Prompt:="Error en la aplicación.Consulte al Soporte Tecnico", Buttons:=vbExclamation
    End Select
End Sub


Public Sub LimpiaEmpleado()
On Error GoTo TipoErrs

 frmEmpleado.TxtTelefono.Text = ""
 frmEmpleado.TxtCuentaBanco.Text = ""
 frmEmpleado.TxtNombre1.Text = ""
 frmEmpleado.TxtNombre2.Text = ""
 frmEmpleado.TxtApellido1.Text = ""
 frmEmpleado.TxtApellido2.Text = ""
 frmEmpleado.TxtDireccion.Text = ""
 frmEmpleado.TxtNacionalidad.Text = ""
 frmEmpleado.TxtNumCedula.Text = ""
 frmEmpleado.TxtCodPostal.Text = ""
 frmEmpleado.CmbSexo.Text = ""
 frmEmpleado.TxtNRuc.Text = ""
 frmEmpleado.TxtNInss.Text = ""
 frmEmpleado.DBCDepartamento.Text = ""
 frmEmpleado.DBCCargo.Text = ""
 frmEmpleado.CmbSindicalista.Text = ""
 frmEmpleado.TxtNumHijos.Text = ""
 frmEmpleado.TxtDiasDescuento.Text = "0"


 'Le Asigna una Imagen predefina cuando no exista el empleado.
 Destino = RutaFoto & frmEmpleado.DBCodigoEmpleado.Text & ".bmp"
 If (Dir(Destino) <> "") Then
   frmEmpleado.Image1.Picture = LoadPicture(Destino)
 Else
    Destino = RutaFoto + "Zw.bmp"
    frmEmpleado.Image1.Picture = LoadPicture(Destino)
 End If
 
 

 Exit Sub
TipoErrs:
  ControlErrores
 
End Sub


Public Sub GrabarInfNomina()
 On Error GoTo TipoErrs
 frmEmpleado.DtaInfNomina.Recordset.Fields("CodEmpleado") = frmEmpleado.DBCodigoEmpleado.Text
 frmEmpleado.DtaInfNomina.Recordset.Fields("TipoPago") = frmEmpleado.CmbTipoPago.Text
 frmEmpleado.DtaInfNomina.Recordset.Fields("SueldoPeriodo") = frmEmpleado.TxtSueldoPeriodo.Text
 frmEmpleado.DtaInfNomina.Recordset.Fields("TarifaHoraria") = frmEmpleado.TxtTarifaHoraria.Text
 frmEmpleado.DtaInfNomina.Recordset.Fields("SalarioMinimo") = frmEmpleado.CmbSalarioMinimo.Text
 frmEmpleado.DtaInfNomina.Recordset.Fields("ExentoInss") = frmEmpleado.CmbExentoInss.Text
 frmEmpleado.DtaInfNomina.Recordset.Fields("ExentoIr") = frmEmpleado.CmbExentoIr.Text
 frmEmpleado.DtaInfNomina.Recordset.Fields("PagoInssPatronal") = frmEmpleado.CmbPagoInssPatronal.Text
 frmEmpleado.DtaInfNomina.Recordset.Fields("CodTipoNomina") = frmEmpleado.TxtCodTipoNomina.Text
Exit Sub
TipoErrs:
  ControlErrores
End Sub

Public Sub GrabaInss()
On Error GoTo TipoErrs
 'FrmInssIR.DtaInss.Recordset.Fields("Categoria") = FrmInssIR.TxtCategoria.Text
 'FrmInssIR.DtaInss.Recordset.Fields("Desde") = FrmInssIR.TxtDesde.Text
 'FrmInssIR.DtaInss.Recordset.Fields("Hasta") = FrmInssIR.TxtHasta.Text
 'FrmInssIR.DtaInss.Recordset.Fields("MontoLaboral") = FrmInssIR.TxtMontolaboral
 'FrmInssIR.DtaInss.Recordset.Fields("MontoPatronal") = FrmInssIR.TxtMontoPatronal
Exit Sub
TipoErrs:
  ControlErrores
End Sub

Public Sub LimpiaHistorico()
frmEmpleado.MaskEdNacimiento.Value = Now
 frmEmpleado.MaskEdContrato.Value = Now
 frmEmpleado.DBCargoInicial.Text = ""
 frmEmpleado.DBCargoActual.Text = ""
 frmEmpleado.DBCargoAnterior.Text = ""
 frmEmpleado.TxtSueldoInicial.Text = "0.00"
 frmEmpleado.TxtSueldoAnterior.Text = "0.00"
 frmEmpleado.TxtSueldoActual.Text = "0.00"
 frmEmpleado.MaskEdBaja.Text = "__/__/____"
 frmEmpleado.TxtMotivoBaja.Text = ""
 frmEmpleado.MaskEdAumento.Text = "__/__/____"
 frmEmpleado.TxtMotivoAumento.Text = ""
 frmEmpleado.MaskEdSuspencion.Text = "__/__/____"
 frmEmpleado.MaskEdFinalSusp.Text = "__/__/____"
 frmEmpleado.TxtMotivoSuspencion.Text = ""
' frmEmpleado.TxtDebito.Text = "11111"
' frmEmpleado.TxtCredito.Text = "11111"
End Sub
Public Sub LimpiaInfNomina()
frmEmpleado.CmbTipoPago.Text = "Sueldo Fijo"
 frmEmpleado.TxtSueldoPeriodo.Text = "0.00"
 frmEmpleado.TxtTarifaHoraria.Text = "0.00"
 frmEmpleado.CmbSalarioMinimo.Text = ""
 frmEmpleado.CmbExentoInss.Text = "No"
 frmEmpleado.CmbExentoIr.Text = "No"
 frmEmpleado.TxtOtrosIngresos.Text = "0.00"
 frmEmpleado.TxtDescripOtrIngre.Text = ""
 frmEmpleado.CmbPagoInssPatronal.Text = "Si"
 frmEmpleado.TxtCodTipoNomina.Text = ""
 frmEmpleado.DBCTipoNomina = ""
End Sub


Public Sub PreparaSalida()
 If Evaluar = True Then
 Salida = False
Else
 Salida = True
End If
End Sub
Public Sub LeeTecla()
Select Case Lectura
       Case 42
          Lectura = "*"
       Case 48
         Lectura = "0"
       Case 49
         Lectura = "1"
       Case 50
           Lectura = "2"
       Case 51
            Lectura = "3"
       Case 52
            Lectura = 4
       Case 53
             Lectura = 5
        Case 54
             Lectura = 6
        Case 55
             Lectura = 7
        Case 56
             Lectura = 8
        Case 57
              Lectura = 9
        Case 59
              Lectura = "Ñ"
        Case 47
              Lectura = "/"
        Case 81
             Lectura = "Q"
         Case 87
             Lectura = "W"
         Case 69
             Lectura = "E"
         Case 82
             Lectura = "R"
        Case 84
              Lectura = "T"
        Case 89
              Lectura = "Y"
        Case 85
              Lectura = "U"
                
         Case 112
             Lectura = "P"
         Case 113
             Lectura = "Q"
         Case 119
             Lectura = "W"
         Case 101
             Lectura = "E"
         Case 114
             Lectura = "R"
         Case 116
             Lectura = "T"
         Case 121
             Lectura = "Y"
         Case 117
             Lectura = "U"
         Case 105
             Lectura = "I"
         Case 111
             Lectura = "O"
         Case 97
             Lectura = "A"
         Case 115
             Lectura = "S"
         Case 100
             Lectura = "D"
         Case 102
             Lectura = "F"
         Case 103
             Lectura = "G"
         Case 104
             Lectura = "H"
         Case 106
             Lectura = "J"
         Case 107
             Lectura = "K"
         Case 108
             Lectura = "L"
         Case 122
             Lectura = "Z"
         Case 120
             Lectura = "X"
         Case 99
             Lectura = "C"
         Case 118
             Lectura = "V"
         Case 98
             Lectura = "B"
         Case 110
             Lectura = "N"
         Case 109
             Lectura = "M"
        Case 73
             Lectura = "I"
         Case 79
             Lectura = "O"
         Case 80
             Lectura = "P"
         Case 65
             Lectura = "A"
        Case 83
              Lectura = "S"
        Case 68
              Lectura = "D"
        Case 70
              Lectura = "F"
        Case 71
             Lectura = "G"
         Case 72
             Lectura = "H"
         Case 74
             Lectura = "J"
         Case 75
             Lectura = "K"
         Case 76
             Lectura = "L"
         Case 90
             Lectura = "Z"
        Case 88
              Lectura = "X"
        Case 67
              Lectura = "C"
        Case 86
              Lectura = "V"
        Case 66
             Lectura = "B"
         Case 78
             Lectura = "N"
         Case 77
             Lectura = "M"
End Select
End Sub

Public Sub ValidaSalida(Var As Variant)
If Salida Then
  On Error GoTo TipoErrs
'se utiliza para cuando el usario sale con la x de windows y poder validar
k% = MsgBox("Guardar Cambios " & Var & "?", vbYesNo, "Sistema de Nominas")
  If k% <> 6 Then
    Contesta = False
  Else
    Salida = True
    Contesta = True
  End If
  Exit Sub
Else
  Contesta = False
End If
Salida = False
Exit Sub
TipoErrs:
 ControlErrores
 
End Sub

Public Sub Check()
 'FrmEditaNiveles.ChLeer.Enabled = True
 FrmEditaNiveles.ChGrabar.Enabled = True
 FrmEditaNiveles.ChEliminar.Enabled = True
 FrmEditaNiveles.ChAbrir.Enabled = False
 'FrmEditaNiveles.ChCopiaCia.Enabled = False
 'FrmEditaNiveles.ChBorraCia.Enabled = False
End Sub

Public Sub GrabaNiveles()
On Error GoTo TipoErrs
'Select Case FrmEditaNiveles.ListAcceso.Text
'       Case "Compañia"
'             If FrmEditaNiveles.ChCopiaCia.Value = 1 Then
'               FrmEditaNiveles.DtaNacceso.Recordset.Fields("CopiaCia") = "s"
'             Else
'               FrmEditaNiveles.DtaNacceso.Recordset.Fields("CopiaCia") = "n"
'             End If
'             If FrmEditaNiveles.ChBorraCia.Value = 1 Then
'               FrmEditaNiveles.DtaNacceso.Recordset.Fields("BorraCia") = "s"
'             Else
'               FrmEditaNiveles.DtaNacceso.Recordset.Fields("BorraCia") = "n"
'             End If
'             If FrmEditaNiveles.ChAbrir.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerCompañia") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerCompañia") = "n"
'             End If
'       Case "Selecciona Compañia"
'             If FrmEditaNiveles.ChAbrir.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("SeleccionarCia") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("SeleccionarCia") = "n"
'             End If
'       Case "Editar Niveles"
'             If FrmEditaNiveles.ChAbrir.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerEditaNiveles") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerEditaNiveles") = "n"
'             End If
'       Case "Registro Empleados"
'             If FrmEditaNiveles.ChGrabar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GEmpleado") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GEmpleado") = "n"
'             End If
'             If FrmEditaNiveles.ChEliminar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BEmpleado") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BEmpleado") = "n"
'             End If
''             If FrmEditaNiveles.ChLeer.Value = 1 Then
''                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerEmpleado") = "s"
''             Else
''                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerEmpleado") = "n"
''             End If
'       Case "Tabla Anotaciones"
'             If FrmEditaNiveles.ChGrabar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GAnotaciones") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GAnotaciones") = "n"
'             End If
'             If FrmEditaNiveles.ChEliminar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BAnotaciones") = "s"
'             Else
'                 FrmEditaNiveles.DtaNacceso.Recordset.Fields("BAnotaciones") = "n"
'             End If
''             If FrmEditaNiveles.ChLeer.Value = 1 Then
''                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerAnotaciones") = "s"
''             Else
''                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerAnotaciones") = "n"
''             End If
'       Case "Tabla Departamentos"
'             If FrmEditaNiveles.ChGrabar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GDepartamento") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GDepartamento") = "n"
'             End If
'             If FrmEditaNiveles.ChEliminar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BDepartamento") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BDepartamento") = "n"
'             End If
''             If FrmEditaNiveles.ChLeer.Value = 1 Then
''                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerDepartamento") = "s"
''             Else
''                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerDepartamento") = "n"
''             End If
'       Case "Tabla Cargo"
'             If FrmEditaNiveles.ChGrabar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GCargo") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GCargo") = "n"
'             End If
'             If FrmEditaNiveles.ChEliminar.Value = 1 Then
'                DtaNacceso.Recordset.Fields("BCargo") = "s"
'             Else
'               FrmEditaNiveles.DtaNacceso.Recordset.Fields("BCargo") = "n"
'             End If
''             If FrmEditaNiveles.ChLeer.Value = 1 Then
''                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerCargo") = "s"
''             Else
''                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerCargo") = "n"
''             End If
'       Case "Tabla Tipo Incapacidad"
'             If FrmEditaNiveles.ChGrabar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GTipoIncapacidad") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GTipoIncapacidad") = "n"
'             End If
'             If FrmEditaNiveles.ChEliminar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BTipoIncapacidad") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BTipoIncapacidad") = "n"
''             End If
''             If FrmEditaNiveles.ChLeer.Value = 1 Then
''                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerTipoIncapacidad") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerTipoIncapacidad") = "n"
'             End If
'       Case "Tabla Incapacidad"
'             If FrmEditaNiveles.ChGrabar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GIncapacidad") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GIncapacidad") = "n"
'             End If
'             If FrmEditaNiveles.ChEliminar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BIncapacidad") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BIncapacidad") = "n"
'             End If
'             If FrmEditaNiveles.ChLeer.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerIncapacidad") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerIncapacidad") = "n"
'             End If
      
'       Case "Tabla Prestamos"
'             If FrmEditaNiveles.ChGrabar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GPrestamos") = "s"
'             Else
'                 FrmEditaNiveles.DtaNacceso.Recordset.Fields("GPrestamos") = "n"
'             End If
'             If FrmEditaNiveles.ChEliminar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BPrestamos") = "s"
'             Else
'               FrmEditaNiveles.DtaNacceso.Recordset.Fields("BPrestamos") = "n"
'             End If
'             If FrmEditaNiveles.ChLeer.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerPrestamos") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerPrestamos") = "n"
'             End If
'       Case "Tabla Tipo Nomina"
'             If FrmEditaNiveles.ChGrabar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GTipoNomina") = "s"
'             Else
'               FrmEditaNiveles.DtaNacceso.Recordset.Fields("GTipoNomina") = "n"
'             End If
'             If FrmEditaNiveles.ChEliminar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BTipoNomina") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BTipoNomina") = "n"
'             End If
'             If FrmEditaNiveles.ChLeer.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerTipoNomina") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerTipoNomina") = "n"
'             End If
'       Case "Tabla INSS/IR"
'            If FrmEditaNiveles.ChGrabar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GInss") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GInss") = "n"
'             End If
'             If FrmEditaNiveles.ChEliminar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BInss") = "s"
'             Else
'               FrmEditaNiveles.DtaNacceso.Recordset.Fields("BInss") = "n"
'             End If
'             If FrmEditaNiveles.ChLeer.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerInss") = "s"
'             Else
'               FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerInss") = "n"
'             End If
'
'       Case "Claves de Usuario"
'             If FrmEditaNiveles.ChLeer.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerUsuarios") = "s"
'             Else
'               FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerUsuarios") = "n"
'             End If
'       Case "Registro Moneda"
'             If FrmEditaNiveles.ChGrabar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("GRegistroMoneda") = "s"
'             Else
'               FrmEditaNiveles.DtaNacceso.Recordset.Fields("GRegistroMoneda") = "n"
'             End If
'             If FrmEditaNiveles.ChEliminar.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BRegistroMoneda") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("BRegistroMoneda") = "n"
'             End If
'             If FrmEditaNiveles.ChLeer.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerRegistroMoneda") = "s"
'             Else
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerRegistroMoneda") = "n"
'             End If
'       Case "Abrir/Cerrar Respaldo"
'            If FrmEditaNiveles.ChAbrir.Value = 1 Then
'                FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerRespaldo") = "s"
'            Else
'              FrmEditaNiveles.DtaNacceso.Recordset.Fields("VerRespaldo") = "n"
'            End If
      
'       End Select
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Public Sub Verifica()
' Select Case FrmEditaNiveles.ListAcceso.Text
'        Case "Compañia"
'              If FrmEditaNiveles.DtaNacceso.Recordset.CopiaCia = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.CopiaCia) Then
'               FrmEditaNiveles.ChCopiaCia.Value = 1
'              Else
'                FrmEditaNiveles.ChCopiaCia.Value = 0
'              End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.BorraCia = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.BorraCia) Then
'               FrmEditaNiveles.ChBorraCia.Value = 1
'             Else
'               FrmEditaNiveles.ChBorraCia.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerCompañia = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerCompañia) Then
'                FrmEditaNiveles.ChAbrir.Value = 1
'             Else
'                FrmEditaNiveles.ChAbrir.Value = 0
'             End If
'             FrmEditaNiveles.ChAbrir.Enabled = True
'             FrmEditaNiveles.ChCopiaCia.Enabled = True
'             FrmEditaNiveles.ChBorraCia.Enabled = True
'            FrmEditaNiveles.ChLeer.Enabled = False
'            FrmEditaNiveles.ChGrabar.Enabled = False
'            FrmEditaNiveles.ChEliminar.Enabled = False
'            Exit Sub
'        Case "Selecciona Compañia"
'             If FrmEditaNiveles.DtaNacceso.Recordset.SeleccionarCia = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.SeleccionarCia) Then
'                FrmEditaNiveles.ChAbrir.Value = 1
'             Else
'                FrmEditaNiveles.ChAbrir.Value = 0
'             End If
'             FrmEditaNiveles.ChLeer.Enabled = False
'             FrmEditaNiveles.ChGrabar.Enabled = False
'             FrmEditaNiveles.ChEliminar.Enabled = False
'             FrmEditaNiveles.ChAbrir.Enabled = True
'             FrmEditaNiveles.ChCopiaCia.Enabled = False
'             FrmEditaNiveles.ChBorraCia.Enabled = False
'       Case "Editar Niveles"
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerEditaNiveles = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerEditaNiveles) Then
'                FrmEditaNiveles.ChAbrir.Value = 1
'             Else
'                FrmEditaNiveles.ChAbrir.Value = 0
'             End If
'             FrmEditaNiveles.ChLeer.Enabled = False
'             FrmEditaNiveles.ChGrabar.Enabled = False
'             FrmEditaNiveles.ChEliminar.Enabled = False
'             FrmEditaNiveles.ChAbrir.Enabled = True
'             FrmEditaNiveles.ChCopiaCia.Enabled = False
'             FrmEditaNiveles.ChBorraCia.Enabled = False
'       Case "Registro Empleados"
'             If FrmEditaNiveles.DtaNacceso.Recordset.GEmpleado = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.GEmpleado) Then
'                FrmEditaNiveles.ChGrabar.Value = 1
'             Else
'                FrmEditaNiveles.ChGrabar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.BEmpleado = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.BEmpleado) Then
'             FrmEditaNiveles.ChEliminar.Value = 1
'             Else
'             FrmEditaNiveles.ChEliminar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerEmpleado = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerEmpleado) Then
'              FrmEditaNiveles.ChLeer.Value = 1
'             Else
'              FrmEditaNiveles.ChLeer.Value = 0
'             End If
'             Check
'       Case "Tabla Anotaciones"
'             If FrmEditaNiveles.DtaNacceso.Recordset.GAnotaciones = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.GAnotaciones) Then
'                FrmEditaNiveles.ChGrabar.Value = 1
'             Else
'                FrmEditaNiveles.ChGrabar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.BAnotaciones = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.BAnotaciones) Then
'             FrmEditaNiveles.ChEliminar.Value = 1
'             Else
'             FrmEditaNiveles.ChEliminar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerAnotaciones = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerAnotaciones) Then
'              FrmEditaNiveles.ChLeer.Value = 1
'             Else
'              FrmEditaNiveles.ChLeer.Value = 0
'             End If
'             Check
'       Case "Tabla Departamentos"
'             If FrmEditaNiveles.DtaNacceso.Recordset.GDepartamento = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.GDepartamento) Then
'                FrmEditaNiveles.ChGrabar.Value = 1
'             Else
'                FrmEditaNiveles.ChGrabar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.BDepartamento = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.BDepartamento) Then
'             FrmEditaNiveles.ChEliminar.Value = 1
'             Else
'             FrmEditaNiveles.ChEliminar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerDepartamento = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerDepartamento) Then
'              FrmEditaNiveles.ChLeer.Value = 1
'             Else
'              FrmEditaNiveles.ChLeer.Value = 0
'             End If
'             Check
'       Case "Tabla Cargo"
'
'
'             If FrmEditaNiveles.DtaNacceso.Recordset.GCargo = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.GCargo) Then
'               FrmEditaNiveles.ChGrabar.Value = 1
'             Else
'                FrmEditaNiveles.ChGrabar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.BCargo = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.BCargo) Then
'             FrmEditaNiveles.ChEliminar.Value = 1
'             Else
'             FrmEditaNiveles.ChEliminar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerCargo = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerCargo) Then
'              FrmEditaNiveles.ChLeer.Value = 1
'             Else
'              FrmEditaNiveles.ChLeer.Value = 0
'             End If
'             Check
'       Case "Tabla Tipo Incapacidad"
'              If FrmEditaNiveles.DtaNacceso.Recordset.GTipoIncapacidad = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.GTipoIncapacidad) Then
'                FrmEditaNiveles.ChGrabar.Value = 1
'             Else
'                FrmEditaNiveles.ChGrabar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.BTipoIncapacidad = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.BTipoIncapacidad) Then
'             FrmEditaNiveles.ChEliminar.Value = 1
'             Else
'             FrmEditaNiveles.ChEliminar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerTipoIncapacidad = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerTipoIncapacidad) Then
'              FrmEditaNiveles.ChLeer.Value = 1
'             Else
'              FrmEditaNiveles.ChLeer.Value = 0
'             End If
'             Check
'       Case "Tabla Incapacidad"
'             If FrmEditaNiveles.DtaNacceso.Recordset.GIncapacidad = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.GIncapacidad) Then
'                FrmEditaNiveles.ChGrabar.Value = 1
'             Else
'                FrmEditaNiveles.ChGrabar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.BIncapacidad = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.BIncapacidad) Then
'             FrmEditaNiveles.ChEliminar.Value = 1
'             Else
'             FrmEditaNiveles.ChEliminar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerIncapacidad = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerIncapacidad) Then
'              FrmEditaNiveles.ChLeer.Value = 1
'             Else
'              FrmEditaNiveles.ChLeer.Value = 0
'             End If
'             Check
'       Case "Tabla Prestamos"
'             If FrmEditaNiveles.DtaNacceso.Recordset.GPrestamos = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.GPrestamos) Then
'                FrmEditaNiveles.ChGrabar.Value = 1
'             Else
'                FrmEditaNiveles.ChGrabar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.BPrestamos = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.BPrestamos) Then
'             FrmEditaNiveles.ChEliminar.Value = 1
'             Else
'             FrmEditaNiveles.ChEliminar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerPrestamos = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerPrestamos) Then
'              FrmEditaNiveles.ChLeer.Value = 1
'             Else
'              FrmEditaNiveles.ChLeer.Value = 0
'             End If
'             Check
'       Case "Tabla Tipo Nomina"
'              If FrmEditaNiveles.DtaNacceso.Recordset.GTipoNomina = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.GTipoNomina) Then
'                FrmEditaNiveles.ChGrabar.Value = 1
'             Else
'                FrmEditaNiveles.ChGrabar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.BTipoNomina = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.BTipoNomina) Then
'             FrmEditaNiveles.ChEliminar.Value = 1
'             Else
'             FrmEditaNiveles.ChEliminar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerTipoNomina = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerTipoNomina) Then
'              FrmEditaNiveles.ChLeer.Value = 1
'             Else
'              FrmEditaNiveles.ChLeer.Value = 0
'             End If
'             Check
'       Case "Tabla INSS/IR"
'            If FrmEditaNiveles.DtaNacceso.Recordset.GInss = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.GInss) Then
'                FrmEditaNiveles.ChGrabar.Value = 1
'             Else
'                FrmEditaNiveles.ChGrabar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.BInss = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.BInss) Then
'             FrmEditaNiveles.ChEliminar.Value = 1
'             Else
'             FrmEditaNiveles.ChEliminar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerInss = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerInss) Then
'              FrmEditaNiveles.ChLeer.Value = 1
'             Else
'              FrmEditaNiveles.ChLeer.Value = 0
'             End If
'            Check
'
'       Case "Claves de Usuario"
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerUsuarios = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerUsuarios) Then
'                FrmEditaNiveles.ChAbrir.Value = 1
'             Else
'                FrmEditaNiveles.ChAbrir.Value = 0
'             End If
'             FrmEditaNiveles.ChLeer.Enabled = True
'             FrmEditaNiveles.ChGrabar.Enabled = False
'             FrmEditaNiveles.ChEliminar.Enabled = False
'             FrmEditaNiveles.ChAbrir.Enabled = False
'             FrmEditaNiveles.ChCopiaCia.Enabled = False
'             FrmEditaNiveles.ChBorraCia.Enabled = False
'       Case "Registro Moneda"
'             If FrmEditaNiveles.DtaNacceso.Recordset.GRegistroMoneda = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.GRegistroMoneda) Then
'                FrmEditaNiveles.ChGrabar.Value = 1
'             Else
'                FrmEditaNiveles.ChGrabar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.BRegistroMoneda = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.BRegistroMoneda) Then
'             FrmEditaNiveles.ChEliminar.Value = 1
'             Else
'             FrmEditaNiveles.ChEliminar.Value = 0
'             End If
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerRegistroMoneda = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerRegistroMoneda) Then
'              FrmEditaNiveles.ChLeer.Value = 1
'             Else
'              FrmEditaNiveles.ChLeer.Value = 0
'             End If
'             Check
'       Case "Abrir/Cerrar Respaldo"
'             If FrmEditaNiveles.DtaNacceso.Recordset.VerRespaldo = "s" Or IsNull(FrmEditaNiveles.DtaNacceso.Recordset.VerRespaldo) Then
'                FrmEditaNiveles.ChAbrir.Value = 1
'             Else
'                FrmEditaNiveles.ChAbrir.Value = 0
'             End If
'             FrmEditaNiveles.ChLeer.Enabled = False
'             FrmEditaNiveles.ChGrabar.Enabled = False
'             FrmEditaNiveles.ChEliminar.Enabled = False
'             FrmEditaNiveles.ChAbrir.Enabled = True
'             FrmEditaNiveles.ChCopiaCia.Enabled = False
'             FrmEditaNiveles.ChBorraCia.Enabled = False
'       End Select
End Sub

Public Sub Activa()
'Select Case FrmEditaNiveles.ListAcceso.Text
'        Case "Compañia"
'               FrmEditaNiveles.ChAbrir.Enabled = True
'             FrmEditaNiveles.ChCopiaCia.Enabled = True
'             FrmEditaNiveles.ChBorraCia.Enabled = True
'            FrmEditaNiveles.ChLeer.Enabled = False
'            FrmEditaNiveles.ChGrabar.Enabled = False
'            FrmEditaNiveles.ChEliminar.Enabled = False
'            Exit Sub
'        Case "Selecciona Compañia"
'             FrmEditaNiveles.ChLeer.Enabled = False
'             FrmEditaNiveles.ChGrabar.Enabled = False
'             FrmEditaNiveles.ChEliminar.Enabled = False
'             FrmEditaNiveles.ChAbrir.Enabled = True
'             FrmEditaNiveles.ChCopiaCia.Enabled = False
'             FrmEditaNiveles.ChBorraCia.Enabled = False
'       Case "Editar Niveles"
'             FrmEditaNiveles.ChLeer.Enabled = False
'             FrmEditaNiveles.ChGrabar.Enabled = False
'             FrmEditaNiveles.ChEliminar.Enabled = False
'             FrmEditaNiveles.ChAbrir.Enabled = True
'             FrmEditaNiveles.ChCopiaCia.Enabled = False
'             FrmEditaNiveles.ChBorraCia.Enabled = False
'       Case "Registro Empleados"
'             Check
'       Case "Tabla Anotaciones"
'
'              Check
'       Case "Tabla Departamentos"
'             Check
'       Case "Tabla Cargo"
'             Check
'       Case "Tabla Tipo Incapacidad"
'              Check
'       Case "Tabla Incapacidad"
'             Check
'       Case "Tabla Prestamos"
'             Check
'       Case "Tabla Tipo Nomina"
'              Check
'       Case "Tabla INSS/IR"
'             Check
'
'       Case "Claves de Usuario"
'             FrmEditaNiveles.ChLeer.Enabled = True
'             FrmEditaNiveles.ChGrabar.Enabled = False
'             FrmEditaNiveles.ChEliminar.Enabled = False
'             FrmEditaNiveles.ChAbrir.Enabled = False
'             FrmEditaNiveles.ChCopiaCia.Enabled = False
'             FrmEditaNiveles.ChBorraCia.Enabled = False
'       Case "Registro Moneda"
'             Check
'       Case "Abrir/Cerrar Respaldo"
'             FrmEditaNiveles.ChLeer.Enabled = True
'             FrmEditaNiveles.ChGrabar.Enabled = False
'             FrmEditaNiveles.ChEliminar.Enabled = False
'             FrmEditaNiveles.ChAbrir.Enabled = False
'             FrmEditaNiveles.ChCopiaCia.Enabled = False
'             FrmEditaNiveles.ChBorraCia.Enabled = False
'       End Select
End Sub


 
Public Sub Otorga()
    
CopiaCia = True
BorraCia = True
VerCia = True
SelecCia = True
GEmpleado = True
BEmpleado = True
VerEmpleado = True
GAnotaciones = True
BAnotaciones = True
VerAnotaciones = True
GDepartamento = True
BDepartamento = True
VerDepartamento = True
GCargo = True
BCargo = True
VerCargo = True
GTipoIncapacidad = True
BTipoIncapacidad = True
VerTipoIncapacidad = True
GIncapacidad = True
BIncapacidad = True
VerIncapacidad = True
GPrestamos = True
BPrestamos = True
VerPrestamos = True
GTipoNomina = True
BTipoNomina = True
 VerTipoNomina = True
GInss = True
BInss = True
VerInss = True
VerUsuarios = True
GRegistroMoneda = True
BRegistroMoneda = True
VerRegistroMoneda = True
VerRespaldo = True
End Sub

Public Sub CreaArchivo()
FrmExporta.DtaHistorico.Refresh
 Do While Not FrmExporta.DtaHistorico.Recordset.EOF
  FrmExporta.DtaEmpleado.Refresh
      Do While Not FrmExporta.DtaEmpleado.Recordset.EOF
           If FrmExporta.DtaEmpleado.Recordset("CodEmpleado") = FrmExporta.DtaHistorico.Recordset("CodEmpleado") Then
            Print #1, "A"; FrmExporta.DtaHistorico.Recordset("CuentaDebito"); Tab(21); FrmExporta.DtaEmpleado.Recordset("Nombre1") & " " & FrmExporta.DtaEmpleado.Recordset("Nombre2") & " " & FrmExporta.DtaEmpleado.Recordset("Apellido1"); Tab(56); "10"
           End If
        FrmExporta.DtaEmpleado.Recordset.MoveNext
      Loop
    FrmExporta.DtaHistorico.Recordset.MoveNext
  Loop
End Sub

Public Sub Backad()
On Error GoTo TipoErrs
       Origen = "c:\Sistema de Nominas\Nominas.mdb"
         
        frmBackup.CmdProcesar.Enabled = False
        frmBackup.CmdCerrar.Enabled = False
        
        If Mid$(frmBackup.File1.Path, Len(frmBackup.File1.Path)) = "\" Then
           Destino = frmBackup.Dir1.Path & "Zeus.Zn"
         Else
           Destino = frmBackup.Dir1.Path & "\" & "Zeus.Zn"
         End If
        FileCopy Origen, Destino
        Cadena = "El respaldo se ha Creado" & vbLf
        Cadena = Cadena & "Satisfactoriamente.."
        R% = MsgBox(Cadena, vbExclamation, "Sistema de Nominas")
               
Exit Sub
TipoErrs:
ControlErrores
End Sub

Public Sub UbicaDepartamento()
FrmDepartamentos.DBCodigo.Text = ""
FrmDepartamentos.txtNombre.Text = ""
FrmDepartamentos.DtaDepartamento.Recordset.MoveLast
FrmDepartamentos.DtaDepartamento.Recordset.MovePrevious
FrmDepartamentos.CmdAnterior.Enabled = True
FrmDepartamentos.CmdSiguiente.Enabled = True
FrmDepartamentos.CmdPrimero.Enabled = True
FrmDepartamentos.CmdUltimo.Enabled = True
FrmDepartamentos.CmdBorrar.Enabled = True
End Sub

Public Sub UbicaEmpleado()
frmEmpleado.CmdAnterior.Enabled = True
frmEmpleado.CmdSiguiente.Enabled = True
frmEmpleado.CmdPrimero.Enabled = True
frmEmpleado.CmdUltimo.Enabled = True
frmEmpleado.CmdBorrar.Enabled = True
End Sub
Public Function Decrypt(Frase As String) As String
Dim Ilen As Integer, X As Integer
Dim sFrase As String, sCurrent As String, sNew As String
Ilen = Len(Frase)
For X = 1 To Ilen
    sCurrent = Mid$(Frase, X, 1)
    sNew = Chr$(Asc(sCurrent) - 110)
    sFrase = sFrase & sNew
Next
Decrypt = sFrase
End Function

Public Function Encrypt(Frase As String) As String
Dim Ilen As Integer, X As Integer
Dim sFrase As String, sCurrent As String, sNew As String
Ilen = Len(Frase)
For X = 1 To Ilen
    sCurrent = Mid$(Frase, X, 1)
    sNew = Chr$(Asc(sCurrent) + 110)
    sFrase = sFrase & sNew
Next
Encrypt = sFrase
End Function



Public Function ConvertirMes(Mes As Integer) As String
 Select Case Mes
  Case 1
     Convertir = "Enero"
  Case 2
     Convertir = "Febrero"
  Case 3
     Convertir = "Marzo"
  Case 4
     Convertir = "Abril"
  Case 5
    Convertir = "Mayo"
  Case 6
    Convertir = "Junio"
  Case 7
    Convertir = "Julio"
  Case 8
   Convertir = "Agosto"
  Case 9
   Convertir = "Septiembre"
  Case 10
   Convertir = "Octubre"
  Case 11
    Convertir = "Noviembre"
  Case 12
    Convertir = "Diciembre"
 End Select
 
 ConvertirMes = Convertir
End Function

Public Function Inicio_Excel() As Boolean
Dim i As Integer
Dim j As Integer

Set objExcel = New Excel.Application
 
objExcel.Visible = True 'lo hacemos visible
objExcel.SheetsInNewWorkbook = 1 'decimos cuantas hojas queremos en el nuevo documento
objExcel.Workbooks.Add ' añadimos el objeto al workbook

End Function


Public Function Formato_Excel(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

With objExcel.ActiveSheet
        
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos - 1)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 9)).Font.Bold = True
        
    For i = 1 To Num_Campos - 1 Step 1
        .Cells(3, i) = Nombre_Campos(i)
    Next i
        'hasta aki pa colocar los titulos
        
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
       
        .Columns("A").ColumnWidth = 10
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 45
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 20
        
   
    
End With
End Function

Public Sub Llenado()
Dim SQlIncentivos As String, SQlDeducciones As String, SqlDetallePrestamo As String, SQlPrestamo As String, SqlDetalleSubsidio As String
Dim Salario As Boolean
Dim CodEmpleado1 As String
Dim numeroPrestamo As Double


 Evaluar = True
 'Al ejecutar algun cambio en el combo actualizo el nombre del Empleado
   frmEmpleado.MousePointer = 11

 LimpiaEmpleado
 LimpiaHistorico
 LimpiaInfNomina
 'LimpiaInfNomina
' DtaEmpleados.Refresh
 frmEmpleado.ChkSuspendido.Visible = False
'Busco el codigo del empleado para que automaticamente ubique el nombre
 'aunque no existe en la data consulta
 CodEmpleado = -1
 
 frmEmpleado.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo From Empleado WHERE (CodEmpleado1 = '" & frmEmpleado.DBCodigoEmpleado.Text & "') And (Activo = 1)"
frmEmpleado.DtaEmpleado.Refresh

If Not frmEmpleado.DtaEmpleado.Recordset.EOF Then
'Do While Not DtaEmpleado.Recordset.EOF
'     If DtaEmpleado.Recordset("CodEmpleado1") = DBCodigoEmpleado.Text Then
'        If DtaEmpleado.Recordset("activo") = False Then
'           MsgBox "Este empleado ya fue dado de Baja"
'        End If
        
     CodEmpleado = frmEmpleado.DtaEmpleado.Recordset("CodEmpleado")
      frmEmpleado.TxtCodEmpleado.Text = frmEmpleado.DtaEmpleado.Recordset("CodEmpleado")
     
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("numeroruc")) Then
           frmEmpleado.TxtNRuc.Text = frmEmpleado.DtaEmpleado.Recordset("numeroruc")
        End If
        'busco el tipo del archivo
        'Destino = ""
        If Dir(RutaFoto & frmEmpleado.DBCodigoEmpleado.Text & ".jpg") <> "" Then
           Destino = RutaFoto & frmEmpleado.DBCodigoEmpleado.Text & ".jpg"
        ElseIf Dir(RutaFoto & frmEmpleado.DBCodigoEmpleado.Text & ".gif") <> "" Then
           Destino = RutaFoto & frmEmpleado.DBCodigoEmpleado.Text & ".gif"
        ElseIf Dir(RutaFoto & frmEmpleado.DBCodigoEmpleado.Text & ".bmp") <> "" Then
           Destino = RutaFoto & frmEmpleado.DBCodigoEmpleado.Text & ".bmp"
        End If
        
        If (Dir(Destino) <> "") Then
         frmEmpleado.Image1.Picture = LoadPicture(Destino)
        Else
         Destino = RutaFoto + "Zw.bmp"
         frmEmpleado.Image1.Picture = LoadPicture(Destino)
        End If
        
        If frmEmpleado.DtaEmpleado.Recordset("PorcientoIncentivo") = 0 Then
         frmEmpleado.Check1.Value = 0
         frmEmpleado.TxtPorcientoHora.Text = 0
         frmEmpleado.TxtPorcientoHora.Visible = False
        Else
         frmEmpleado.Check1.Value = 1
         frmEmpleado.TxtPorcientoHora.Text = DtaEmpleado.Recordset("PorcientoIncentivo")
         frmEmpleado.TxtPorcientoHora.Visible = True
        End If
        
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("numcedula")) Then
        frmEmpleado.TxtNumCedula.Text = frmEmpleado.DtaEmpleado.Recordset("numcedula")
        End If
        frmEmpleado.ChkSuspendido.Visible = True
        frmEmpleado.TxtNombre1.Text = frmEmpleado.DtaEmpleado.Recordset("Nombre1")
        
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("Nombre2")) Then
        frmEmpleado.TxtNombre2.Text = frmEmpleado.DtaEmpleado.Recordset("Nombre2")
        End If
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("Apellido1")) Then
          frmEmpleado.TxtApellido1.Text = frmEmpleado.DtaEmpleado.Recordset("Apellido1")
        End If
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("Apellido2")) Then
         frmEmpleado.TxtApellido2.Text = frmEmpleado.DtaEmpleado.Recordset("Apellido2")
        End If
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("Direccion")) Then
           frmEmpleado.TxtDireccion.Text = frmEmpleado.DtaEmpleado.Recordset("Direccion")
        End If
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("Nacionalidad")) Then
         frmEmpleado.TxtNacionalidad.Text = frmEmpleado.DtaEmpleado.Recordset("Nacionalidad")
        End If
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("Codgrupo")) Then
            frmEmpleado.TxtCodGrupo = frmEmpleado.DtaEmpleado.Recordset("Codgrupo")
        Else
            frmEmpleado.TxtCodGrupo = ""
            DBCGrupo.Text = ""
        End If
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("CodigoPostal")) Then
          frmEmpleado.TxtCodPostal.Text = frmEmpleado.DtaEmpleado.Recordset("CodigoPostal")
        End If
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("sexo")) Then
          frmEmpleado.CmbSexo.Text = frmEmpleado.DtaEmpleado.Recordset("sexo")
        End If
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("NumeroInss")) Then
        frmEmpleado.TxtNInss.Text = frmEmpleado.DtaEmpleado.Recordset("NumeroInss")
        End If
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("CodDepartamento")) Then
        frmEmpleado.TxtCodDepartamento.Text = frmEmpleado.DtaEmpleado.Recordset("CodDepartamento")
        End If
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("CodCargo")) Then
          frmEmpleado.TxtCodCargo.Text = frmEmpleado.DtaEmpleado.Recordset("CodCargo")
        End If
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("Sindicalista")) Then
          frmEmpleado.CmbSindicalista.Text = frmEmpleado.DtaEmpleado.Recordset("Sindicalista")
        End If
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("numhijos")) Then
          frmEmpleado.TxtNumHijos.Text = frmEmpleado.DtaEmpleado.Recordset("numhijos")
        End If
        frmEmpleado.Caption = "Registro del Empleado: " & frmEmpleado.DBCodigoEmpleado.Text & ": " & frmEmpleado.TxtNombre1.Text & " " & frmEmpleado.TxtNombre2.Text & " " & frmEmpleado.TxtApellido1.Text & " " & frmEmpleado.TxtApellido2.Text
'        frmEmpleado.CmdAcercade.Caption = frmEmpleado.DBCodigoEmpleado.Text & ":   " & frmEmpleado.txtNombre1.Text & " " & frmEmpleado.txtNombre2.Text & " " & frmEmpleado.txtApellido1.Text & " " & frmEmpleado.txtApellido2.Text
'        frmEmpleado.xp_canvas1.Caption = "Registro del Empleado: " & frmEmpleado.DBCodigoEmpleado.Text & ": " & frmEmpleado.txtNombre1.Text & " " & frmEmpleado.txtNombre2.Text & " " & frmEmpleado.txtApellido1.Text & " " & frmEmpleado.txtApellido2.Text
        
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("DiasDescuento")) Then
            frmEmpleado.TxtDiasDescuento.Text = frmEmpleado.DtaEmpleado.Recordset("DiasDescuento")
        Else
            frmEmpleado.TxtDiasDescuento.Text = 0
        End If
        Bandera = False
        
        If frmEmpleado.DtaEmpleado.Recordset("SalarioFijo") = "S" Then
          Salario = True
        Else
          Salario = False
        End If
        
        If frmEmpleado.DtaEmpleado.Recordset("ausente") = True Then
           frmEmpleado.ChkSuspendido.Value = 1
           frmEmpleado.LblSuspendido.Visible = True
        Else
           frmEmpleado.LblSuspendido.Visible = False
           frmEmpleado.ChkSuspendido.Value = 0
        End If
        
        If frmEmpleado.DtaEmpleado.Recordset("salariominimo") = True Then
            frmEmpleado.CmbSalarioMinimo.Text = "Verdaderp"
        Else
           frmEmpleado.CmbSalarioMinimo.Text = "Falso"
        End If
        
        If frmEmpleado.DtaEmpleado.Recordset("ExentoInss") = True Then
            frmEmpleado.CmbExentoInss.Text = "Verdadero"
        Else
           frmEmpleado.CmbExentoInss.Text = "Falso"
        End If
           
        If frmEmpleado.DtaEmpleado.Recordset("ExentoIr") = True Then
            frmEmpleado.CmbExentoIr.Text = "Verdadero"
        Else
           frmEmpleado.CmbExentoIr.Text = "Falso"
        End If
        
        If frmEmpleado.DtaEmpleado.Recordset("PagoInssPatronal") = True Then
            frmEmpleado.CmbPagoInssPatronal.Text = "Verdadero"
        Else
           frmEmpleado.CmbPagoInssPatronal.Text = "Falso"
        End If
        Bandera = True
    
    frmEmpleado.SSTab1.TabEnabled(0) = True
    frmEmpleado.SSTab1.TabEnabled(1) = True
    frmEmpleado.SSTab1.TabEnabled(2) = True
    frmEmpleado.SSTab1.TabEnabled(3) = True
    frmEmpleado.SSTab1.TabEnabled(4) = True
    frmEmpleado.SSTab1.TabEnabled(5) = True
    frmEmpleado.SSTab1.TabEnabled(6) = True

    ' datos de la Nómina

'no olvidar los valores nomina

        frmEmpleado.DtaTipoNomina.Refresh
        Do While Not frmEmpleado.DtaTipoNomina.Recordset.EOF
           If frmEmpleado.DtaTipoNomina.Recordset("CodTipoNomina") = frmEmpleado.DtaEmpleado.Recordset("CodTipoNomina") Then
              frmEmpleado.DBCTipoNomina.Text = frmEmpleado.DtaTipoNomina.Recordset("nomina")
              Exit Do
            End If
        frmEmpleado.DtaTipoNomina.Recordset.MoveNext
        Loop
        
            If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("SueldoPeriodo")) Then
            frmEmpleado.TxtSueldoPeriodo.Text = frmEmpleado.DtaEmpleado.Recordset("SueldoPeriodo")
            End If
            
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("TarifaHoraria")) Then
            frmEmpleado.TxtTarifaHoraria.Text = frmEmpleado.DtaEmpleado.Recordset("TarifaHoraria")
        End If
        
        If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("PorcentajeComision")) Then
            frmEmpleado.TxtComision.Text = frmEmpleado.DtaEmpleado.Recordset("PorcentajeComision")
        End If
      
     
       If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("OtrosIngresos")) Then
          frmEmpleado.TxtOtrosIngresos.Text = frmEmpleado.DtaEmpleado.Recordset("OtrosIngresos")
       End If
       
       If Not IsNull(frmEmpleado.DtaEmpleado.Recordset("DescripOtrIngre")) Then
          frmEmpleado.TxtDescripOtrIngre.Text = frmEmpleado.DtaEmpleado.Recordset("DescripOtrIngre")
       End If
'        Exit Do
        Else 'si no lo encuentra
  
              frmEmpleado.SSTab1.TabEnabled(0) = True
              frmEmpleado.SSTab1.TabEnabled(1) = True
              frmEmpleado.SSTab1.TabEnabled(2) = True
              frmEmpleado.SSTab1.TabEnabled(3) = False
              frmEmpleado.SSTab1.TabEnabled(4) = False
              frmEmpleado.SSTab1.TabEnabled(5) = False
              frmEmpleado.SSTab1.TabEnabled(6) = False
        'frmEmpleado.Caption = "Registro del Empleado: " & frmEmpleado.txtNombre1.Text & " " & frmEmpleado.txtNombre2.Text & " " & frmEmpleado.txtApellido1.Text & " " & frmEmpleado.txtApellido2.Text
     End If
'frmEmpleado.DtaEmpleado.Recordset.MoveNext
'Loop

Evaluar = True
frmEmpleado.DtaHistorico.RecordSource = "SELECT  Codempleado, FechaBaja, MotivoBaja, FechaAumento, MotivoAumento, FechaInicialSusp, FechaFinalSusp, MotivoSuspencion, FechaNacimiento, FechaContrato , CargoInicial, CargoActual, CargoAnterior, SueldoInicial, SueldoAnterior, SueldoActual, CuentaDebito, CuentaCredito From Historico Where (CodEmpleado = " & CodEmpleado & " )"
frmEmpleado.DtaHistorico.Refresh
Do While Not frmEmpleado.DtaHistorico.Recordset.EOF
     If frmEmpleado.DtaHistorico.Recordset("CodEmpleado") = CodEmpleado Then
        If Not IsNull(frmEmpleado.DtaHistorico.Recordset("FechaNacimiento")) Then
            frmEmpleado.MaskEdNacimiento.Value = Format(frmEmpleado.DtaHistorico.Recordset("FechaNacimiento"), "dd/mm/yyyy")
        End If
        
        If Not IsNull(frmEmpleado.DtaHistorico.Recordset("FechaContrato")) Then
            frmEmpleado.MaskEdContrato.Value = Format(frmEmpleado.DtaHistorico.Recordset("FechaContrato"), "dd/mm/yyyy")
        End If
        
        If Not IsNull(frmEmpleado.DtaHistorico.Recordset("CargoInicial")) Then
          frmEmpleado.DBCargoInicial.Text = frmEmpleado.DtaHistorico.Recordset("CargoInicial")
        End If
        If Not IsNull(frmEmpleado.DtaHistorico.Recordset("CargoAnterior")) Then
           frmEmpleado.DBCargoAnterior.Text = frmEmpleado.DtaHistorico.Recordset("CargoAnterior")
        End If
        If Not IsNull(frmEmpleado.DtaHistorico.Recordset("CargoActual")) Then
             frmEmpleado.DBCargoActual.Text = frmEmpleado.DtaHistorico.Recordset("CargoActual")
        End If
        If Not IsNull(frmEmpleado.DtaHistorico.Recordset("MOTIVOBAJA")) Then
              frmEmpleado.TxtMotivoBaja.Text = frmEmpleado.DtaHistorico.Recordset("MOTIVOBAJA")
        End If
        If Not IsNull(frmEmpleado.DtaHistorico.Recordset("MotivoAumento")) Then
             frmEmpleado.TxtMotivoAumento.Text = frmEmpleado.DtaHistorico.Recordset("MotivoAumento")
        End If
        If Not IsNull(frmEmpleado.DtaHistorico.Recordset("MotivoSuspencion")) Then
             frmEmpleado.TxtMotivoSuspencion.Text = frmEmpleado.DtaHistorico.Recordset("MotivoSuspencion")
        End If
        
        frmEmpleado.TxtSueldoInicial.Text = Format((frmEmpleado.DtaHistorico.Recordset("SueldoInicial")), "##,##0.00")
        frmEmpleado.TxtSueldoAnterior.Text = Format((frmEmpleado.DtaHistorico.Recordset("SueldoAnterior")), "##,##0.00")
        frmEmpleado.TxtSueldoActual.Text = Format((frmEmpleado.DtaHistorico.Recordset("SueldoActual")), "##,##0.00")
        
        If Not IsNull(frmEmpleado.DtaHistorico.Recordset("fechabaja")) Then
             MaskEdBaja.Text = frmEmpleado.DtaHistorico.Recordset("fechabaja")
        End If
        
        If Not IsNull(frmEmpleado.DtaHistorico.Recordset("FechaAumento")) Then
             MaskEdAumento.Text = Format(frmEmpleado.DtaHistorico.Recordset("FechaAumento"), "dd/mm/yyyy")
        End If
         
         If Not IsNull(frmEmpleado.DtaHistorico.Recordset("FechaInicialSusp")) Then
            MaskEdSuspencion.Text = frmEmpleado.DtaHistorico.Recordset("FechaInicialSusp")
        End If
        
        If Not IsNull(frmEmpleado.DtaHistorico.Recordset("FechaInicialSusp")) Then
           MaskEdFinalSusp.Text = frmEmpleado.DtaHistorico.Recordset("FechaInicialSusp")
        End If
        
'        If Not IsNull(frmEmpleado.DtaHistorico.Recordset("CuentaDebito")) Then
'         frmEmpleado.TxtDebito.Text = frmEmpleado.DtaHistorico.Recordset("CuentaDebito")
'        End If
        
'        If Not IsNull(frmEmpleado.DtaHistorico.Recordset("cuentacredito")) Then
'         frmEmpleado.TxtCredito.Text = frmEmpleado.DtaHistorico.Recordset("cuentacredito")
'        End If
        Exit Do
   End If
       frmEmpleado.DtaHistorico.Recordset.MoveNext
   Loop

frmEmpleado.DtaInfNomina.Refresh
Do While Not frmEmpleado.DtaInfNomina.Recordset.EOF
     If frmEmpleado.DtaInfNomina.Recordset("CodEmpleado") = CodEmpleado Then
        If Not IsNull(frmEmpleado.DtaInfNomina.Recordset("salariominimo")) Then
             frmEmpleado.CmbSalarioMinimo.Text = frmEmpleado.DtaInfNomina.Recordset("salariominimo")
        End If
        If Not IsNull(frmEmpleado.DtaInfNomina.Recordset("ExentoInss")) Then
             frmEmpleado.CmbExentoInss.Text = frmEmpleado.DtaInfNomina.Recordset("ExentoInss")
        End If
        If Not IsNull(frmEmpleado.DtaInfNomina.Recordset("ExentoIr")) Then
             frmEmpleado.CmbExentoIr.Text = frmEmpleado.DtaInfNomina.Recordset("ExentoIr")
        End If
       
        If Not IsNull(frmEmpleado.DtaInfNomina.Recordset("PagoInssPatronal")) Then
            frmEmpleado.CmbPagoInssPatronal.Text = frmEmpleado.DtaInfNomina.Recordset("PagoInssPatronal")
        End If
       
        frmEmpleado.TxtCodTipoNomina.Text = frmEmpleado.DtaInfNomina.Recordset("CodTipoNomina")
          
        Evaluar = True
        Exit Do
 
     End If
       frmEmpleado.DtaInfNomina.Recordset.MoveNext
   Loop

  frmEmpleado.DtaDepartamento.Refresh
   Do While Not frmEmpleado.DtaDepartamento.Recordset.EOF
     If frmEmpleado.DtaDepartamento.Recordset("CodDepartamento") = frmEmpleado.TxtCodDepartamento.Text Then
        frmEmpleado.DBCDepartamento.Text = frmEmpleado.DtaDepartamento.Recordset("departamento")
        Exit Do
     Else
      'frmEmpleado.DBCDepartamento.Text = ""
     End If
       frmEmpleado.DtaDepartamento.Recordset.MoveNext
   Loop

   frmEmpleado.DtaCargo.Refresh
   Do While Not frmEmpleado.DtaCargo.Recordset.EOF
     If frmEmpleado.DtaCargo.Recordset("CodCargo") = frmEmpleado.TxtCodCargo.Text Then
        frmEmpleado.DBCCargo.Text = frmEmpleado.DtaCargo.Recordset("Cargo")
        Exit Do
     Else
        frmEmpleado.DBCCargo.Text = ""
     End If
       frmEmpleado.DtaCargo.Recordset.MoveNext
   Loop
   
  Evaluar = True
 frmEmpleado.CmbIncapacidad.Text = "No"
  frmEmpleado.DtaIncapacidades.Refresh
   Do While Not frmEmpleado.DtaIncapacidades.Recordset.EOF
     If frmEmpleado.DtaIncapacidades.Recordset("CodEmpleado") = CodEmpleado Then
        frmEmpleado.CmbIncapacidad.Text = "Si"
        Exit Do
     Else
        frmEmpleado.CmbIncapacidad.Text = "No"
     End If
       frmEmpleado.DtaIncapacidades.Recordset.MoveNext
   Loop
Evaluar = True
   Salida = False

   



'realizo los sql's de los incentivos y de las deducciones

CodEmpleado1 = frmEmpleado.DBCodigoEmpleado.Text

SQlIncentivos = "SELECT Incentivo.NumIncentivo, TipoIncentivo.Incentivo, Incentivo.CodEmpleado, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado FROM TipoIncentivo INNER JOIN (Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo) ON TipoIncentivo.CodTipoIncentivo = Incentivo.CodTipoIncentivo WHERE (Incentivo.CodEmpleado= " & CodEmpleado & ") And (DetalleIncentivo.Pagado = " & 0 & ")"
frmEmpleado.DtaDetalleIncentivo.RecordSource = SQlIncentivos
frmEmpleado.DtaDetalleIncentivo.Refresh

frmEmpleado.DbGIncentivos.Columns(0).Visible = False
frmEmpleado.DbGIncentivos.Columns(2).Visible = False
frmEmpleado.DbGIncentivos.Columns(5).Visible = False

SQlDeducciones = "SELECT Deduccion.NumDeduccion, TipoDeduccion.Deduccion, Deduccion.CodEmpleado, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado FROM TipoDeduccion INNER JOIN (Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion) ON (TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion) WHERE Deduccion.CodEmpleado=" & CodEmpleado & " AND DetalleDeduccion.Pagado= " & 0 & " "
frmEmpleado.DtaDetalleDeduccion.RecordSource = SQlDeducciones
frmEmpleado.DtaDetalleDeduccion.Refresh

frmEmpleado.DbgDeducciones.Columns(0).Visible = False
frmEmpleado.DbgDeducciones.Columns(2).Visible = False
frmEmpleado.DbgDeducciones.Columns(5).Visible = False
SQlPrestamo = "SELECT NumPrestamo, CuentaDebito, CuentaCredito, Monto, CantCuotas, Interes, Saldo, FechaInicial, Cancelado, Moneda, CuotasIguales, CodEmpleado From Prestamo WHERE Prestamo.CodEmpleado=" & CodEmpleado & " AND Prestamo.Cancelado=0"
frmEmpleado.DtaPrestamo.RecordSource = SQlPrestamo
frmEmpleado.DtaPrestamo.Refresh
If Not frmEmpleado.DtaPrestamo.Recordset.EOF Then
numeroPrestamo = frmEmpleado.DtaPrestamo.Recordset("NumPrestamo")
Else
 numeroPrestamo = -100
End If
SqlDetallePrestamo = "SELECT MovPrestamo.ID,MovPrestamo.NumPrestamo, MovPrestamo.NumCuota, MovPrestamo.Monto, MovPrestamo.Interes, MovPrestamo.CuotaIgual,MovPrestamo.SaldoCuota , MovPrestamo.Cancelado FROM Prestamo INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo Where (MovPrestamo.Cancelado = 0) And (MovPrestamo.NumPrestamo = " & numeroPrestamo & ")"
frmEmpleado.DtaMovPrestamo.RecordSource = SqlDetallePrestamo
frmEmpleado.DtaMovPrestamo.Refresh






If Not frmEmpleado.DtaPrestamo.Recordset.EOF Then
   frmEmpleado.TxtCreditoPrestamo.Text = frmEmpleado.DtaPrestamo.Recordset("cuentacredito")
   frmEmpleado.TxtDebitoPrestamo.Text = frmEmpleado.DtaPrestamo.Recordset("CuentaDebito")
   frmEmpleado.DbgrLibreta.Columns(0).Visible = False
frmEmpleado.DbgrLibreta.Columns(1).Visible = False
frmEmpleado.DbgrLibreta.Columns(7).Visible = False

Else
   frmEmpleado.TxtCreditoPrestamo.Text = " "
   frmEmpleado.TxtDebitoPrestamo.Text = " "
End If




SqlDetalleSubsidio = "SELECT Subsidio.NumSubsidio, Subsidio.CodEmpleado, Subsidio.CodTipoSubsidio, TipoSubsidio.Subsidio,DetalleSubsidio.Descripcion, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.Pagado FROM TipoSubsidio INNER JOIN (Subsidio INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio) ON TipoSubsidio.CodTipoSubsidio = Subsidio.CodTipoSubsidio WHERE DetalleSubsidio.Pagado=0 And Subsidio.CodEmpleado=" & CodEmpleado & " "
frmEmpleado.DtaDetalleSubsidio.RecordSource = SqlDetalleSubsidio
frmEmpleado.DtaDetalleSubsidio.Refresh

frmEmpleado.DbgrSubsidios.Columns(0).Visible = False
frmEmpleado.DbgrSubsidios.Columns(1).Visible = False
frmEmpleado.DbgrSubsidios.Columns(2).Visible = False
frmEmpleado.DbgrSubsidios.Columns(7).Visible = False
frmEmpleado.DbgrSubsidios.Columns(5).Width = 1200
frmEmpleado.DbgrSubsidios.Columns(6).Width = 500



If frmEmpleado.TxtNombre1.Text = "" Then
    frmEmpleado.SSTab1.TabEnabled(1) = True
    frmEmpleado.SSTab1.TabEnabled(2) = True
    frmEmpleado.SSTab1.TabEnabled(3) = True
    frmEmpleado.SSTab1.TabEnabled(4) = True
    frmEmpleado.SSTab1.TabEnabled(5) = True
    frmEmpleado.SSTab1.TabEnabled(6) = True
End If

      If Salario = True Then
         frmEmpleado.ChkSalarioFijo.Value = 1
         frmEmpleado.TxtComision.Enabled = False
        Else
         frmEmpleado.ChkSalarioFijo.Value = 0
         frmEmpleado.TxtComision.Enabled = True
        End If


frmEmpleado.DtaEmpleado.RecordSource = "SELECT CodEmpleado,CodEmpleado1,Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo From Empleado WHERE     (CodEmpleado1 = '" & frmEmpleado.DBCodigoEmpleado.Text & "') "
frmEmpleado.DtaEmpleado.Refresh

If Not frmEmpleado.DtaEmpleado.Recordset.EOF Then
 frmEmpleado.DBCodigoEmpleado.Text = frmEmpleado.DtaEmpleado.Recordset("CodEmpleado1")
End If

'frmEmpleado.DBCodigoEmpleado.Columns(0).Visible = False
'frmEmpleado.DBCodigoEmpleado.Columns(1).Caption = "Codigo"
'frmEmpleado.DBCodigoEmpleado.Columns(1).Width = 800
'frmEmpleado.DBCodigoEmpleado.Columns(2).Visible = False


frmEmpleado.MousePointer = 0
frmEmpleado.AutoRedraw = True




End Sub




