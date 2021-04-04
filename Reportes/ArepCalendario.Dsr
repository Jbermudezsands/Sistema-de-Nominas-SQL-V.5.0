VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepCalendario 
   Caption         =   "Reporte de Calendario"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepCalendario.dsx":0000
End
Attribute VB_Name = "ArepCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
     Me.LblTitulo3.Caption = "Reporte Calendario de Nominas"
     Me.LblTitulo.Caption = Titulo
     Me.LblSubtitulo.Caption = SubTitulo
     If Dir(RutaLogo) <> "" Then
       Me.ImgLogo.Picture = LoadPicture(RutaLogo)
     End If
     
     Me.LblTipo.Caption = "Nomina: " & frmFecha.DbTipoNomina.Text & "    Desde: " & frmFecha.mskInicio.Value & " Hasta: " & frmFecha.mskFinal.Value
     
End Sub

Private Sub Detail_Format()
 Dim sql As String, CodTipoNomina As String, Mes As String, Año As String
        frmFecha.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, UltFecha, TipoPago, Moneda, MantValor From TipoNomina WHERE  (Nomina = '" & frmFecha.DbTipoNomina.Text & "')"
        frmFecha.DtaConsulta.Refresh
        If frmFecha.DtaConsulta.Recordset.EOF Then
          Exit Sub
        Else
         TipoNomina = frmFecha.DtaConsulta.Recordset("Periodo")
         CodTipoNomina = frmFecha.DtaConsulta.Recordset("CodTipoNomina")
        End If
      Año = frmFecha.TxtAño.Text
       
       '----------------------------MES DE ENERO---------------------------------------------------
       Mes = "01"
       sql = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla " & _
            "WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (mes = '" & Mes & "') AND (año = " & Año & ") ORDER BY Periodo"
       Set Me.SubReportEnero.object = New ArepSub
       Me.SubReportEnero.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReportEnero.object.DataControl1.Source = sql

       '----------------------------MES DE FEBRERO---------------------------------------------------
       Mes = "02"
       sql = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla " & _
            "WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (mes = '" & Mes & "') AND (año = " & Año & ") ORDER BY Periodo"
       Set Me.SubReportFebrero.object = New ArepSub
       Me.SubReportFebrero.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReportFebrero.object.DataControl1.Source = sql
       
              '----------------------------MES DE MARZO---------------------------------------------------
       Mes = "03"
       sql = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla " & _
            "WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (mes = '" & Mes & "') AND (año = " & Año & ") ORDER BY Periodo"
       Set Me.SubReportMarzo.object = New ArepSub
       Me.SubReportMarzo.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReportMarzo.object.DataControl1.Source = sql
       
              '----------------------------MES DE ABRIL---------------------------------------------------
       Mes = "04"
       sql = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla " & _
            "WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (mes = '" & Mes & "') AND (año = " & Año & ") ORDER BY Periodo"
       Set Me.SubReportAbril.object = New ArepSub
       Me.SubReportAbril.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReportAbril.object.DataControl1.Source = sql
       
              '----------------------------MES DE MAYO---------------------------------------------------
       Mes = "05"
       sql = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla " & _
            "WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (mes = '" & Mes & "') AND (año = " & Año & ") ORDER BY Periodo"
       Set Me.SubReportMayo.object = New ArepSub
       Me.SubReportMayo.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReportMayo.object.DataControl1.Source = sql
       
              '----------------------------MES DE JUNIO---------------------------------------------------
       Mes = "06"
       sql = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla " & _
            "WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (mes = '" & Mes & "') AND (año = " & Año & ") ORDER BY Periodo"
       Set Me.SubReportJunio.object = New ArepSub
       Me.SubReportJunio.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReportJunio.object.DataControl1.Source = sql
       
              '----------------------------MES DE JULIO---------------------------------------------------
       Mes = "07"
       sql = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla " & _
            "WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (mes = '" & Mes & "') AND (año = " & Año & ") ORDER BY Periodo"
       Set Me.SubReportJulio.object = New ArepSub
       Me.SubReportJulio.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReportJulio.object.DataControl1.Source = sql
       
              '----------------------------MES DE AGOSTO---------------------------------------------------
       Mes = "08"
       sql = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla " & _
            "WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (mes = '" & Mes & "') AND (año = " & Año & ") ORDER BY Periodo"
       Set Me.SubReportAgosto.object = New ArepSub
       Me.SubReportAgosto.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReportAgosto.object.DataControl1.Source = sql
       
              '----------------------------MES DE SEPTIEMBRE---------------------------------------------------
       Mes = "09"
       sql = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla " & _
            "WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (mes = '" & Mes & "') AND (año = " & Año & ") ORDER BY Periodo"
       Set Me.SubReportSeptiembre.object = New ArepSub
       Me.SubReportSeptiembre.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReportSeptiembre.object.DataControl1.Source = sql
       
                     '----------------------------MES DE OCTUBRE---------------------------------------------------
       Mes = "10"
       sql = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla " & _
            "WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (mes = '" & Mes & "') AND (año = " & Año & ") ORDER BY Periodo"
       Set Me.SubReportOctubre.object = New ArepSub
       Me.SubReportOctubre.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReportOctubre.object.DataControl1.Source = sql
       
                     '----------------------------MES DE NOVIEMBRE---------------------------------------------------
       Mes = "11"
       sql = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla " & _
            "WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (mes = '" & Mes & "') AND (año = " & Año & ") ORDER BY Periodo"
       Set Me.SubReportNoviembre.object = New ArepSub
       Me.SubReportNoviembre.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReportNoviembre.object.DataControl1.Source = sql
       
                     '----------------------------MES DE DICIEMBRE---------------------------------------------------
       Mes = "12"
       sql = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla " & _
            "WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (mes = '" & Mes & "') AND (año = " & Año & ") ORDER BY Periodo"
       Set Me.SubReportDiciembre.object = New ArepSub
       Me.SubReportDiciembre.object.DataControl1.ConnectionString = ConexionReporte
       Me.SubReportDiciembre.object.DataControl1.Source = sql
End Sub

