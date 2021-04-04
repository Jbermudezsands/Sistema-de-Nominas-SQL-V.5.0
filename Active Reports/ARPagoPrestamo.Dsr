VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARPagoPrestamo 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "ARPagoPrestamo.dsx":0000
End
Attribute VB_Name = "ARPagoPrestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Dim SqlPagoPrestamo As String
'DaoDtaPagoPrestamo.DatabaseName = ruta
'DaoDtaPagoPrestamo.ConnectionString = Conexion

'SqlPagoPrestamo = "SELECT MovPrestamo.NumPrestamo, Prestamo.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Prestamo.Monto AS Prestamo, Prestamo.Saldo, Prestamo.CuotasIguales, MovPrestamo.NumCuota, MovPrestamo.Monto, MovPrestamo.Interes, MovPrestamo.CuotaIgual, MovPrestamo.SaldoCuota, MovPrestamo.Cancelado, MovPrestamo.NumNomina FROM (Empleado INNER JOIN Prestamo ON Empleado.CodEmpleado = Prestamo.CodEmpleado) INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo WHERE MovPrestamo.Cancelado=True AND MovPrestamo.NumNomina= " & NumNomina & ""
'SqlPagoPrestamo = "SELECT MovPrestamo.NumPrestamo, Prestamo.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Prestamo.Monto AS Prestamo, Prestamo.Saldo, Prestamo.CuotasIguales, MovPrestamo.NumCuota, MovPrestamo.Monto, MovPrestamo.Interes, MovPrestamo.CuotaIgual, MovPrestamo.SaldoCuota, MovPrestamo.Cancelado, MovPrestamo.NumNomina FROM (Empleado INNER JOIN Prestamo ON Empleado.CodEmpleado = Prestamo.CodEmpleado) INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo WHERE MovPrestamo.Cancelado=True AND MovPrestamo.NumNomina= " & NumNomina & ""
'DaoDtaPagoPrestamo.RecordSource = SqlPagoPrestamo
'DaoDtaPagoPrestamo.Refresh
LblTitulo.Caption = Titulo
LblSubtitulo.Caption = SubTitulo
ImgLogo.Picture = LoadPicture(RutaLogo)
End Sub

