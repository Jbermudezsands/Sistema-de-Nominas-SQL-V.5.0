VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arepInasistenciaGeneral 
   Caption         =   "Inasistencia"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "arepInasistenciaGeneral.dsx":0000
End
Attribute VB_Name = "arepInasistenciaGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()

Me.txtNombreCompleto.Text = Me.txtNombre1.Text & " " & Me.txtNombre2.Text & " " & Me.txtApellido1.Text & " " & Me.txtApellido2.Text

End Sub

Private Sub GroupFooter1_Format()


frmRepAsistencia.adoAsistencia.CommandType = adCmdText
frmRepAsistencia.adoAsistencia.RecordSource = "SELECT dbo.Empleado.CodEmpleado, dbo.Empleado.CodEmpleado1, dbo.Departamento.Departamento, dbo.Empleado.Activo " & _
                                              "FROM dbo.Empleado INNER JOIN dbo.Departamento ON dbo.Empleado.CodDepartamento = dbo.Departamento.CodDepartamento Where (dbo.Empleado.Activo = 1) " & _
                                              "AND dbo.Departamento.Departamento ='" & Me.txtDepto.Text & "' ORDER BY dbo.Departamento.Departamento"
                                              
frmRepAsistencia.adoAsistencia.Refresh
 
 
Me.lblTotalDepto.Caption = frmRepAsistencia.adoAsistencia.Recordset.RecordCount
 
 
 
 
 

End Sub

Private Sub PageHeader_Format()

Me.lblHora.Caption = Time


End Sub
