Imports System.Data.OleDb
Public Class ConDB
	Dim Cadena As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=|DataDirectory|\DataBase.accdb; Jet OLEDB:Database Password=Test"
	Dim Conectar As New OleDbConnection(Cadena)
	Private Con As New OleDbConnection(Cadena)
    'Probar Conexión con Base de datos
    Sub Conexion()
        Try
            Conectar.Open()
            MessageBox.Show("Conectado")
            Conectar.Close()
        Catch ex As Exception
            MessageBox.Show("Error: " + ex.Message)
        End Try
    End Sub
    'procedimiento mostrar datos en un DataGridView mediante consultas SELECT
    Sub Consulta(ByVal Tabla As DataGridView, ByVal SQL As String)
        Try
            Dim DA As New OleDbDataAdapter(SQL, Conectar)
            Dim DT As New DataTable
            DA.Fill(DT)
            Tabla.DataSource = DT
        Catch ex As Exception
            MessageBox.Show("Error: " + ex.Message)
        End Try
    End Sub

    'Procedimiento para rellenar ComboBox con datos de una tabla
    Function Rellenar(ByVal SQL As String)
        Dim DT As New DataTable
        Try
            Dim DA As New OleDbDataAdapter(SQL, Con)
            DA.Fill(DT)
        Catch ex As Exception
            MessageBox.Show("Error: " + ex.Message)
        End Try
        Return DT
    End Function
End Class
