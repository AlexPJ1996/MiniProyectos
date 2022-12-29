Imports System.Data.SqlClient
Public Class ConDB
    Dim Cadena As String = "Data Source=(LocalDB)\MSSQLLocalDB; AttachDbFilename=|DataDirectory|\DB.mdf; Integrated Security=True"
    Dim Conectar As New SqlConnection(Cadena)

    Sub Consulta(ByVal Tabla As DataGridView, ByVal SQL As String)
        Try
            Dim DA As New SqlDataAdapter(SQL, Conectar)
            Dim DT As New DataTable
            DA.Fill(DT)
            Tabla.DataSource = DT
        Catch ex As Exception
            MessageBox.Show("Error: " + ex.Message)
        End Try
    End Sub

    Sub Operaciones(ByVal Tabla As DataGridView, ByVal SQL As String)
        Conectar.Open()
        Try
            Dim CMD As New SqlCommand(SQL, Conectar)
            CMD.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Error: " + ex.Message)
        End Try
        Conectar.Close()
    End Sub
End Class
