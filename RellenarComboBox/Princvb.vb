Public Class Princvb
    Dim OBJ As New ConDB
    Dim CanPaises As Integer
    Private Sub Princvb_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = My.Application.Info.Title
        PropCB()
    End Sub

    Sub PropCB()
        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
        ComboBox1.FlatStyle = FlatStyle.Popup
        ComboBox2.DropDownStyle = ComboBoxStyle.DropDownList
        ComboBox2.FlatStyle = FlatStyle.Popup
    End Sub

    Private Sub TCon_Click(sender As Object, e As EventArgs) Handles TCon.Click
        OBJ.Conexion()
    End Sub

    Private Sub Rellenar_Click(sender As Object, e As EventArgs) Handles Rellenar.Click
        CanRegistros()
        RellenarCB1()
        RellenarCB2()
    End Sub

    Sub CanRegistros()
        Dim SQL1 As String = "SELECT COUNT(*) FROM Paises"
        OBJ.Consulta(DataGridView1, SQL1)
        If DataGridView1.RowCount > 0 Then
            CanPaises = DataGridView1.CurrentRow.Cells(0).Value
        End If
    End Sub

    Sub RellenarCB1()
        Try
            'Crear variable que guardará la consulta SQL
            Dim SQL As String = "SELECT Pais FROM Paises ORDER BY Pais ASC"
            'Crear DataTable que recibirá los datos de la consulta
            Dim DT As New DataTable
            'Asignamos a nuestro DataTable el objeto y llamamos nuestra función
            'A la que le pasamos nuestra consulta SQL
            DT = OBJ.Rellenar(SQL)
            'Asignamos como origen de datos del ComboBox en su propiedad (DataSource)
            'Nuestro DataTable que ya tiene los valores devuelto por nuestro procedimiento
            ComboBox1.DataSource = DT
            'Le indicamos a nuestro ComboBox que columna debe mostrar
            ComboBox1.DisplayMember = "Pais"
            'Le indicamos que simpre muestre el primer valor
            ComboBox1.SelectedIndex = 0
        Catch ex As Exception

        End Try
    End Sub

    Sub RellenarCB2()
        Try
            If Val(CanPaises) > 0 Then
                Dim SQL As String = "SELECT Estado FROM Estados WHERE (Pais = '" + ComboBox1.Text + "')  ORDER BY Estado ASC"
                Dim DT As New DataTable
                DT = OBJ.Rellenar(SQL)
                ComboBox2.DataSource = DT
                ComboBox2.DisplayMember = "Estado"
                ComboBox2.SelectedIndex = 0
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text <> "" Then
            RellenarCB2()
        End If
    End Sub
End Class