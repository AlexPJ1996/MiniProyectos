Public Class Princ
    Dim OBJ As New ConDB
    Dim SumaE As Double
    Dim SumaS As Double
    Dim ST As Double

    Private Sub Princ_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Tiempo.Enabled = True
        Tiempo_Tick(sender, e)
        FText()
        PSel()
    End Sub

    Sub FText()
        Me.Text = My.Application.Info.Title
    End Sub

    Sub PSel()
        Dim SQLE As String = "SELECT ID, Cantidad, Fecha, Hora FROM Entrada ORDER BY ID DESC"
        OBJ.Consulta(RegsE, SQLE)
        Dim SQLS As String = "SELECT ID, Cantidad, Fecha, Hora FROM Salida ORDER BY ID DESC"
        OBJ.Consulta(RegsS, SQLS)
        PSum()
        CellE()
        DGVF()
    End Sub

    Sub PSum()
        If RegsE.RowCount > 0 Then
            Dim SQLE As String = "SELECT Sum(Cantidad) AS 'Total' FROM Entrada"
            OBJ.Consulta(ESuma, SQLE)
            SumaE = ESuma.CurrentRow.Cells(0).Value
        Else
            SumaE = 0
        End If

        If RegsS.RowCount > 0 Then
            Dim SQLS As String = "SELECT SUM(Cantidad) AS 'Total' FROM Salida"
            OBJ.Consulta(SSuma, SQLS)
            SumaS = SSuma.CurrentRow.Cells(0).Value
        Else
            SumaS = 0
        End If
        ST = Val(SumaE) - Val(SumaS)
        LST.Text = ST
        LST.Text = FormatCurrency((LST.Text), "$ 0")
    End Sub

    Sub CellE()
        If RegsE.RowCount > 0 Then
            EID.Text = RegsE.CurrentRow.Cells(0).Value
        End If

        If RegsS.RowCount > 0 Then
            SID.Text = RegsS.CurrentRow.Cells(0).Value
        End If
    End Sub

    Sub DGVF()
        RegsE.Columns(0).Width = 65
        RegsE.Columns(1).Width = 125
        RegsE.Columns(2).Width = 100
        RegsE.Columns(3).Width = 110

        RegsS.Columns(0).Width = 65
        RegsS.Columns(1).Width = 125
        RegsS.Columns(2).Width = 100
        RegsS.Columns(3).Width = 110
    End Sub

    Private Sub CanE_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CanE.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub HFecE_CheckedChanged(sender As Object, e As EventArgs) Handles HFecE.CheckedChanged
        If HFecE.Checked = True Then
            FecE.Enabled = True
        Else
            FecE.Enabled = False
            FecE.Text = Today
        End If
    End Sub

    Private Sub HHorE_CheckedChanged(sender As Object, e As EventArgs) Handles HHorE.CheckedChanged
        If HHorE.Checked = True Then
            HorE.Enabled = True
        Else
            HorE.Enabled = False
            HorE.Text = TimeOfDay
        End If
    End Sub

    Private Sub LimE_Click(sender As Object, e As EventArgs) Handles LimE.Click
        CanE.Clear()
        FecE.Text = Today
        HorE.Text = TimeOfDay
        CanE.Focus()
    End Sub

    Private Sub AgE_Click(sender As Object, e As EventArgs) Handles AgE.Click
        If CanE.Text = "" Then
            MessageBox.Show("Ingrese un valor")
            CanE.Focus()
        Else
            PInsE()
            PSel()
            LimE_Click(sender, e)
        End If
    End Sub

    Sub PInsE()
        Dim SQL As String = "INSERT INTO Entrada (Cantidad, Fecha, Hora) SELECT " & Val(CanE.Text) & ", '" & FecE.Text & "', '" & HorE.Text & "'"
        OBJ.Operaciones(RegsE, SQL)
    End Sub

    Private Sub ReE_Click(sender As Object, e As EventArgs) Handles ReE.Click
        If RegsE.RowCount < 1 Then
            MessageBox.Show("Error: no hay datos con que trabajar")
            CanE.Focus()
        Else
            If MessageBox.Show("¿Desea remover esta entrada?", "", MessageBoxButtons.YesNo) = vbYes Then
                PDelE()
                PSel()
                LimE_Click(sender, e)
            Else
                CanE.Focus()
            End If
        End If
    End Sub

    Sub PDelE()
        Dim SQL As String = "DELETE FROM Entrada WHERE (ID = " & EID.Text & ")"
        OBJ.Operaciones(RegsE, SQL)
    End Sub

    Private Sub RegsE_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles RegsE.CellEnter
        CellE()
    End Sub

    Private Sub CanS_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CanS.KeyPress
        e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub

    Private Sub HFecS_CheckedChanged(sender As Object, e As EventArgs) Handles HFecS.CheckedChanged
        If HFecS.Checked = True Then
            FecS.Enabled = True
        Else
            FecS.Enabled = False
            FecS.Text = Today
        End If
    End Sub

    Private Sub HHorS_CheckedChanged(sender As Object, e As EventArgs) Handles HHorS.CheckedChanged
        If HHorS.Checked = True Then
            HorS.Enabled = True
        Else
            HorS.Enabled = False
            HorS.Text = TimeOfDay
        End If
    End Sub

    Private Sub LimS_Click(sender As Object, e As EventArgs) Handles LimS.Click
        CanS.Clear()
        FecS.Text = Today
        HorS.Text = TimeOfDay
        CanS.Focus()
    End Sub

    Private Sub AgS_Click(sender As Object, e As EventArgs) Handles AgS.Click
        If CanS.Text = "" Then
            MessageBox.Show("Ingrese un valor")
            CanS.Focus()
        Else
            PInsS()
            PSel()
            LimS_Click(sender, e)
        End If
    End Sub

    Sub PInsS()
        Dim SQL As String = "INSERT INTO Salida (Cantidad, Fecha, Hora) SELECT " & Val(CanS.Text) & ", '" & FecS.Text & "', '" & HorS.Text & "'"
        OBJ.Operaciones(RegsS, SQL)
    End Sub

    Private Sub ReS_Click(sender As Object, e As EventArgs) Handles ReS.Click
        If RegsS.RowCount < 1 Then
            MessageBox.Show("Error: no hay datos con que trabajar")
            CanS.Focus()
        Else
            If MessageBox.Show("¿Desea remover esta entrada?", "", MessageBoxButtons.YesNo) = vbYes Then
                PDelS()
                PSel()
                LimS_Click(sender, e)
            Else
                CanS.Focus()
            End If
        End If
    End Sub

    Sub PDelS()
        Dim SQL As String = "DELETE FROM Salida WHERE (ID = " & SID.Text & ")"
        OBJ.Operaciones(RegsS, SQL)
    End Sub

    Private Sub RegsS_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles RegsS.CellEnter
        CellE()
    End Sub

    Private Sub Tiempo_Tick(sender As Object, e As EventArgs) Handles Tiempo.Tick
        Tempo()
    End Sub

    Sub Tempo()
        If HFecE.Checked = False Then
            FecE.Text = Today
        End If

        If HHorE.Checked = False Then
            HorE.Text = TimeOfDay
        End If

        If HFecS.Checked = False Then
            FecS.Text = Today
        End If

        If HHorS.Checked = False Then
            HorS.Text = TimeOfDay
        End If
    End Sub

    Private Sub Vaciar_Click(sender As Object, e As EventArgs) Handles Vaciar.Click
        If MessageBox.Show("¿Está seguro que desea eliminar todos los registros del sistema?", "Atención: Acción irreversible", MessageBoxButtons.YesNo) = vbYes Then
            PDelT()
            PSel()
            MessageBox.Show("Todos los datos han sido eliminados")
            LimE_Click(sender, e)
            LimS_Click(sender, e)
            CanE.Focus()
        End If
    End Sub

    Sub PDelT()
        Dim SQLE As String = "TRUNCATE TABLE Entrada"
        OBJ.Operaciones(RegsS, SQLE)

        Dim SQLS As String = "TRUNCATE TABLE Salida"
        OBJ.Operaciones(RegsS, SQLS)
    End Sub
End Class