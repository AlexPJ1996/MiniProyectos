Public Class Ejemplos
    
	'Procedimiento para solo recibir numeros en un TextBox, mediante el evento KeyPress
	Private Sub MyKeyPress(ByRef e As System.Windows.Forms.KeyPressEventArgs)
		e.Handled = Not IsNumeric(e.KeyChar) And Not Char.IsControl(e.KeyChar)
    End Sub
	
	'Procedimiento para mostrar datos de un DataGridView en Label/TextBox
	Private Sub CellE()
        [Label/TextBox] = [DataGridView].CurrentRow.Cells([Columna N°]).Value
    End Sub
	
	'Procedimiento para mostrar texto de un Label/TextBox con formato moneda
	Sub Moneda()
	    [Label/TextBox] = FormatCurrency(([Label/TextBox]), "$ 0")
		[Label/TextBox] = Format([Label/TextBox], "$ 0")
	End Sub
End Class