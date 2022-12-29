<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Princ
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Princ))
        Me.GBE = New System.Windows.Forms.GroupBox()
        Me.LimE = New System.Windows.Forms.Button()
        Me.HHorE = New System.Windows.Forms.CheckBox()
        Me.HFecE = New System.Windows.Forms.CheckBox()
        Me.ReE = New System.Windows.Forms.Button()
        Me.AgE = New System.Windows.Forms.Button()
        Me.HorE = New System.Windows.Forms.DateTimePicker()
        Me.FecE = New System.Windows.Forms.DateTimePicker()
        Me.CanE = New System.Windows.Forms.TextBox()
        Me.LCanE = New System.Windows.Forms.Label()
        Me.Tiempo = New System.Windows.Forms.Timer(Me.components)
        Me.RegsE = New System.Windows.Forms.DataGridView()
        Me.RegsS = New System.Windows.Forms.DataGridView()
        Me.GBS = New System.Windows.Forms.GroupBox()
        Me.LimS = New System.Windows.Forms.Button()
        Me.HHorS = New System.Windows.Forms.CheckBox()
        Me.HFecS = New System.Windows.Forms.CheckBox()
        Me.ReS = New System.Windows.Forms.Button()
        Me.AgS = New System.Windows.Forms.Button()
        Me.HorS = New System.Windows.Forms.DateTimePicker()
        Me.FecS = New System.Windows.Forms.DateTimePicker()
        Me.CanS = New System.Windows.Forms.TextBox()
        Me.LCanS = New System.Windows.Forms.Label()
        Me.EID = New System.Windows.Forms.TextBox()
        Me.ESuma = New System.Windows.Forms.DataGridView()
        Me.SID = New System.Windows.Forms.TextBox()
        Me.SSuma = New System.Windows.Forms.DataGridView()
        Me.Suma = New System.Windows.Forms.ToolStrip()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.LCanA = New System.Windows.Forms.ToolStripLabel()
        Me.LST = New System.Windows.Forms.ToolStripLabel()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.Vaciar = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.GBE.SuspendLayout()
        CType(Me.RegsE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RegsS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GBS.SuspendLayout()
        CType(Me.ESuma, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SSuma, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Suma.SuspendLayout()
        Me.SuspendLayout()
        '
        'GBE
        '
        Me.GBE.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.GBE.Controls.Add(Me.LimE)
        Me.GBE.Controls.Add(Me.HHorE)
        Me.GBE.Controls.Add(Me.HFecE)
        Me.GBE.Controls.Add(Me.ReE)
        Me.GBE.Controls.Add(Me.AgE)
        Me.GBE.Controls.Add(Me.HorE)
        Me.GBE.Controls.Add(Me.FecE)
        Me.GBE.Controls.Add(Me.CanE)
        Me.GBE.Controls.Add(Me.LCanE)
        Me.GBE.Location = New System.Drawing.Point(14, 13)
        Me.GBE.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GBE.Name = "GBE"
        Me.GBE.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GBE.Size = New System.Drawing.Size(333, 194)
        Me.GBE.TabIndex = 0
        Me.GBE.TabStop = False
        Me.GBE.Text = "Entrada de dinero"
        '
        'LimE
        '
        Me.LimE.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.LimE.Location = New System.Drawing.Point(242, 58)
        Me.LimE.Name = "LimE"
        Me.LimE.Size = New System.Drawing.Size(84, 31)
        Me.LimE.TabIndex = 6
        Me.LimE.Text = "Limpiar"
        Me.LimE.UseVisualStyleBackColor = True
        '
        'HHorE
        '
        Me.HHorE.AutoSize = True
        Me.HHorE.Location = New System.Drawing.Point(10, 95)
        Me.HHorE.Name = "HHorE"
        Me.HHorE.Size = New System.Drawing.Size(63, 23)
        Me.HHorE.TabIndex = 4
        Me.HHorE.Text = "Hora"
        Me.HHorE.UseVisualStyleBackColor = True
        '
        'HFecE
        '
        Me.HFecE.AutoSize = True
        Me.HFecE.Location = New System.Drawing.Point(10, 63)
        Me.HFecE.Name = "HFecE"
        Me.HFecE.Size = New System.Drawing.Size(68, 23)
        Me.HFecE.TabIndex = 2
        Me.HFecE.Text = "Fecha"
        Me.HFecE.UseVisualStyleBackColor = True
        '
        'ReE
        '
        Me.ReE.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ReE.Location = New System.Drawing.Point(150, 158)
        Me.ReE.Name = "ReE"
        Me.ReE.Size = New System.Drawing.Size(86, 31)
        Me.ReE.TabIndex = 8
        Me.ReE.Text = "Remover"
        Me.ReE.UseVisualStyleBackColor = True
        '
        'AgE
        '
        Me.AgE.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.AgE.Location = New System.Drawing.Point(8, 158)
        Me.AgE.Name = "AgE"
        Me.AgE.Size = New System.Drawing.Size(86, 31)
        Me.AgE.TabIndex = 7
        Me.AgE.Text = "Agregar"
        Me.AgE.UseVisualStyleBackColor = True
        '
        'HorE
        '
        Me.HorE.CustomFormat = ""
        Me.HorE.Enabled = False
        Me.HorE.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.HorE.Location = New System.Drawing.Point(86, 91)
        Me.HorE.Name = "HorE"
        Me.HorE.ShowUpDown = True
        Me.HorE.Size = New System.Drawing.Size(150, 26)
        Me.HorE.TabIndex = 5
        '
        'FecE
        '
        Me.FecE.CustomFormat = ""
        Me.FecE.Enabled = False
        Me.FecE.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.FecE.Location = New System.Drawing.Point(86, 58)
        Me.FecE.Name = "FecE"
        Me.FecE.Size = New System.Drawing.Size(150, 26)
        Me.FecE.TabIndex = 3
        '
        'CanE
        '
        Me.CanE.Location = New System.Drawing.Point(86, 26)
        Me.CanE.MaxLength = 11
        Me.CanE.Name = "CanE"
        Me.CanE.Size = New System.Drawing.Size(150, 26)
        Me.CanE.TabIndex = 1
        '
        'LCanE
        '
        Me.LCanE.AutoSize = True
        Me.LCanE.Location = New System.Drawing.Point(8, 29)
        Me.LCanE.Name = "LCanE"
        Me.LCanE.Size = New System.Drawing.Size(70, 19)
        Me.LCanE.TabIndex = 0
        Me.LCanE.Text = "Cantidad"
        '
        'Tiempo
        '
        Me.Tiempo.Enabled = True
        '
        'RegsE
        '
        Me.RegsE.AllowUserToAddRows = False
        Me.RegsE.AllowUserToDeleteRows = False
        Me.RegsE.AllowUserToOrderColumns = True
        Me.RegsE.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.RegsE.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.RegsE.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.ActiveCaption
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.RegsE.DefaultCellStyle = DataGridViewCellStyle3
        Me.RegsE.Location = New System.Drawing.Point(14, 215)
        Me.RegsE.Name = "RegsE"
        Me.RegsE.ReadOnly = True
        Me.RegsE.Size = New System.Drawing.Size(333, 434)
        Me.RegsE.TabIndex = 1
        '
        'RegsS
        '
        Me.RegsS.AllowUserToAddRows = False
        Me.RegsS.AllowUserToDeleteRows = False
        Me.RegsS.AllowUserToOrderColumns = True
        Me.RegsS.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.RegsS.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.RegsS.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.ActiveCaption
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.RegsS.DefaultCellStyle = DataGridViewCellStyle4
        Me.RegsS.Location = New System.Drawing.Point(354, 215)
        Me.RegsS.Name = "RegsS"
        Me.RegsS.ReadOnly = True
        Me.RegsS.Size = New System.Drawing.Size(334, 434)
        Me.RegsS.TabIndex = 6
        '
        'GBS
        '
        Me.GBS.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.GBS.Controls.Add(Me.LimS)
        Me.GBS.Controls.Add(Me.HHorS)
        Me.GBS.Controls.Add(Me.HFecS)
        Me.GBS.Controls.Add(Me.ReS)
        Me.GBS.Controls.Add(Me.AgS)
        Me.GBS.Controls.Add(Me.HorS)
        Me.GBS.Controls.Add(Me.FecS)
        Me.GBS.Controls.Add(Me.CanS)
        Me.GBS.Controls.Add(Me.LCanS)
        Me.GBS.Location = New System.Drawing.Point(355, 13)
        Me.GBS.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GBS.Name = "GBS"
        Me.GBS.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GBS.Size = New System.Drawing.Size(333, 194)
        Me.GBS.TabIndex = 5
        Me.GBS.TabStop = False
        Me.GBS.Text = "Salida de dinero"
        '
        'LimS
        '
        Me.LimS.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.LimS.Location = New System.Drawing.Point(242, 58)
        Me.LimS.Name = "LimS"
        Me.LimS.Size = New System.Drawing.Size(84, 31)
        Me.LimS.TabIndex = 6
        Me.LimS.Text = "Limpiar"
        Me.LimS.UseVisualStyleBackColor = True
        '
        'HHorS
        '
        Me.HHorS.AutoSize = True
        Me.HHorS.Location = New System.Drawing.Point(10, 95)
        Me.HHorS.Name = "HHorS"
        Me.HHorS.Size = New System.Drawing.Size(63, 23)
        Me.HHorS.TabIndex = 4
        Me.HHorS.Text = "Hora"
        Me.HHorS.UseVisualStyleBackColor = True
        '
        'HFecS
        '
        Me.HFecS.AutoSize = True
        Me.HFecS.Location = New System.Drawing.Point(10, 63)
        Me.HFecS.Name = "HFecS"
        Me.HFecS.Size = New System.Drawing.Size(68, 23)
        Me.HFecS.TabIndex = 2
        Me.HFecS.Text = "Fecha"
        Me.HFecS.UseVisualStyleBackColor = True
        '
        'ReS
        '
        Me.ReS.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ReS.Location = New System.Drawing.Point(150, 158)
        Me.ReS.Name = "ReS"
        Me.ReS.Size = New System.Drawing.Size(86, 31)
        Me.ReS.TabIndex = 8
        Me.ReS.Text = "Remover"
        Me.ReS.UseVisualStyleBackColor = True
        '
        'AgS
        '
        Me.AgS.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.AgS.Location = New System.Drawing.Point(8, 158)
        Me.AgS.Name = "AgS"
        Me.AgS.Size = New System.Drawing.Size(86, 31)
        Me.AgS.TabIndex = 7
        Me.AgS.Text = "Agregar"
        Me.AgS.UseVisualStyleBackColor = True
        '
        'HorS
        '
        Me.HorS.Enabled = False
        Me.HorS.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.HorS.Location = New System.Drawing.Point(86, 91)
        Me.HorS.Name = "HorS"
        Me.HorS.Size = New System.Drawing.Size(150, 26)
        Me.HorS.TabIndex = 5
        '
        'FecS
        '
        Me.FecS.Enabled = False
        Me.FecS.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.FecS.Location = New System.Drawing.Point(86, 58)
        Me.FecS.Name = "FecS"
        Me.FecS.Size = New System.Drawing.Size(150, 26)
        Me.FecS.TabIndex = 3
        '
        'CanS
        '
        Me.CanS.Location = New System.Drawing.Point(86, 26)
        Me.CanS.MaxLength = 11
        Me.CanS.Name = "CanS"
        Me.CanS.Size = New System.Drawing.Size(150, 26)
        Me.CanS.TabIndex = 1
        '
        'LCanS
        '
        Me.LCanS.AutoSize = True
        Me.LCanS.Location = New System.Drawing.Point(8, 29)
        Me.LCanS.Name = "LCanS"
        Me.LCanS.Size = New System.Drawing.Size(70, 19)
        Me.LCanS.TabIndex = 0
        Me.LCanS.Text = "Cantidad"
        '
        'EID
        '
        Me.EID.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.EID.Enabled = False
        Me.EID.Location = New System.Drawing.Point(98, 317)
        Me.EID.MaxLength = 11
        Me.EID.Name = "EID"
        Me.EID.Size = New System.Drawing.Size(150, 26)
        Me.EID.TabIndex = 3
        '
        'ESuma
        '
        Me.ESuma.AllowUserToAddRows = False
        Me.ESuma.AllowUserToDeleteRows = False
        Me.ESuma.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ESuma.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ESuma.Enabled = False
        Me.ESuma.Location = New System.Drawing.Point(98, 349)
        Me.ESuma.Name = "ESuma"
        Me.ESuma.ReadOnly = True
        Me.ESuma.Size = New System.Drawing.Size(150, 50)
        Me.ESuma.TabIndex = 4
        '
        'SID
        '
        Me.SID.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.SID.Enabled = False
        Me.SID.Location = New System.Drawing.Point(440, 317)
        Me.SID.MaxLength = 11
        Me.SID.Name = "SID"
        Me.SID.Size = New System.Drawing.Size(150, 26)
        Me.SID.TabIndex = 7
        '
        'SSuma
        '
        Me.SSuma.AllowUserToAddRows = False
        Me.SSuma.AllowUserToDeleteRows = False
        Me.SSuma.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.SSuma.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.SSuma.Enabled = False
        Me.SSuma.Location = New System.Drawing.Point(440, 349)
        Me.SSuma.Name = "SSuma"
        Me.SSuma.ReadOnly = True
        Me.SSuma.Size = New System.Drawing.Size(150, 50)
        Me.SSuma.TabIndex = 8
        '
        'Suma
        '
        Me.Suma.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Suma.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Suma.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.Suma.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator1, Me.LCanA, Me.LST, Me.ToolStripSeparator2, Me.ToolStripSeparator3, Me.Vaciar, Me.ToolStripSeparator4})
        Me.Suma.Location = New System.Drawing.Point(0, 652)
        Me.Suma.Name = "Suma"
        Me.Suma.Padding = New System.Windows.Forms.Padding(0, 0, 2, 0)
        Me.Suma.Size = New System.Drawing.Size(700, 26)
        Me.Suma.TabIndex = 10
        Me.Suma.Text = "ToolStrip1"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 26)
        '
        'LCanA
        '
        Me.LCanA.Name = "LCanA"
        Me.LCanA.Size = New System.Drawing.Size(123, 23)
        Me.LCanA.Text = "Cantidad actual: "
        '
        'LST
        '
        Me.LST.Name = "LST"
        Me.LST.Size = New System.Drawing.Size(0, 23)
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 26)
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 26)
        '
        'Vaciar
        '
        Me.Vaciar.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.Vaciar.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.Vaciar.Image = CType(resources.GetObject("Vaciar.Image"), System.Drawing.Image)
        Me.Vaciar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.Vaciar.Name = "Vaciar"
        Me.Vaciar.Size = New System.Drawing.Size(98, 23)
        Me.Vaciar.Text = "Vaciar datos"
        Me.Vaciar.ToolTipText = "Eliminar todos los datos (Entrada y salidas de dinero)"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 26)
        '
        'Princ
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.ClientSize = New System.Drawing.Size(700, 678)
        Me.Controls.Add(Me.RegsS)
        Me.Controls.Add(Me.RegsE)
        Me.Controls.Add(Me.Suma)
        Me.Controls.Add(Me.SSuma)
        Me.Controls.Add(Me.SID)
        Me.Controls.Add(Me.ESuma)
        Me.Controls.Add(Me.EID)
        Me.Controls.Add(Me.GBS)
        Me.Controls.Add(Me.GBE)
        Me.Font = New System.Drawing.Font("Cambria", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.Name = "Princ"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Princ"
        Me.GBE.ResumeLayout(False)
        Me.GBE.PerformLayout()
        CType(Me.RegsE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RegsS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GBS.ResumeLayout(False)
        Me.GBS.PerformLayout()
        CType(Me.ESuma, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SSuma, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Suma.ResumeLayout(False)
        Me.Suma.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GBE As GroupBox
    Friend WithEvents HorE As DateTimePicker
    Friend WithEvents FecE As DateTimePicker
    Friend WithEvents LCanE As Label
    Friend WithEvents Tiempo As Timer
    Friend WithEvents ReE As Button
    Friend WithEvents AgE As Button
    Friend WithEvents CanE As TextBox
    Friend WithEvents RegsE As DataGridView
    Friend WithEvents RegsS As DataGridView
    Friend WithEvents HHorE As CheckBox
    Friend WithEvents HFecE As CheckBox
    Friend WithEvents GBS As GroupBox
    Friend WithEvents HHorS As CheckBox
    Friend WithEvents HFecS As CheckBox
    Friend WithEvents ReS As Button
    Friend WithEvents AgS As Button
    Friend WithEvents HorS As DateTimePicker
    Friend WithEvents FecS As DateTimePicker
    Friend WithEvents CanS As TextBox
    Friend WithEvents LCanS As Label
    Friend WithEvents LimE As Button
    Friend WithEvents LimS As Button
    Friend WithEvents EID As TextBox
    Friend WithEvents ESuma As DataGridView
    Friend WithEvents SID As TextBox
    Friend WithEvents SSuma As DataGridView
    Friend WithEvents Suma As ToolStrip
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents LCanA As ToolStripLabel
    Friend WithEvents LST As ToolStripLabel
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents Vaciar As ToolStripButton
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
End Class
