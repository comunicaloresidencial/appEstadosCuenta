<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Generador
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.generar = New System.Windows.Forms.Button()
        Me.txtContrato = New System.Windows.Forms.TextBox()
        Me.gvContratos = New System.Windows.Forms.DataGridView()
        Me.todos = New System.Windows.Forms.Button()
        Me.dtpFecha = New System.Windows.Forms.DateTimePicker()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.btnEstado = New System.Windows.Forms.Button()
        Me.txtEstado = New System.Windows.Forms.TextBox()
        Me.regenerar_pdf = New System.Windows.Forms.Button()
        Me.enviar = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.txtEdoId = New System.Windows.Forms.TextBox()
        Me.btnReenvio = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.gvContratos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'generar
        '
        Me.generar.Location = New System.Drawing.Point(698, 12)
        Me.generar.Name = "generar"
        Me.generar.Size = New System.Drawing.Size(101, 23)
        Me.generar.TabIndex = 0
        Me.generar.Text = "Generar estado"
        Me.ToolTip1.SetToolTip(Me.generar, "Genera el estado de cuenta correspondiente al contrato indicado")
        Me.generar.UseVisualStyleBackColor = True
        '
        'txtContrato
        '
        Me.txtContrato.Location = New System.Drawing.Point(581, 15)
        Me.txtContrato.Name = "txtContrato"
        Me.txtContrato.Size = New System.Drawing.Size(111, 20)
        Me.txtContrato.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtContrato, "Número de contrato")
        '
        'gvContratos
        '
        Me.gvContratos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gvContratos.Location = New System.Drawing.Point(12, 12)
        Me.gvContratos.Name = "gvContratos"
        Me.gvContratos.Size = New System.Drawing.Size(559, 358)
        Me.gvContratos.TabIndex = 3
        '
        'todos
        '
        Me.todos.Location = New System.Drawing.Point(698, 57)
        Me.todos.Name = "todos"
        Me.todos.Size = New System.Drawing.Size(101, 23)
        Me.todos.TabIndex = 4
        Me.todos.Text = "TODOS"
        Me.todos.UseVisualStyleBackColor = True
        '
        'dtpFecha
        '
        Me.dtpFecha.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFecha.Location = New System.Drawing.Point(581, 60)
        Me.dtpFecha.Name = "dtpFecha"
        Me.dtpFecha.Size = New System.Drawing.Size(112, 20)
        Me.dtpFecha.TabIndex = 5
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(713, 141)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 10000
        '
        'btnEstado
        '
        Me.btnEstado.Location = New System.Drawing.Point(581, 196)
        Me.btnEstado.Name = "btnEstado"
        Me.btnEstado.Size = New System.Drawing.Size(207, 23)
        Me.btnEstado.TabIndex = 7
        Me.btnEstado.Text = "Regenerar Pago"
        Me.ToolTip1.SetToolTip(Me.btnEstado, "Regenera el pago online para tiendas")
        Me.btnEstado.UseVisualStyleBackColor = True
        '
        'txtEstado
        '
        Me.txtEstado.Location = New System.Drawing.Point(581, 170)
        Me.txtEstado.Name = "txtEstado"
        Me.txtEstado.Size = New System.Drawing.Size(207, 20)
        Me.txtEstado.TabIndex = 8
        '
        'regenerar_pdf
        '
        Me.regenerar_pdf.Location = New System.Drawing.Point(581, 225)
        Me.regenerar_pdf.Name = "regenerar_pdf"
        Me.regenerar_pdf.Size = New System.Drawing.Size(207, 23)
        Me.regenerar_pdf.TabIndex = 9
        Me.regenerar_pdf.Text = "Regenerar pdf"
        Me.ToolTip1.SetToolTip(Me.regenerar_pdf, "Regenera el pdf del estado cuenta indicado")
        Me.regenerar_pdf.UseVisualStyleBackColor = True
        '
        'enviar
        '
        Me.enviar.Location = New System.Drawing.Point(581, 254)
        Me.enviar.Name = "enviar"
        Me.enviar.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.enviar.Size = New System.Drawing.Size(207, 23)
        Me.enviar.TabIndex = 10
        Me.enviar.Text = "Enviar Alerta"
        Me.ToolTip1.SetToolTip(Me.enviar, "Vuelve a enviar el correo con el estado de cuenta")
        Me.enviar.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(581, 112)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(218, 23)
        Me.Button2.TabIndex = 11
        Me.Button2.Text = "Gerar Estados Cuenta Del Día"
        Me.ToolTip1.SetToolTip(Me.Button2, "Genera los estados de cuenta de la fecha actual")
        Me.Button2.UseVisualStyleBackColor = True
        '
        'txtEdoId
        '
        Me.txtEdoId.Location = New System.Drawing.Point(577, 350)
        Me.txtEdoId.Name = "txtEdoId"
        Me.txtEdoId.Size = New System.Drawing.Size(87, 20)
        Me.txtEdoId.TabIndex = 12
        '
        'btnReenvio
        '
        Me.btnReenvio.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.btnReenvio.Location = New System.Drawing.Point(713, 347)
        Me.btnReenvio.Name = "btnReenvio"
        Me.btnReenvio.Size = New System.Drawing.Size(75, 23)
        Me.btnReenvio.TabIndex = 13
        Me.btnReenvio.Text = "Reenviar"
        Me.ToolTip1.SetToolTip(Me.btnReenvio, "Regenera y envia los correos de los estados de cuenta con id mayor o igual al ing" &
        "resado a la izquierda")
        Me.btnReenvio.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(574, 334)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(105, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "ID Estado de cuenta"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(581, 154)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(105, 13)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "ID Estado de cuenta"
        '
        'Generador
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(807, 450)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnReenvio)
        Me.Controls.Add(Me.txtEdoId)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.enviar)
        Me.Controls.Add(Me.regenerar_pdf)
        Me.Controls.Add(Me.txtEstado)
        Me.Controls.Add(Me.btnEstado)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.dtpFecha)
        Me.Controls.Add(Me.todos)
        Me.Controls.Add(Me.gvContratos)
        Me.Controls.Add(Me.txtContrato)
        Me.Controls.Add(Me.generar)
        Me.Name = "Generador"
        Me.Text = "Generador estados de cuenta"
        CType(Me.gvContratos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents generar As Button
    Friend WithEvents txtContrato As TextBox
    Friend WithEvents gvContratos As DataGridView
    Friend WithEvents todos As Button
    Friend WithEvents dtpFecha As DateTimePicker
    Friend WithEvents Button1 As Button
    Friend WithEvents Timer1 As Timer
    Friend WithEvents btnEstado As Button
    Friend WithEvents txtEstado As TextBox
    Friend WithEvents regenerar_pdf As Button
    Friend WithEvents enviar As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents txtEdoId As TextBox
    Friend WithEvents btnReenvio As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents Label2 As Label
End Class
