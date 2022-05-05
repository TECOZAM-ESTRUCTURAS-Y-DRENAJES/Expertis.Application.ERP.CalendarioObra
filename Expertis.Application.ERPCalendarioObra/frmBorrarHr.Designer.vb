<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBorrarHr
    Inherits Solmicro.Expertis.Engine.UI.FormBase

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBorrarHr))
        Me.cbFecha = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.adOperario = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.adObra = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.btnCancelar = New Solmicro.Expertis.Engine.UI.Button
        Me.btnBorrar = New Solmicro.Expertis.Engine.UI.Button
        Me.lblFecha = New Solmicro.Expertis.Engine.UI.Label
        Me.lblOperario = New Solmicro.Expertis.Engine.UI.Label
        Me.lblObra = New Solmicro.Expertis.Engine.UI.Label
        Me.txtFecha = New Solmicro.Expertis.Engine.UI.TextBox
        Me.Label1 = New Solmicro.Expertis.Engine.UI.Label
        Me.SuspendLayout()
        '
        'cbFecha
        '
        Me.cbFecha.DisabledBackColor = System.Drawing.Color.White
        Me.cbFecha.Location = New System.Drawing.Point(96, 228)
        Me.cbFecha.Name = "cbFecha"
        Me.cbFecha.Size = New System.Drawing.Size(190, 21)
        Me.cbFecha.TabIndex = 0
        Me.cbFecha.Visible = False
        '
        'adOperario
        '
        Me.adOperario.DisabledBackColor = System.Drawing.Color.White
        Me.adOperario.DisplayField = "IDOperario"
        Me.adOperario.EntityName = "Operario"
        Me.adOperario.Location = New System.Drawing.Point(127, 105)
        Me.adOperario.Name = "adOperario"
        Me.adOperario.SecondaryDataFields = "IdOperario"
        Me.adOperario.Size = New System.Drawing.Size(190, 23)
        Me.adOperario.TabIndex = 1
        Me.adOperario.ViewName = "frmMntoOperario"
        '
        'adObra
        '
        Me.adObra.DisabledBackColor = System.Drawing.Color.White
        Me.adObra.DisplayField = "NObra"
        Me.adObra.EntityName = "ObraCabecera"
        Me.adObra.Location = New System.Drawing.Point(127, 167)
        Me.adObra.Name = "adObra"
        Me.adObra.SecondaryDataFields = "IDObra"
        Me.adObra.Size = New System.Drawing.Size(190, 23)
        Me.adObra.TabIndex = 2
        Me.adObra.ViewName = "tbObraCabecera"
        '
        'btnCancelar
        '
        Me.btnCancelar.Icon = CType(resources.GetObject("btnCancelar.Icon"), System.Drawing.Icon)
        Me.btnCancelar.Location = New System.Drawing.Point(72, 265)
        Me.btnCancelar.Name = "btnCancelar"
        Me.btnCancelar.Size = New System.Drawing.Size(87, 23)
        Me.btnCancelar.TabIndex = 3
        Me.btnCancelar.Text = "Cancelar"
        '
        'btnBorrar
        '
        Me.btnBorrar.Icon = CType(resources.GetObject("btnBorrar.Icon"), System.Drawing.Icon)
        Me.btnBorrar.Location = New System.Drawing.Point(215, 265)
        Me.btnBorrar.Name = "btnBorrar"
        Me.btnBorrar.Size = New System.Drawing.Size(87, 23)
        Me.btnBorrar.TabIndex = 4
        Me.btnBorrar.Text = "Borrar"
        '
        'lblFecha
        '
        Me.lblFecha.Location = New System.Drawing.Point(52, 51)
        Me.lblFecha.Name = "lblFecha"
        Me.lblFecha.Size = New System.Drawing.Size(40, 13)
        Me.lblFecha.TabIndex = 5
        Me.lblFecha.Text = "Fecha"
        '
        'lblOperario
        '
        Me.lblOperario.Location = New System.Drawing.Point(52, 110)
        Me.lblOperario.Name = "lblOperario"
        Me.lblOperario.Size = New System.Drawing.Size(57, 13)
        Me.lblOperario.TabIndex = 6
        Me.lblOperario.Text = "Operario"
        '
        'lblObra
        '
        Me.lblObra.Location = New System.Drawing.Point(52, 172)
        Me.lblObra.Name = "lblObra"
        Me.lblObra.Size = New System.Drawing.Size(35, 13)
        Me.lblObra.TabIndex = 7
        Me.lblObra.Text = "Obra"
        '
        'txtFecha
        '
        Me.txtFecha.DisabledBackColor = System.Drawing.Color.White
        Me.txtFecha.Location = New System.Drawing.Point(127, 51)
        Me.txtFecha.Name = "txtFecha"
        Me.txtFecha.Size = New System.Drawing.Size(190, 21)
        Me.txtFecha.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(127, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(137, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Formato: aaaa-mm-dd"
        '
        'frmBorrarHr
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(351, 333)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtFecha)
        Me.Controls.Add(Me.lblObra)
        Me.Controls.Add(Me.lblOperario)
        Me.Controls.Add(Me.lblFecha)
        Me.Controls.Add(Me.btnBorrar)
        Me.Controls.Add(Me.btnCancelar)
        Me.Controls.Add(Me.adObra)
        Me.Controls.Add(Me.adOperario)
        Me.Controls.Add(Me.cbFecha)
        Me.Name = "frmBorrarHr"
        Me.Text = "frmBorrarHr"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbFecha As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents adOperario As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents adObra As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents btnCancelar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents btnBorrar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents lblFecha As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents lblOperario As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents lblObra As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents txtFecha As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents Label1 As Solmicro.Expertis.Engine.UI.Label
End Class
