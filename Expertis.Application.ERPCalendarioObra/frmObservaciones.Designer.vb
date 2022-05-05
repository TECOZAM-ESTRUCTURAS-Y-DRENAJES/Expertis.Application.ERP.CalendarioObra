<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmObservaciones
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmObservaciones))
        Me.txtComentario = New Solmicro.Expertis.Engine.UI.TextBox
        Me.cmdAceptar = New Solmicro.Expertis.Engine.UI.Button
        Me.cmdSalir = New Solmicro.Expertis.Engine.UI.Button
        Me.SuspendLayout()
        '
        'txtComentario
        '
        Me.txtComentario.DisabledBackColor = System.Drawing.Color.White
        Me.txtComentario.Location = New System.Drawing.Point(27, 24)
        Me.txtComentario.Multiline = True
        Me.txtComentario.Name = "txtComentario"
        Me.txtComentario.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtComentario.Size = New System.Drawing.Size(496, 214)
        Me.txtComentario.TabIndex = 0
        '
        'cmdAceptar
        '
        Me.cmdAceptar.Icon = CType(resources.GetObject("cmdAceptar.Icon"), System.Drawing.Icon)
        Me.cmdAceptar.Location = New System.Drawing.Point(136, 259)
        Me.cmdAceptar.Name = "cmdAceptar"
        Me.cmdAceptar.Size = New System.Drawing.Size(119, 44)
        Me.cmdAceptar.TabIndex = 1
        Me.cmdAceptar.Text = "Aceptar"
        '
        'cmdSalir
        '
        Me.cmdSalir.Icon = CType(resources.GetObject("cmdSalir.Icon"), System.Drawing.Icon)
        Me.cmdSalir.Location = New System.Drawing.Point(287, 259)
        Me.cmdSalir.Name = "cmdSalir"
        Me.cmdSalir.Size = New System.Drawing.Size(119, 44)
        Me.cmdSalir.TabIndex = 1
        Me.cmdSalir.Text = "Salir"
        '
        'frmObservaciones
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(547, 325)
        Me.Controls.Add(Me.cmdSalir)
        Me.Controls.Add(Me.cmdAceptar)
        Me.Controls.Add(Me.txtComentario)
        Me.Name = "frmObservaciones"
        Me.Text = "frmObservaciones"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtComentario As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents cmdAceptar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents cmdSalir As Solmicro.Expertis.Engine.UI.Button
End Class
