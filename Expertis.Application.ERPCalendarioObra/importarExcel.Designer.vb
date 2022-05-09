<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class importarExcel
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(importarExcel))
        Me.Label1 = New Solmicro.Expertis.Engine.UI.Label
        Me.PvProgreso = New System.Windows.Forms.ProgressBar
        Me.LProgreso = New Solmicro.Expertis.Engine.UI.Label
        Me.Label3 = New Solmicro.Expertis.Engine.UI.Label
        Me.lblRuta = New Solmicro.Expertis.Engine.UI.Label
        Me.cmdUbicacion = New Solmicro.Expertis.Engine.UI.Button
        Me.btnAceptar = New Solmicro.Expertis.Engine.UI.Button
        Me.btnSalir = New Solmicro.Expertis.Engine.UI.Button
        Me.borrahoras = New Solmicro.Expertis.Engine.UI.Button
        Me.CD = New System.Windows.Forms.OpenFileDialog
        Me.bBorrarExcel = New Solmicro.Expertis.Engine.UI.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(286, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(176, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "IMPORTACIÓN HORAS EXCEL"
        '
        'PvProgreso
        '
        Me.PvProgreso.Location = New System.Drawing.Point(188, 91)
        Me.PvProgreso.Name = "PvProgreso"
        Me.PvProgreso.Size = New System.Drawing.Size(574, 33)
        Me.PvProgreso.TabIndex = 1
        '
        'LProgreso
        '
        Me.LProgreso.Location = New System.Drawing.Point(188, 131)
        Me.LProgreso.Name = "LProgreso"
        Me.LProgreso.Size = New System.Drawing.Size(97, 13)
        Me.LProgreso.TabIndex = 2
        Me.LProgreso.Text = "Progreso Actual"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(31, 226)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(105, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Establecer Ruta: "
        '
        'lblRuta
        '
        Me.lblRuta.AutoSize = False
        Me.lblRuta.Location = New System.Drawing.Point(188, 226)
        Me.lblRuta.Name = "lblRuta"
        Me.lblRuta.Size = New System.Drawing.Size(574, 38)
        Me.lblRuta.TabIndex = 4
        Me.lblRuta.Text = "Ruta"
        '
        'cmdUbicacion
        '
        Me.cmdUbicacion.Icon = CType(resources.GetObject("cmdUbicacion.Icon"), System.Drawing.Icon)
        Me.cmdUbicacion.Location = New System.Drawing.Point(784, 226)
        Me.cmdUbicacion.Name = "cmdUbicacion"
        Me.cmdUbicacion.Size = New System.Drawing.Size(120, 38)
        Me.cmdUbicacion.TabIndex = 5
        Me.cmdUbicacion.Text = "Buscar"
        '
        'btnAceptar
        '
        Me.btnAceptar.Icon = CType(resources.GetObject("btnAceptar.Icon"), System.Drawing.Icon)
        Me.btnAceptar.Location = New System.Drawing.Point(191, 296)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(120, 38)
        Me.btnAceptar.TabIndex = 5
        Me.btnAceptar.Text = "Aceptar"
        '
        'btnSalir
        '
        Me.btnSalir.Icon = CType(resources.GetObject("btnSalir.Icon"), System.Drawing.Icon)
        Me.btnSalir.Location = New System.Drawing.Point(642, 296)
        Me.btnSalir.Name = "btnSalir"
        Me.btnSalir.Size = New System.Drawing.Size(120, 38)
        Me.btnSalir.TabIndex = 5
        Me.btnSalir.Text = "Salir"
        '
        'borrahoras
        '
        Me.borrahoras.Icon = CType(resources.GetObject("borrahoras.Icon"), System.Drawing.Icon)
        Me.borrahoras.Location = New System.Drawing.Point(191, 381)
        Me.borrahoras.Name = "borrahoras"
        Me.borrahoras.Size = New System.Drawing.Size(162, 38)
        Me.borrahoras.TabIndex = 5
        Me.borrahoras.Text = "Borrar(Por App y por Excel)"
        Me.borrahoras.Visible = False
        '
        'CD
        '
        Me.CD.FileName = "CD"
        '
        'bBorrarExcel
        '
        Me.bBorrarExcel.Icon = CType(resources.GetObject("bBorrarExcel.Icon"), System.Drawing.Icon)
        Me.bBorrarExcel.Location = New System.Drawing.Point(405, 296)
        Me.bBorrarExcel.Name = "bBorrarExcel"
        Me.bBorrarExcel.Size = New System.Drawing.Size(165, 38)
        Me.bBorrarExcel.TabIndex = 6
        Me.bBorrarExcel.Text = "Borrar horas"
        '
        'importarExcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(951, 459)
        Me.Controls.Add(Me.bBorrarExcel)
        Me.Controls.Add(Me.borrahoras)
        Me.Controls.Add(Me.btnSalir)
        Me.Controls.Add(Me.btnAceptar)
        Me.Controls.Add(Me.cmdUbicacion)
        Me.Controls.Add(Me.lblRuta)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.LProgreso)
        Me.Controls.Add(Me.PvProgreso)
        Me.Controls.Add(Me.Label1)
        Me.Name = "importarExcel"
        Me.Text = "importarExcel"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents PvProgreso As System.Windows.Forms.ProgressBar
    Friend WithEvents LProgreso As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label3 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents lblRuta As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents cmdUbicacion As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents btnAceptar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents btnSalir As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents borrahoras As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents CD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents bBorrarExcel As Solmicro.Expertis.Engine.UI.Button
End Class
