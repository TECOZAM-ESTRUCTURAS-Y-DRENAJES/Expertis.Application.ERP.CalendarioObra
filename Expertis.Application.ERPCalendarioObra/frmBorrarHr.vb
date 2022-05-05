Imports Solmicro.Expertis.Business.ClasesTecozam
Imports System.Windows.Forms
Imports Solmicro.Expertis.Business.Obra
Imports Solmicro.Expertis.Engine

Public Class frmBorrarHr

    Inherits Solmicro.Expertis.Engine.UI.FormBase

    Public operario As String
    Public obra As String

    Public Sub frmBorrarHr_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim dt As New DataTable

        adOperario.Text = operario
        adObra.Text = obra

    End Sub

    Private Sub btnCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelar.Click
        Me.Close()
    End Sub

    Private Sub btnBorrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBorrar.Click

        Dim op As New OperarioCalendario
        Dim ob As New ObraMODControl

        Dim fecha As Date '= cbFecha.Value
        fecha = txtFecha.Text
        Dim operario As String = adOperario.Text
        Dim obra As String = adObra.Text

        Dim dt As DataTable
        Dim f As New Filter
        'Obtengo el IDObra del NObra
        Dim DE As New Expertis.Engine.BE.DataEngine

        Dim dt2 As New DataTable

        dt2 = DE.Filter("tbObraCabecera", "IDObra", "NObra='" & obra & "'")

        obra = dt2.Rows(0)("IDObra").ToString
        'Obtenido
        Dim dSql As String = "DELETE FROM tbObraMODControl WHERE FechaInicio = '" & fecha & "' AND IDOperario = '" & operario & "' AND IDObra = '" & obra & "'"
        Dim sSQL As String = "SELECT * FROM tbObraMODControl WHERE FechaInicio = '" & fecha & "' AND IDOperario = '" & operario & "' AND IDObra = '" & obra & "'"

        f.Add("FechaInicio", FilterOperator.Equal, fecha)
        f.Add("IDOperario", FilterOperator.Equal, operario)
        f.Add("IDObra", FilterOperator.Equal, obra)

        dt = DE.Filter("tbObraMODControl", "", "FechaInicio = '" & fecha & "' AND IDOperario = '" & operario & "' AND IDObra = '" & obra & "'")

        'op.Ejecutar(dSql)

        If dt.Rows.Count > 0 Then
            op.Ejecutar(dSql)
            MessageBox.Show("Horas Eliminadas")
        Else
            MessageBox.Show("El operario no tiene registros para esa fecha y/o esa obra")
        End If

    End Sub
End Class