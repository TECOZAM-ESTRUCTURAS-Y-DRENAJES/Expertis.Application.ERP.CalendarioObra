Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Business.ClasesTecozam
Imports Solmicro.Expertis.Engine

Public Class frmObservaciones
    Inherits Solmicro.Expertis.Engine.UI.FormBase

    Public filtroHistorico As New Filter

    Private Sub cmdSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSalir.Click
        Me.Close()
    End Sub

    Private Sub cmdAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAceptar.Click
        GuardarDatos()
    End Sub

    Private Sub frmObservaciones_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rs As New DataTable
        Dim op As New OperarioCalendario

        rs = op.Filter(filtroHistorico)
        Dim contador As Integer = 0

        'If contador < rs.Rows.Count = False Then
        '    txtComentario.Text = rs.Rows(contador)("Comentario")
        '    contador = contador + 1
        'End If
        Try
            txtComentario.Text = rs.Rows(contador)("Comentario")
        Catch ex As Exception

        End Try

        'Libero memoria
        rs = Nothing
        filtroHistorico = Nothing
    End Sub
    Function GuardarDatos()
        Dim sSQL As String
        Dim op As New OperarioCalendario
        If MsgBox("¿Deseas realmente guardar los datos?", vbQuestion + vbYesNo) = vbNo Then
            Exit Function
        End If

        sSQL = "update tboperariocalendario set Comentario= '" & txtComentario.Text & "' WHERE " & Me.Params
        op.Ejecutar(sSQL)

        Me.Close()
    End Function

End Class