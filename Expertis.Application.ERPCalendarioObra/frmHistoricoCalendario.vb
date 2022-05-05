Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Business.ClasesTecozam

Public Class frmHistoricoCalendario
    Inherits Solmicro.Expertis.Engine.UI.FormBase


    Public operario As String
    Public mes As Integer
    Public anio As Integer

    Dim rs As New DataTable
    Dim f As New Filter
    Private Sub frmHistoricoCalendario_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim rs As New Recordset

        'rs = AdminData.GetData("tbHistoricoCalendario", , MyBase.Params)
        'Dim f As New Filter

        Dim cal As New HistoricoCalendario

        f.Add("idoperario", operario)
        f.Add("mes", mes)
        f.Add("anyo", anio)

        rs = cal.Filter(f)
        Grid1.DataSource = rs


    End Sub

    Private Sub cmdGuardar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGuardar.Click
        Grid1.UpdateData()
        'Me.UpdateData()
        Me.Close()
    End Sub

    Private Sub cmdCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancelar.Click
        Me.Close()
    End Sub
End Class
