Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Business.Obra
Imports Solmicro.Expertis.Engine.UI
Imports Solmicro.Expertis.Engine.Global
Imports System.Math
Imports Solmicro.Expertis.Business.General
Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Business.ClasesTecozam
Imports System.Windows.Forms
Imports Solmicro.Expertis.Business
Imports System.Data.SqlClient

Public Class CalendarioObra
    Public sValor As String
    Public filtroHistorico As New Filter
    Friend dbDieta As Double = 0

    'Dim vIntervalo() As Intervalo
    Dim sObra As String 'Obra predeterminada
    Dim gHDM As Double 'Guarda las horas de menos
    Function LimpiarDatos(Optional ByVal bcargaCorrecta As Boolean = False)

        lbl0.Text = ""
        lbl1.Text = ""
        lbl2.Text = ""
        lbl3.Text = ""
        lbl4.Text = ""
        lbl5.Text = ""
        lbl6.Text = ""
        lbl7.Text = ""
        lbl8.Text = ""
        lbl9.Text = ""
        lbl10.Text = ""
        lbl11.Text = ""
        lbl12.Text = ""
        lbl13.Text = ""
        lbl14.Text = ""
        lbl15.Text = ""
        lbl16.Text = ""
        lbl17.Text = ""
        lbl18.Text = ""
        lbl19.Text = ""
        lbl20.Text = ""
        lbl21.Text = ""
        lbl22.Text = ""
        lbl23.Text = ""
        lbl24.Text = ""
        lbl25.Text = ""
        lbl26.Text = ""
        lbl27.Text = ""
        lbl28.Text = ""
        lbl29.Text = ""
        lbl30.Text = ""
        'lblExplicacion.Caption = ""
        'lblExtra.Caption = ""
        txtUltimaFB.Text = ""
        ntbDiasTrab1.Text = 0
        txtNHNormal.Text = 0
        txtNHExtra.Text = 0
        txtNHEspecial.Text = 0
        'txtIRPF.Text = 0
        txtNAnticipo.Text = 0
        txtTotal.Text = 0
        txtAtrasos.Text = 0
        txtPagare.Text = 0
        ' Valores de Dietas
        txtHorasD.Text = 0
        txtImporteD.Text = 0
        txtTotalRegD.Text = 0
        lblComentarios.BackColor = System.Drawing.Color.Gray
        txtHorasReg.BackColor = System.Drawing.Color.White
        txtPagaExtra.Text = 0
        txtPrima.Text = 0
        'David Velasco 15/11/21
        ntbDiasTrab2.Text = 0
        ntbDiasLab1.Text = 0
        ntbDiasLab2.Text = 0
        ntbHorasTeoricas.Text = 0
        txtHReg.Text = 0
        'Fin David
        ' Comentario
        If bcargaCorrecta = False Then txtCondiciones.Text = ""

    End Function

    Private Sub btnSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Private Sub CalendarioObra_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DE As New Expertis.Engine.BE.DataEngine
        Dim dtMes As New DataTable
        Dim dtAnio As New DataTable
        'Dim clsBE As New BE.DataEngine
        'Dim FilHist As New Filter
        'FilHist.Add("Mes", FilterOperator.Equal, cbxMes.Value)
        'FilHist.Add("anyo", FilterOperator.Equal, cbxAnyo.Value)
        'dtMes = clsBE.Filter("tbMes", "*")
        dtMes = DE.Filter("tbMes", "*", "")
        Me.cbxMes.DataSource = dtMes
        dtAnio = DE.Filter("tbAño", "*", "")
        Me.cbxAnyo.DataSource = dtAnio
        cbxMes.Value = Month(Today)
        cbxAnyo.Value = Year(Today)

        rellenaCombos()


    End Sub
    Public Sub rellenaCombos()
        clbACCDesde.Value = "01/01/2022"
        clbCCDesde.Value = "01/01/2022"
        clbVACAVDesde.Value = "01/01/2022"

        clbACCHasta.Value = "31/12/2022"
        clbCCHasta.Value = "31/12/2022"
        clbVACAVHasta.Value = "31/12/2022"
    End Sub
    Private Sub cbxMes_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxMes.ValueChanged
        ulMes.Text = SeleccionarMes(cbxMes.Value)
    End Sub
    Function SeleccionarMes(ByVal i As Integer) As String
        Select Case i
            Case 1
                SeleccionarMes = "Enero"
            Case 2
                SeleccionarMes = "Febrero"
            Case 3
                SeleccionarMes = "Marzo"
            Case 4
                SeleccionarMes = "Abril"
            Case 5
                SeleccionarMes = "Mayo"
            Case 6
                SeleccionarMes = "Junio"
            Case 7
                SeleccionarMes = "Julio"
            Case 8
                SeleccionarMes = "Agosto"
            Case 9
                SeleccionarMes = "Septiembre"
            Case 10
                SeleccionarMes = "Octubre"
            Case 11
                SeleccionarMes = "Noviembre"
            Case 12
                SeleccionarMes = "Diciembre"
        End Select
    End Function

    Private Sub advOperario1_SelectionChanged(ByVal sender As Object, ByVal e As Engine.UI.AdvSearchSelectionChangedEventArgs) Handles advOperario1.SelectionChanged
        Try
            ulOperario.Text = IIf(IsDBNull(e.Selected.Rows(0).Item("DescOperario")), "", e.Selected.Rows(0).Item("DescOperario"))
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Use 26/01/2009. Obtener texto condiciones especiales.
            txtCondiciones.Text = IIf(IsDBNull(e.Selected.Rows(0).Item("textoCondiciones")), "", e.Selected.Rows(0).Item("textoCondiciones"))
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            If Nz(e.Selected.Rows(0).Item("Incentivos"), False) = False Then
                chxIncentivo.Checked = False
            Else
                chxIncentivo.Checked = True
            End If
            ulJornada.Text = Nz(e.Selected.Rows(0).Item("jornada_laboral"), "")
            txtFechaAlta.Text = Nz(e.Selected.Rows(0).Item("fechaalta"), "")
            txtFechaBaja.Text = Nz(e.Selected.Rows(0).Item("Fecha_Baja"), "")
            ulTipoContrato.Text = Nz(e.Selected.Rows(0).Item("TipoContrato"), "")
        Catch ex As Exception

        End Try
    End Sub

    Function CargarDatos2()
        Dim op As New Operario
        Dim opHist As New ClasesTecozam.OperarioHistorico
        Dim dtOperario As DataTable
        Dim dt As DataTable
        Dim iContador As Integer
        Dim dFechaInicio As Date
        Dim dFechaFin As Date
        Dim tooltip As New Windows.Forms.ToolTip
        Dim ultFechaBaja As Date

        'Cargamos la última Fecha de Baja del trabajador
        Dim DE As New Expertis.Engine.BE.DataEngine

        'Dim filtroHistorico As New Filter
        'filtroHistorico.Add("IDOperario", advOperario1)
        'dt = opHist.Filter(filtroHistorico)
        dt = DE.Filter("tbOperarioHistorico", "FechaBaja", "IdOperario='" & advOperario1.Text & "'", "FechaBaja DESC")
        'ExpertisApp.GenerateMessage("el DT tiene " & dt.Rows.Count & " filas")
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                If Nz(dr("FechaBaja"), Nothing) > Nz(ultFechaBaja, Nothing) Then
                    ultFechaBaja = dr("FechaBaja")
                End If
            Next
            'Else
            '    ultFechaBaja = ""
        End If
        txtUltimaFB.Text = ultFechaBaja

        'Obtenemos los datos del Operario
        'Dim fOperario As New Filter
        'fOperario.Add("IdOperario", FilterOperator.Equal, FilterType.String)
        'dtOperario = op.Filter(fOperario)
        'ExpertisApp.GenerateMessage("el DT tiene " & dtOperario.Rows.Count & " filas")
        'dtOperario = DE.Filter("frmMntoOperario", "*", "IdOperario='" & advOperario1.Text & "'")
        'ExpertisApp.GenerateMessage("el DT tiene " & dtOperario.Rows.Count & " filas")

        Dim drOperario As DataRow = advOperario1.SelectedRow()
        'Obtenemos el Sueldo
        txtSueldo.Text = Round(Nz(drOperario("Sueldo"), 0), 2)
        'Obtenemos la Obra Predeterminada
        txtObraPredeterminada.Text = Nz(drOperario("Obra_Predeterminada"), "")
        Dim sObra = Nz(drOperario("Obra_Predeterminada"), "")

        Dim sSQL As String

        sSQL = "(Fecha >='01/" & cbxMes.Value & "/" & cbxAnyo.Value & "'" & _
               " and Fecha  <='20/" & cbxMes.Value & "/" & cbxAnyo.Value & "')" & _
               " and idcentro='" & sObra & "' and Tipodia = 1"
        Dim dtDias As DataTable

        dtDias = DE.Filter("tbCalendarioCentro", "*", sSQL)
        'ExpertisApp.GenerateMessage("el DTDias tiene " & dtDias.Rows.Count & " filas")

        'Obtenemos las horas 
        If dtDias.Rows.Count > 0 Then
            ntbDiasLab1.Text = (20 - dtDias.Rows.Count) * Nz(ulJornada.Text, 0)
        Else
            ntbDiasLab1.Text = (20 * ulJornada.Text)
        End If

        dFechaInicio = ObtenerFechaInicio()
        dFechaFin = ObtenerFechaFin()

        'Obtenemos los dias laborables (todos menos los festivos) (21 al 31 mes actual) "E"
        'teniendo en cuenta que se le puede dar de baja antes...

        sSQL = "(Fecha >='" & dFechaInicio & "' and Fecha  <='" & dFechaFin & "')" & _
               " and idcentro='" & sObra & "' and Tipodia = 1"
        'rs = AdminData.GetData("tbCalendarioCentro", , sSQL)

        Dim dtDLab As DataTable

        dtDLab = DE.Filter("tbCalendarioCentro", "*", sSQL)

        'COMENTADO DAVID VELASCO 15/3/22 NO SEGURO
        'If dtDLab.Rows.Count > 0 = False Then
        '    ntbDiasLab1.Text = ((DateDiff("d", dFechaInicio, dFechaFin) + 1) - dtDLab.Rows.Count) * ulJornada.Text
        'Else
        '    ntbDiasLab1.Text = (DateDiff("d", dFechaInicio, dFechaFin) + 1) * ulJornada.Text
        'End If
        'DAVID VELASCO 15/3/22

        'Obtenemos las horas trabajadas (01 al 20) "F"

        sSQL = "(FechaInicio >='01/" & cbxMes.Value & "/" & cbxAnyo.Value & "'" & _
             " and FechaInicio  <='20/" & cbxMes.Value & "/" & cbxAnyo.Value & "')" & _
             " and idoperario='" & advOperario1.Value & "' and idhora ='HO'"
        Dim dtHorTra As DataTable

        dtHorTra = DE.Filter("vAgrupadoMODHora", "*", sSQL)

        For Each drHorTra As DataRow In dtHorTra.Rows
            ntbDiasTrab1.Text = CDbl(Nz(ntbDiasTrab1.Text, 0)) + CDbl(drHorTra("Horas"))

        Next


        'ATENCION: Se supone que los días trabajados son igual que los días laborables...
        'COMENTADO POR DAVID VELASCO 15/3/22
        'If BuscarHorasTeoricas() Then
        '    ntbDiasTrab2.Text = ntbDiasLab2.Text
        'Else
        '    ntbDiasTrab2.Text = 0
        'End If
        '15/3/22

        'ATENCION ESTE ES UN CALCULO QUE VUELVO HACER PORQUE LOS DIAS LABORABLES NO ES = A LOS DIAS TEORICAMENTE TRABAJADOS
        dFechaFin = BuscarDia("01/" & cbxMes.Value & "/" & cbxAnyo.Value)

        sSQL = "(Fecha >='" & dFechaInicio & "' and Fecha  <='" & dFechaFin & "')" & _
               " and idcentro='" & sObra & "' and Tipodia = 1"


        Dim dtDLab2 As DataTable = DE.Filter("tbCalendarioCentro", "*", sSQL)

        If dtDLab2.Rows.Count > 0 Then
            ntbDiasLab2.Text = ((DateDiff("d", dFechaInicio, dFechaFin) + 1) - dtDLab2.Rows.Count) * ulJornada.Text
        Else
            ntbDiasLab2.Text = (DateDiff("d", dFechaInicio, dFechaFin) + 1) * ulJornada.Text
        End If

        'DAVID VELASCO 15/3/22 NO ESTOY SEGURO
        If BuscarHorasTeoricas() Then
            ntbDiasTrab2.Text = ntbDiasLab2.Text
        Else
            ntbDiasTrab2.Text = 0
        End If
        'DAVID VELASCO

        '1
        ntbDiasTot1.Text = Round((txtSueldo.Text / (CDbl(ntbDiasLab1.Text) + CDbl(ntbDiasLab2.Text))) * ntbDiasTrab1.Text, 2)

        '2

        ntbDiasTot2.Text = Round((txtSueldo.Text / (CDbl(ntbDiasLab1.Text) + CDbl(ntbDiasLab2.Text))) * ntbDiasTrab2.Text, 2)

        'H = (C/(D+E)) * (F+G)

        ntbDiasTot.Text = Round(CDbl(ntbDiasTot1.Text) + CDbl(ntbDiasTot2.Text), 2)

        'Obtengo las horas correspondientes al mes
        'sSQL = "(FechaInicio >='21/" & MesAnterior(cbxMes.Value, cbxAnyo.Value) & _
        '       " and FechaInicio <='20/" & cbxMes.Value & "/" & cbxAnyo.Value & "')" & _
        '       " and idoperario='" & advOperario1.Value & "'"
        Dim fHorReal As New Filter
        fHorReal.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, ("21/" & MesAnterior(cbxMes.Value, cbxAnyo.Value)))
        fHorReal.Add("FechaInicio", FilterOperator.LessThanOrEqual, ("20/" & cbxMes.Value & "/" & cbxAnyo.Value))
        fHorReal.Add("IDOperario", FilterOperator.Equal, advOperario1.Value)
        Dim dtHorReal = DE.Filter("vAgrupadoMODHora", fHorReal)
        'DE.Filter("vAgrupadoMODHora", "*", sSQL)

        'sumo las horas del trabajador segun la hora
        'Do While Not rs.EOF
        For Each drHorReal As DataRow In dtHorReal.Rows

            Select Case drHorReal("idhora")
                'Hora ordinaria
                'Case "HO"
                '"I"
                '    txtNHNormal.Text = txtNHNormal.Text + Nz(rs!Horas, 0)
                'Hora extra
                Case "HX"
                    '"J"
                    txtNHExtra.Text = txtNHExtra.Text + Nz(drHorReal("Horas"), 0)
                    'Hora especial
                Case "HE"
                    '"K"
                    txtNHEspecial.Text = txtNHEspecial.Text + Nz(drHorReal("Horas"), 0)
            End Select

        Next

        'sSQL = "(FechaInicio >='01/" & cbxMes.Value & "/" & cbxAnyo.Value & "'" & _
        '     " and FechaInicio  <='20/" & cbxMes.Value & "/" & cbxAnyo.Value & "')" & _
        '     " and idoperario='" & advOperario1.Value & "' and idhora='HO'"
        Dim clsBE As New BE.DataEngine
        Dim Filtro As New Filter
        'Dim FIDes As Date = "01/" & cbxMes.Value & "/" & cbxAnyo.Value & "'"
        Filtro.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, "01/" & cbxMes.Value & "/" & cbxAnyo.Value)
        Filtro.Add("FechaInicio", FilterOperator.LessThanOrEqual, "20/" & cbxMes.Value & "/" & cbxAnyo.Value)
        Filtro.Add("IDOperario", FilterOperator.Equal, advOperario1.Text)
        Filtro.Add("IDhora", FilterOperator.Equal, "HO")
        Dim dtHorOrd = clsBE.Filter("vAgrupadoMODHora", Filtro)


        'Dim dtHorOrd = DE.Filter("vAgrupadoMODHora", "*", sSQL)
        'Dim dtHorOrd = DE.Filter("vAgrupadoMODHora", "*", sSQL)

        tooltip.SetToolTip(txtNHNormal, "")

        'Hora ordinaria
        '"I"
        For Each drHoraOrd As DataRow In dtHorOrd.Rows

            txtNHNormal.Text = txtNHNormal.Text + Nz(drHoraOrd("Horas"), 0)

        Next
        tooltip.SetToolTip(txtNHNormal, "(Horas 1 - 31: " & txtNHNormal.Text & ")")

        '"L"
        txtNHNormalImporte.Text = Round(Nz(drOperario("c_h_n"), 0), 2)
        '"M"
        txtNHExtraImporte.Text = Round(Nz(drOperario("c_h_x"), 0), 2)
        '"N"
        txtNHEspecialImporte.Text = Round(Nz(drOperario("c_h_e"), 0), 2)
        'Calculo de horas por importe
        txtNHNormalTotal.Text = Round(txtNHNormal.Text * txtNHNormalImporte.Text, 2)
        txtNHExtraTotal.Text = Round(txtNHExtra.Text * txtNHExtraImporte.Text, 2)
        txtNHEspecialTotal.Text = Round(txtNHEspecial.Text * txtNHEspecialImporte.Text, 2)
        'Mostrar información del operario
        txtPlus.Text = Round(Nz(drOperario("Plus"), 0), 2)
        'Dieta - modificado manuel 18/06/08
        txtDieta.Text = Round((Nz(drOperario("Dieta"), 0) / (CDbl(ntbDiasLab1.Text) + CDbl(ntbDiasLab2.Text))) * (CDbl(ntbDiasTrab1.Text) + CDbl(ntbDiasTrab2.Text)), 2)
        If txtDieta.Text > Nz(drOperario("Dieta"), 0) Then
            txtDieta.Text = Round(Nz(drOperario("Dieta"), 0), 2)
        End If
        ' Coger el valor de la dieta para regularizar posteriormente
        Dim dbDieta As Double = Round(Nz(drOperario("Dieta"), 0), 2)
        txtVarios.Text = Round(Nz(drOperario("Varios"), 0), 2)

        'A RELLENAR POR EL USUARIO
        txtNMensual.Text = "0"
        txtNAnticipo.Text = "0"
        txtSSocial.Text = "0"
        txtAnticipo.Text = 0
        txtVarios2.Text = 0
        txtAtrasos.Text = "0"
        txtPagare.Text = "0"
        txtPagaExtra.Text = "0"
        txtPrima.Text = "0"

        txtACC.Text = "0"
        txtACCImporte.Text = "0"
        txtACCTotal.Text = "0"
        txtCC.Text = "0"
        txtCCImporte.Text = "0"
        txtCCTotal.Text = "0"
        txtVACAV.Text = "0"
        txtVACAVImporte.Text = "0"
        txtVACAVTotal.Text = "0"

        '----------------------
        ntbAmortizacion.Text = Round(Nz(drOperario("amortizacion"), 0), 2)
        txtEmbargos.Text = Round(Nz(drOperario("embargo"), 0), 2)
        txtAnticipoFijo.Text = Round(Nz(drOperario("anticipo_fijo"), 0), 2)
        txtObraPredeterminada.Text = BuscarObraPredet(sObra)

        dFechaInicio = ObtenerFechaInicio()

        sSQL = "(Fecha >='" & dFechaInicio & "'" & _
               " and Fecha  <='" & BuscarDia("01/" & cbxMes.Value & "/" & cbxAnyo.Value) & "')" & _
               " and idcentro='" & sObra & "' and Tipodia = 1"
        ' rs = AdminData.GetData("tbCalendarioCentro", , sSQL)
        Dim dtA = DE.Filter("tbCalendarioCentro", "*", sSQL)
        dFechaFin = BuscarDia("01/" & cbxMes.Value & "/" & cbxAnyo.Value)
        'Esto evita que rellene datos de un mes que no corresponde

        If BuscarHorasTeoricas() Then
            If dtA.Rows.Count > 0 Then
                ntbHorasTeoricas.Text = (DateDiff("d", dFechaInicio, dFechaFin) + 1 - dtA.Rows.Count) * ulJornada.Text
            Else
                ntbHorasTeoricas.Text = (DateDiff("d", dFechaInicio, dFechaFin) + 1) * ulJornada.Text
            End If
        Else
            ntbHorasTeoricas.Text = 0
        End If

        tooltip.SetToolTip(txtNHNormal, tooltip.GetToolTip(txtNHNormal) & " (Horas Teoricas: " & ntbHorasTeoricas.Text & ")")

        ''Modifico las horas ordinarias **** no suma horas teoricas, si no horas reales teoricas
        txtNHNormal.Text = CDbl(txtNHNormal.Text) + CDbl(ntbDiasTrab2.Text) '+ HorasDeMenos(sObra)
        HorasDeMenos(sObra)
        ''Recalculo de horas por importe
        txtNHNormalTotal.Text = txtNHNormal.Text * txtNHNormalImporte.Text

        ''Dias laborables del mes anterior
        sSQL = "(Fecha >='01/" & MesAnterior(cbxMes.Value, cbxAnyo.Value) & "'" & _
               " and Fecha <='" & BuscarDia("01/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "'", "")) & "')" & _
              " and idcentro='" & sObra & "' and Tipodia = 1"
        'MsgBox(sSQL)
        'rs = AdminData.GetData("tbCalendarioCentro", , sSQL)
        Dim dtDLabMA = DE.Filter("tbCalendarioCentro", "*", sSQL)
        dFechaInicio = "01/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "'", "")
        dFechaFin = BuscarDia("01/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "'", ""))

        If txtNHNormalImporte.Text = 0 Then
            txtImpReg.Text = Round(Round(((txtSueldo.Text / (DateDiff("d", dFechaInicio, dFechaFin) - dtDLabMA.Rows.Count + 1))), 2) / ulJornada.Text, 2)
        Else
            txtImpReg.Text = txtNHNormalImporte.Text
        End If
        txtTotalReg.Text = Round(CDbl(txtHorasReg.Text) * CDbl(txtImpReg.Text), 2)

    End Function

    Private Function ObtenerFechaInicio() As Date
        Dim dFechaInicio As Date
        dFechaInicio = "21/" & cbxMes.Value & "/" & cbxAnyo.Value
        'Obtenemos la fecha de inicio real
        If Len(txtFechaAlta.Text) > 0 Then
            If cbxMes.Value = Trim(Month(txtFechaAlta.Text)) And cbxAnyo.Value = Trim(Year(txtFechaAlta.Text)) And Trim(Day(txtFechaAlta.Text)) > 21 Then
                dFechaInicio = txtFechaAlta.Text
                If Len(Nz(txtUltimaFB.Text, "")) > 0 Then
                    If DateDiff("d", txtUltimaFB.Text, dFechaInicio) >= 1 And DateDiff("d", txtUltimaFB.Text, dFechaInicio) <= 3 Then dFechaInicio = "21/" & cbxMes.Value & "/" & cbxAnyo.Value
                End If
            End If
        End If
        ObtenerFechaInicio = dFechaInicio
    End Function

    Private Function ObtenerFechaFin() As Date
        Dim fechafinal As Date
        fechafinal = "01/" & cbxMes.Value & "/" & cbxAnyo.Value
        'Obtenemos el último dia del mes
        fechafinal = DateAdd(DateInterval.Month, 1, fechafinal)
        fechafinal = DateAdd(DateInterval.Day, -1, fechafinal)
        'Obtenemos la fecha de fin real
        If Len(txtFechaBaja.Text) > 0 And IsDate(txtFechaBaja.Text) Then
            If cbxMes.Value = Trim(Month(txtFechaBaja.Text)) And cbxAnyo.Value = Trim(Year(txtFechaBaja.Text)) Then
                'Ahora verificamos el día
                If Trim(Day(txtFechaBaja.Text)) > 20 Then
                    fechafinal = txtFechaBaja.Text
                Else
                    fechafinal = "20/" & cbxMes.Value & "/" & cbxAnyo.Value
                End If
            End If
        End If
        ObtenerFechaFin = fechafinal
    End Function

    Private Function ObtenerFechaInicioMesAnterior() As Date
        Dim dFechaA As Date
        Dim dFechaInicio As Date
        dFechaA = "21/" & cbxMes.Value & "/" & cbxAnyo.Value
        dFechaInicio = DateAdd("m", -1, dFechaA)
        'Obtenemos la fecha de inicio real
        If Len(txtFechaAlta.Text) > 0 Then
            If Trim(Month(dFechaInicio)) = Trim(Month(txtFechaAlta.Text)) And Trim(Year(dFechaInicio)) = Trim(Year(txtFechaAlta.Text)) And Trim(Day(txtFechaAlta.Text)) > 21 Then
                dFechaInicio = txtFechaAlta.Text
            End If
        End If
        ObtenerFechaInicioMesAnterior = dFechaInicio
    End Function

    Private Function BuscarHorasdeMenos() As Boolean
        Dim dFechaA As Date
        BuscarHorasdeMenos = True
        dFechaA = "21/" & cbxMes.Value & "/" & cbxAnyo.Value
        dFechaA = DateAdd("m", -1, dFechaA)
        If Len(txtFechaAlta.Text) > 0 Then
            If CDate(txtFechaAlta.Text) > dFechaA Then BuscarHorasdeMenos = False
        End If

    End Function

    Private Function BuscarHorasTeoricas() As Boolean
        Dim dFechaA As Date
        BuscarHorasTeoricas = True
        dFechaA = "01/" & cbxMes.Value & "/" & cbxAnyo.Value
        dFechaA = DateAdd("m", 1, dFechaA)
        dFechaA = DateAdd("d", -1, dFechaA)

        '1 Con la fecha de alta
        If Len(txtFechaAlta.Text) > 0 Then
            If CDate(txtFechaAlta.Text) > dFechaA Then BuscarHorasTeoricas = False
        End If

        dFechaA = "01/" & cbxMes.Value & "/" & cbxAnyo.Value
        If Len(txtFechaBaja.Text) > 0 Then
            If CDate(txtFechaBaja.Text) > dFechaA Then BuscarHorasTeoricas = False
        End If
    End Function

    Function Month(ByVal Fecha As Date) As Integer
        Month = Fecha.Month
    End Function

    Function Day(ByVal Fecha As Date) As Integer
        Day = Fecha.Day
    End Function

    Function BuscarDia(ByVal dFecha As Date) As Date
        Select Case Month(dFecha)
            Case 1, 3, 5, 7, 8, 10, 12
                BuscarDia = "31/" & Month(dFecha) & "/" & Year(dFecha)
            Case 4, 6, 9, 11
                BuscarDia = "30/" & Month(dFecha) & "/" & Year(dFecha)
            Case 2
                'miro haber si este mes tiene 28 dias o 29 para calcular bien los Dias laborables.

                'fechafinal = Day(ObtenerFechaFin)
                BuscarDia = DateAdd(DateInterval.Day, -1, CDate("01/" & Month(dFecha) + 1 & "/" & Year(dFecha)))
        End Select
    End Function

    Function MesAnterior(ByVal MES As Integer, ByVal Ano As Integer) As String
        Select Case MES
            Case 1
                MesAnterior = "12/" & Ano - 1
            Case Else
                MesAnterior = MES - 1 & "/" & Ano
        End Select
    End Function

    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAceptar.Click
        If advOperario1.Text <> "" And cbxMes.Value Is Nothing = False And cbxAnyo.Value Is Nothing = False Then
            If Nz(ulJornada.Text, 0) <> 0 Then

                LimpiarDatos(True)
                CargarDatos2()
                CargarDatos()
                LimpiarFestivos()
                DiasFestivos()
                CargarDatosGuardados()
                aviso9m(advOperario1.Text, cbxMes.Value, cbxAnyo.Value)
            Else
                LimpiarDatos()
                LimpiarFestivos()
                MsgBox("El empleado NO tiene jornada laboral", vbExclamation + vbOKOnly, "Criterios")
            End If
        Else
            MsgBox("Debes seleccionar un Operario, un Mes y un Año", vbInformation + vbOKOnly, "Criterios")
        End If
        If txtHorasReg.Text < 0 Then
            txtHorasReg.BackColor = System.Drawing.Color.Red
        End If
    End Sub

    Function BuscarObraPredet(ByVal sObra As String) As String
        Dim fwnOC As ObraCabecera
        Dim sSQL As String
        Dim dt As DataTable
        fwnOC = New ObraCabecera
        sSQL = "Idobra='" & sObra & "'"
        dt = fwnOC.Filter("*", sSQL)
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                BuscarObraPredet = Nz(dr("descobra"), "No tiene obra predeterminada")
            Next
        Else
            BuscarObraPredet = "No tiene obra predeterminada"
        End If
    End Function

    Function HorasDeMenos(ByVal sObra As String) As Double
        Dim sSQL As String
        Dim dtHoras As DataTable
        'Dim rs As New Recordset
        Dim dblHoras As Double
        Dim dblHorasTeoricas As Double
        Dim dFechaInicio As Date
        Dim dFechaFin As Date
        Dim DE As New Expertis.Engine.BE.DataEngine

        ' Obtener las horas ordinarias entre el 21 y el último día del mes anterior
        sSQL = "(FechaInicio >='21/" & MesAnterior(cbxMes.Value, cbxAnyo.Value) & _
               " and FechaInicio  <='" & BuscarDia("01/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "'", "")) & "')" & _
               " and idoperario='" & advOperario1.Value & "' and idhora='HO'"
        Dim Filtro As New Filter
        Filtro.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, "21/" & MesAnterior(cbxMes.Value, cbxAnyo.Value))
        Filtro.Add("FechaInicio", FilterOperator.LessThanOrEqual, BuscarDia("01/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "", "")))
        Filtro.Add("IDOperario", FilterOperator.Equal, advOperario1.Value)
        Filtro.Add("IDHora", FilterOperator.Equal, "HO")
        dtHoras = DE.Filter("vAgrupadoMODHora", Filtro)
        'dtHoras = AdminData.GetData("vAgrupadoMODHora", , sSQL)

        'Horas Ordinarias del mes anterior
        'Do While Not rs.EOF
        If dtHoras.Rows.Count > 0 Then
            For Each drhoras As DataRow In dtHoras.Rows
                dblHoras = dblHoras + drhoras("Horas")
            Next
        End If

        'Loop

        'Compruebo si se le ha dado de alta en el mes anterior
        '**anterior
        If "21/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "'", "") >= txtFechaAlta.Text And BuscarDia("01/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "'", "")) <= txtFechaAlta.Text Then
            sSQL = "(Fecha >='" & txtFechaAlta.Text & _
               "' and Fecha <='" & BuscarDia("01/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "'", "")) & "')" & _
               " and idcentro='" & sObra & "' and Tipodia = 1"
            dFechaInicio = txtFechaAlta.Text
        Else
            If cbxMes.Value = Trim(Month(txtFechaAlta.Text)) And cbxAnyo.Value = Trim(Year(txtFechaAlta.Text)) Then
                'No tiene horas del mes anterior, porque se le ha dado de alta en este mes
                sSQL = "(Fecha >='21/" & MesAnterior(cbxMes.Value, cbxAnyo.Value) & _
                   " and Fecha <='" & BuscarDia("01/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "'", "")) & "')" & _
                   " and idcentro='" & sObra & "' and Tipodia = 1"
                dFechaInicio = DateAdd(DateInterval.Day, -1, BuscarDia("01/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "'", "")))
            Else
                sSQL = "(Fecha >='21/" & MesAnterior(cbxMes.Value, cbxAnyo.Value) & _
                   " and Fecha <='" & BuscarDia("01/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "'", "")) & "')" & _
                   " and idcentro='" & sObra & "' and Tipodia = 1"
                dFechaInicio = "21/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "'", "")
            End If
        End If
        '*******quita lo anterior
        If BuscarHorasdeMenos() Then
            ' Desde el primer día del mes anterior (21 o fecha de alta) hasta el último.
            dFechaInicio = ObtenerFechaInicioMesAnterior()
            dFechaFin = BuscarDia("01/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "'", ""))

            'Obtengo sql
            sSQL = "(Fecha >='" & dFechaInicio & " '" & _
                       " and Fecha <='" & dFechaFin & "')" & _
                       " and idcentro='" & sObra & "' and Tipodia = 1"

            ' rs = AdminData.GetData("tbCalendarioCentro", , sSQL)
            Dim dtHorMen = DE.Filter("tbCalendarioCentro", "*", sSQL)
            'Horas Teoricas mes Anterior (- festivos)
            If dtHorMen.Rows.Count > 0 Then
                dblHorasTeoricas = (DateDiff("d", dFechaInicio, dFechaFin) + 1 - dtHorMen.Rows.Count) * ulJornada.Text
            Else
                dblHorasTeoricas = (DateDiff("d", dFechaInicio, dFechaFin) + 1) * ulJornada.Text
            End If
            'Horas realmente trabajadas - las horas que tenia que haber trabajado
            HorasDeMenos = dblHoras - dblHorasTeoricas
            txtHReg.Text = dblHorasTeoricas

            ' Control de regularización de importe de dietas
            If dbDieta > 0 Then
                Dim dblHorasDietaEstimadas, DbLHorasResta As Double
                dblHorasDietaEstimadas = HorasCalendarioCentro(CDate("21/" & MesAnterior(cbxMes.Value, cbxAnyo.Value)), BuscarDia("01/" & Replace(MesAnterior(cbxMes.Value, cbxAnyo.Value), "'", "")))
                ' Ahora horas trabajadas desde el 21-31
                DbLHorasResta = dblHoras - dblHorasDietaEstimadas
                ' Control si es negativo
                If DbLHorasResta < 0 Then
                    txtHorasD.Text = DbLHorasResta
                    txtImporteD.Text = Round(dbDieta / (dblHorasDietaEstimadas + HorasCalendarioCentro(CDate("01/" & MesAnterior(cbxMes.Value, cbxAnyo.Value)), CDate("20/" & MesAnterior(cbxMes.Value, cbxAnyo.Value)))), 6)
                    txtTotalRegD.Text = Round(CDbl(txtHorasD.Text) * CDbl(txtImporteD.Text), 2)
                End If
            End If
        Else
            HorasDeMenos = 0
            txtHReg.Text = 0
            ' Sin horas de menos dieta de menos = 0 siempre
            txtTotalRegD.Text = 0
        End If
        gHDM = HorasDeMenos

        txtHorasRegC.Text = HorasDeMenos
        txtHorasReg.Text = HorasDeMenos
    End Function

    Function HorasCalendarioCentro(ByVal dDesde As Date, ByVal dHasta As Date) As Double
        ' Función que optiene el Nº de días laborales de cada centro entre unas fechas dadas. 02/02/2010
        Dim DE As New Expertis.Engine.BE.DataEngine
        Try
            Dim dtDias As DataTable
            Dim Ssql As String

            Ssql = "(Fecha >='" & dDesde & " '" & _
              " and Fecha  <='" & dHasta & "')" & _
              " and idcentro='" & sObra & "' and Tipodia = 1"
            dtDias = DE.Filter("tbCalendarioCentro", "*", Ssql)
            ' Control de días 1ª 15na o 2ª
            If dDesde.Day = 1 Then
                If dtDias.Rows.Count > 0 Then
                    HorasCalendarioCentro = (20 - dtDias.Rows.Count) * Nz(ulJornada.Text, 0)
                Else
                    HorasCalendarioCentro = 20 * ulJornada.Text
                End If
            Else
                ' Controlar El final de mes
                Dim ShDia As Short = 0
                'Select Case dDesde.DaysInMonth(dDesde.Year, dDesde.Month)
                Select Case Date.DaysInMonth(dDesde.Year, dDesde.Month)
                    Case 28
                        ShDia = 8
                    Case 29
                        ShDia = 9
                    Case 30
                        ShDia = 10
                    Case 31
                        ShDia = 11
                End Select
                If dtDias.Rows.Count > 0 Then
                    HorasCalendarioCentro = (ShDia - dtDias.Rows.Count) * Nz(ulJornada.Text, 0)
                Else
                    HorasCalendarioCentro = ShDia * ulJornada.Text
                End If
            End If
        Catch ex As Exception
            ExpertisApp.GenerateMessage("Error al calcular las Horas del calendario." & ex.Message, MsgBoxStyle.Exclamation, "Error en cálculo.")
            'MsgBox("Error al calcular las Horas del calendario." & ex.Message, MsgBoxStyle.Exclamation, "Error en cálculo.")
            HorasCalendarioCentro = 0
        End Try
    End Function

    Function CargarDatos()
        Dim fwnMOD As ObraMODControl
        Dim sSQL As String
        Dim dt As DataTable
        Dim I As Integer 'Sirve para decidir si va dentro del cuadro o se usa con el "+"
        Dim dFechaAnterior As Date
        Dim dblHoras As Double

        Dim iLaborables As Integer
        Dim iTrabajados As Integer
        Dim sCodigo As String
        Dim DE As New Engine.BE.DataEngine
        Dim codOperario As String = advOperario1.Text
        Dim f As New Filter

        'ExpertisApp.GenerateMessage("El operario seleccionado es: " & codOperario)

        fwnMOD = New ObraMODControl
        '+++++++++++++++++++++++++++++++++++++++++++++  
        ''CODIGO ORIGINAL
        'Dim fECHAfINAL As Date
        'fECHAfINAL = "01/" & cbxMes.Value & "/" & cbxAnyo.Value
        'Obtenemos el último dia del mes
        'fECHAfINAL = DateAdd(DateInterval.Month, 1, fECHAfINAL)
        'fECHAfINAL = DateAdd(DateInterval.Day, -1, fECHAfINAL)

        'sSQL = "(FechaInicio >='01/" & cbxMes.Value & "/" & cbxAnyo.Value & _
        '      "' and FechaInicio <='" & Day(fECHAfINAL) & "/" & cbxMes.Value & "/" & cbxAnyo.Value & "')" & _
        '     " and idoperario='" & advOperario1.Value & "'"
        '++++++++++++++++++++++++++++++++++++++++++++
        sSQL = "(FechaInicio >='21/" & MesAnterior(cbxMes.Value, cbxAnyo.Value) & "'" & _
        " and FechaInicio <='20/" & cbxMes.Value & "/" & cbxAnyo.Value & "')" & _
        " and idoperario='" & advOperario1.Text & "'"

        f.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, "21/" & MesAnterior(cbxMes.Value, cbxAnyo.Value))
        f.Add("FechaInicio", FilterOperator.LessThanOrEqual, "20/" & cbxMes.Value & "/" & cbxAnyo.Value)
        f.Add("IDOperario", FilterOperator.Equal, advOperario1.Value)

        dt = DE.Filter("vAgrupadoMOD", f, "*")
        I = 0

        Dim contar As Integer = 21
        Dim contador As Integer = 0

        If dt.Rows.Count > 0 Then
            'Do While Not rs.EOF
            For Each dr As DataRow In dt.Rows
                dblHoras = dblHoras + dr("Horas")
                'Si es cero quiere decir que es V, FI, etc...
                If dr("Horas") = 0 Then
                    sCodigo = BuscarCodigo(dr("IdObra"), dr("idoperario"), dr("FechaInicio"))
                Else
                    sCodigo = ""
                End If

                If dr("FechaInicio") = dFechaAnterior Then
                    I = I + 1
                Else
                    I = 0
                End If

                ''''''''''SUSTITUYE A LOS SELECT CASE'''''''''''

                Dim etiquetas() As Windows.Forms.Label = {lbl0, lbl1, lbl2, lbl3, lbl4, lbl5, lbl6, lbl7, lbl8, lbl9, lbl10, lbl11, lbl12, lbl13, lbl14, lbl15, lbl16, lbl17, lbl18, lbl19, lbl20, lbl21, lbl22, lbl23, lbl24, lbl25, lbl26, lbl27, lbl28, lbl29, lbl30}
                Select Day(dr("FechaInicio"))
                    Case 21
                        Select Case I
                            Case 0
                                lbl0.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl0.Text = lbl0.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(0).Visible = True
                                ''InsertarVector(0, dr("FechaInicio)
                        End Select
                    Case 22
                        Select Case I
                            Case 0
                                lbl1.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl1.Text = lbl1.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(1).Visible = True
                                'InsertarVector(1, dr("FechaInicio)
                        End Select
                    Case 23
                        Select Case I
                            Case 0
                                lbl2.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl2.Text = lbl2.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(2).Visible = True
                                'InsertarVector(2, dr("FechaInicio)
                        End Select
                    Case 24
                        Select Case I
                            Case 0
                                lbl3.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl3.Text = lbl3.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(3).Visible = True
                                'InsertarVector(3, dr("FechaInicio)
                        End Select
                    Case 25
                        Select Case I
                            Case 0
                                lbl4.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl4.Text = lbl4.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(4).Visible = True
                                'InsertarVector(4, dr("FechaInicio)
                        End Select
                    Case 26
                        Select Case I
                            Case 0
                                lbl5.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl5.Text = lbl5.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(5).Visible = True
                                'InsertarVector(5, dr("FechaInicio)
                        End Select
                    Case 27
                        Select Case I
                            Case 0
                                lbl6.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl6.Text = lbl6.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(6).Visible = True
                                'InsertarVector(6, dr("FechaInicio)
                        End Select
                    Case 28
                        Select Case I
                            Case 0
                                lbl7.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl7.Text = lbl7.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(7).Visible = True
                                'InsertarVector(7, dr("FechaInicio)
                        End Select
                    Case 29
                        Select Case I
                            Case 0
                                lbl8.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl8.Text = lbl8.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(8).Visible = True
                                'InsertarVector(8, dr("FechaInicio)
                        End Select
                    Case 30
                        Select Case I
                            Case 0
                                lbl9.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl9.Text = lbl9.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(9).Visible = True
                                'InsertarVector(9, dr("FechaInicio)
                        End Select
                    Case 31
                        Select Case I
                            Case 0
                                lbl10.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl10.Text = lbl10.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(10).Visible = True
                                'InsertarVector(10, dr("FechaInicio)
                        End Select
                    Case 1
                        Select Case I
                            Case 0
                                lbl11.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl11.Text = lbl11.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(11).Visible = True
                                'InsertarVector(11, dr("FechaInicio)
                        End Select
                    Case 2
                        Select Case I
                            Case 0
                                lbl12.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl12.Text = lbl12.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(12).Visible = True
                                'InsertarVector(12, dr("FechaInicio)
                        End Select
                    Case 3
                        Select Case I
                            Case 0
                                lbl13.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl13.Text = lbl13.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(13).Visible = True
                                'InsertarVector(13, dr("FechaInicio)

                        End Select
                    Case 4
                        Select Case I
                            Case 0
                                lbl14.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl14.Text = lbl14.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(14).Visible = True
                                'InsertarVector(14, dr("FechaInicio)
                        End Select
                    Case 5
                        Select Case I
                            Case 0
                                lbl15.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl15.Text = lbl15.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(15).Visible = True
                                'InsertarVector(15, dr("FechaInicio)
                        End Select
                    Case 6
                        Select Case I
                            Case 0
                                lbl16.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl16.Text = lbl16.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(16).Visible = True
                                'InsertarVector(16, dr("FechaInicio)
                        End Select
                    Case 7
                        Select Case I
                            Case 0
                                lbl17.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl17.Text = lbl17.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(17).Visible = True
                                'InsertarVector(17, dr("FechaInicio)
                        End Select
                    Case 8
                        Select Case I
                            Case 0
                                lbl18.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl18.Text = lbl18.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(18).Visible = True
                                'InsertarVector(18, dr("FechaInicio)
                        End Select
                    Case 9
                        Select Case I
                            Case 0
                                lbl19.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl19.Text = lbl19.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(19).Visible = True
                                'InsertarVector(19, dr("FechaInicio)
                        End Select
                    Case 10
                        Select Case I
                            Case 0
                                lbl20.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl20.Text = lbl20.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(20).Visible = True
                                'InsertarVector(20, dr("FechaInicio)
                        End Select
                    Case 11
                        Select Case I
                            Case 0
                                lbl21.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl21.Text = lbl21.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(21).Visible = True
                                'InsertarVector(21, dr("FechaInicio)
                        End Select
                    Case 12
                        Select Case I
                            Case 0
                                lbl22.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl22.Text = lbl22.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(22).Visible = True
                                'InsertarVector(22, dr("FechaInicio)
                        End Select
                    Case 13
                        Select Case I
                            Case 0
                                lbl23.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl23.Text = lbl23.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(23).Visible = True
                                'InsertarVector(23, dr("FechaInicio)
                        End Select
                    Case 14
                        Select Case I
                            Case 0
                                lbl24.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl24.Text = lbl24.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(24).Visible = True
                                'InsertarVector(24, dr("FechaInicio)
                        End Select
                    Case 15
                        Select Case I
                            Case 0
                                lbl25.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl25.Text = lbl25.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(25).Visible = True
                                'InsertarVector(25, dr("FechaInicio)
                        End Select
                    Case 16
                        Select Case I
                            Case 0
                                lbl26.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl26.Text = lbl26.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(26).Visible = True
                                'InsertarVector(26, dr("FechaInicio)
                        End Select
                    Case 17
                        Select Case I
                            Case 0
                                lbl27.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl27.Text = lbl27.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(27).Visible = True
                                'InsertarVector(27, dr("FechaInicio)
                        End Select
                    Case 18
                        Select Case I
                            Case 0
                                lbl28.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl28.Text = lbl28.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(28).Visible = True
                                'InsertarVector(28, dr("FechaInicio)
                        End Select
                    Case 19
                        Select Case I
                            Case 0
                                lbl29.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl29.Text = lbl29.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(29).Visible = True
                                'InsertarVector(29, dr("FechaInicio)
                        End Select
                    Case 20
                        Select Case I
                            Case 0
                                lbl30.Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                            Case 1
                                lbl30.Text = lbl30.Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                            Case 2
                                'lblMas(30).Visible = True
                                'InsertarVector(30, dr("FechaInicio)
                        End Select
                End Select

                'If Day(dr("FechaInicio")) = contar Then
                '    If I = 0 Then
                '        etiquetas(contador).Text = "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2) & vbCrLf & sCodigo
                '        Windows.Forms.Application.DoEvents()
                '    Else
                '        etiquetas(contador).Text = etiquetas(contador).Text & vbCrLf & vbCrLf & "Obra: " & dr("Nobra") & vbCrLf & "Horas: " & Round(dr("Horas"), 2)
                '    End If
                '    If contar = 31 Then
                '        contar = 1
                '    Else
                '        contar = contar + 1
                '    End If

                '    contador = contador + 1
                'End If


                '''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Windows.Forms.Application.DoEvents()

                dFechaAnterior = dr("FechaInicio")

            Next
        End If

        lblTotHoras.Text = dblHoras

        Dim toolTip As New Windows.Forms.ToolTip

        toolTip.SetToolTip(lbl0, MostrarEtiqueta(lbl0))
        toolTip.SetToolTip(lbl1, MostrarEtiqueta(lbl1))
        toolTip.SetToolTip(lbl2, MostrarEtiqueta(lbl2))
        toolTip.SetToolTip(lbl3, MostrarEtiqueta(lbl3))
        toolTip.SetToolTip(lbl4, MostrarEtiqueta(lbl4))
        toolTip.SetToolTip(lbl5, MostrarEtiqueta(lbl5))
        toolTip.SetToolTip(lbl6, MostrarEtiqueta(lbl6))
        toolTip.SetToolTip(lbl7, MostrarEtiqueta(lbl7))
        toolTip.SetToolTip(lbl8, MostrarEtiqueta(lbl8))
        toolTip.SetToolTip(lbl9, MostrarEtiqueta(lbl9))
        toolTip.SetToolTip(lbl10, MostrarEtiqueta(lbl10))
        toolTip.SetToolTip(lbl11, MostrarEtiqueta(lbl11))
        toolTip.SetToolTip(lbl12, MostrarEtiqueta(lbl12))
        toolTip.SetToolTip(lbl13, MostrarEtiqueta(lbl13))
        toolTip.SetToolTip(lbl14, MostrarEtiqueta(lbl14))
        toolTip.SetToolTip(lbl15, MostrarEtiqueta(lbl15))
        toolTip.SetToolTip(lbl16, MostrarEtiqueta(lbl16))
        toolTip.SetToolTip(lbl17, MostrarEtiqueta(lbl17))
        toolTip.SetToolTip(lbl18, MostrarEtiqueta(lbl18))
        toolTip.SetToolTip(lbl19, MostrarEtiqueta(lbl19))
        toolTip.SetToolTip(lbl20, MostrarEtiqueta(lbl20))
        toolTip.SetToolTip(lbl21, MostrarEtiqueta(lbl21))
        toolTip.SetToolTip(lbl22, MostrarEtiqueta(lbl22))
        toolTip.SetToolTip(lbl23, MostrarEtiqueta(lbl23))
        toolTip.SetToolTip(lbl24, MostrarEtiqueta(lbl24))
        toolTip.SetToolTip(lbl25, MostrarEtiqueta(lbl25))
        toolTip.SetToolTip(lbl26, MostrarEtiqueta(lbl26))
        toolTip.SetToolTip(lbl27, MostrarEtiqueta(lbl27))
        toolTip.SetToolTip(lbl28, MostrarEtiqueta(lbl28))
        toolTip.SetToolTip(lbl29, MostrarEtiqueta(lbl29))
        toolTip.SetToolTip(lbl30, MostrarEtiqueta(lbl30))

        'Libero memoria
        dt = Nothing
    End Function

    Function MostrarEtiqueta(ByVal e As Control) As String
        Dim tArray() As String
        Dim sTmp As String
        Dim I As Long
        Dim sTexto As String = ""

        If e.Text = "" Then
            Exit Function
        End If

        'Como maximo va a tener 2 obras
        sTmp = e.Text

        tArray = Split(sTmp, vbCrLf)

        For I = LBound(tArray) To UBound(tArray)
            sTexto = sTexto & Reemplazar(tArray(I)) & " "
        Next

        MostrarEtiqueta = sTexto
    End Function

    Function Reemplazar(ByVal sTexto As String) As String
        Dim sObra As String
        Dim fwnOC As New ObraCabecera
        Dim dt As DataTable
        Dim isql As String
        Dim dtnobra As DataTable
        Dim numero As String
        Dim DE As New Engine.BE.DataEngine

        If InStr(1, sTexto, "Obra: ") = 0 Then
            Reemplazar = sTexto
            Exit Function
        End If


        sObra = Mid(sTexto, InStr(1, sTexto, "Obra: ") + 6, Len(sTexto))

        'añadido para que busque el IDObra en funcion del numero de obra

        isql = "select IdObra from tbobracabecera where Nobra = '" & sObra & "'"
        dtnobra = DE.Filter("tbObraCabecera", "IdObra", "Nobra = '" & sObra & "'")

        'rsnobra = AdminData.GetData(isql)

        'If rsnobra.EOF = False Then
        Dim row As DataRow = dtnobra.Rows(dtnobra.Rows.Count - 1)

        numero = row("IdObra")

        ' fin nuevo

        If sObra <> "" Then

            'rs = fwnOC.Filter(, "idObra=" & sObra)  '& "'")
            'dt = fwnOC.Filter("*", "idObra=" & numero)  '& "'")
            Dim filtro As New Filter
            filtro.Add("IdObra", numero)
            dt = fwnOC.Filter(filtro)
            If dt.Rows.Count > 0 Then
                Dim dr As DataRow = dt.Rows(dt.Rows.Count - 1)
                sTexto = "'" & dr("DescObra") & "'"
            End If

            'Libero memoria
            dt = Nothing
            fwnOC = Nothing
        End If
        Reemplazar = sTexto
    End Function

    Function BuscarCodigo(ByVal sObra As String, ByVal sOperario As String, ByVal dFecha As Date) As String
        Dim sSQL As String
        Dim dt As DataTable
        Dim DE As New Expertis.Engine.BE.DataEngine

        sSQL = "IdObra='" & sObra & "' and IdOperario='" & sOperario & "' and Fechainicio='" & dFecha & "'"

        Try
            dt = DE.Filter("tbobramodcontrol", "*", sSQL)
            If dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    BuscarCodigo = dr("idhora")
                Next
            Else
                BuscarCodigo = ""
            End If
        Catch ex As Exception

        End Try
        

        'libero memoria
        dt = Nothing
    End Function

    Function CargarDatosGuardados()
        Dim txtSQL As String
        'Dim rs As New Recordset
        'Dim rs1 As New Recordset
        Dim dt As DataTable
        Dim dtl As DataTable
        Dim DE As New Expertis.Engine.BE.DataEngine
        txtSQL = "idoperario='" & advOperario1.Value & "' and Mes=" & cbxMes.Value & _
                 " and anyo=" & cbxAnyo.Value
        dt = DE.Filter("tbOperarioCalendario", "*", txtSQL)
        dtl = DE.Filter("tbHistoricoCalendario", "*", txtSQL)

        'If rs.EOF = False Then
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(dt.Rows.Count - 1)
            txtNMensual.Text = Round(dr("NMensual"), 2)
            txtNAnticipo.Text = Round(dr("NAnticipo"), 2)
            txtSSocial.Text = Round(dr("SSocial"), 2)
            txtACC.Text = Round(dr("ACC"), 2)
            txtACCImporte.Text = Round(dr("ACCImporte"), 2)
            txtACCTotal.Text = Round(dr("ACCTotal"), 2)
            txtCC.Text = Round(dr("CC"), 2)
            txtCCImporte.Text = Round(dr("CCImporte"), 2)
            txtCCTotal.Text = Round(dr("CCTotal"), 2)
            txtVACAV.Text = Round(dr("Vacav"), 2)
            txtVACAVImporte.Text = Round(dr("VacavImporte"), 2)
            txtVACAVTotal.Text = Round(dr("VacavTotal"), 2)
            txtAnticipo.Text = Round(dr("Anticipo"), 2)
            txtVarios2.Text = Round(dr("Varios2"), 2)
            clbACCDesde.Value = dr("ACCDesde")
            clbACCHasta.Value = dr("ACCHasta")
            clbCCDesde.Value = dr("CCDesde")
            clbCCHasta.Value = dr("CCHasta")
            clbVACAVDesde.Value = dr("VACAVDesde")
            clbVACAVHasta.Value = dr("VACAVHasta")
            txtTotal.Text = Nz(Round(dr("Total"), 2), 0)
            txtDerechos.Text = Nz(Round(dr("Derechos"), 2), 0)
            txtIRPF.Text = Nz(Round(dr("IRPF"), 2), 0)
            txtPrima.Text = Nz(Round(dr("Prima"), 2), 0)

            If dr("Atrasos") Is DBNull.Value = True Then
                txtAtrasos.Text = Nz(dr("Atrasos"), 0)
            Else
                txtAtrasos.Text = Nz(Round(dr("Atrasos"), 2), 0)
            End If
            If dr("Pagare") Is DBNull.Value = True Then
                txtPagare.Text = Nz(dr("Pagare"), 0)
            Else
                txtPagare.Text = Nz(Round(dr("Pagare"), 2), 0)
            End If

            If dr("PagaExtra") Is DBNull.Value = True Then
                txtPagaExtra.Text = Nz(dr("PagaExtra"), 0)
            Else
                txtPagaExtra.Text = Nz(Round(dr("PagaExtra"), 2), 0)
            End If

            'If dr("HorasReg") Is DBNull.Value = False Then
            '    txtHorasReg.Text = Nz(Round(dr("HorasReg"), 2), 0)
            'End If

            If dr("ImpReg") Is DBNull.Value = False Then
                txtImpReg.Text = Nz(Round(dr("ImpReg"), 2), 0)
            End If
            'If dr("TotalReg") Is DBNull.Value = False Then
            '    txtTotalReg.Text = Nz(Round(dr("TotalReg"), 2), 0)
            'End If
            If dr("Comentario") <> "" Then
                lblComentarios.BackColor = System.Drawing.Color.Yellow
            Else
                lblComentarios.BackColor = System.Drawing.Color.Gray
            End If
            'If rs1.Fields("id").Value <> "" Then
            'lblHistoricoCalendario.BackColor = System.Drawing.Color.Yellow
            ' Else
            ' lblHistoricoCalendario.BackColor = System.Drawing.Color.Gray
            ' End If
        End If
        'Libero memoria
        dt = Nothing
    End Function

    Function LimpiarFestivos()
        lb2.ForeColor = System.Drawing.Color.Black
        lb3.ForeColor = System.Drawing.Color.Black
        lb4.ForeColor = System.Drawing.Color.Black
        lb5.ForeColor = System.Drawing.Color.Black
        lb6.ForeColor = System.Drawing.Color.Black
        lb7.ForeColor = System.Drawing.Color.Black
        lb8.ForeColor = System.Drawing.Color.Black
        lb9.ForeColor = System.Drawing.Color.Black
        lb10.ForeColor = System.Drawing.Color.Black
        lb11.ForeColor = System.Drawing.Color.Black
        lb12.ForeColor = System.Drawing.Color.Black
        lb13.ForeColor = System.Drawing.Color.Black
        lb14.ForeColor = System.Drawing.Color.Black
        lb15.ForeColor = System.Drawing.Color.Black
        lb16.ForeColor = System.Drawing.Color.Black
        lb17.ForeColor = System.Drawing.Color.Black
        lb18.ForeColor = System.Drawing.Color.Black
        lb19.ForeColor = System.Drawing.Color.Black
        lb20.ForeColor = System.Drawing.Color.Black
        lb21.ForeColor = System.Drawing.Color.Black
        lb22.ForeColor = System.Drawing.Color.Black
        lb23.ForeColor = System.Drawing.Color.Black
        lb24.ForeColor = System.Drawing.Color.Black
        lb25.ForeColor = System.Drawing.Color.Black
        lb26.ForeColor = System.Drawing.Color.Black
        lb27.ForeColor = System.Drawing.Color.Black
        lb28.ForeColor = System.Drawing.Color.Black
        lb29.ForeColor = System.Drawing.Color.Black
        lb30.ForeColor = System.Drawing.Color.Black
        lb31.ForeColor = System.Drawing.Color.Black
        lb31.ForeColor = System.Drawing.Color.Black
    End Function

    Function DiasFestivos()
        Dim f As New Filter
        'Dim rs As Recordset
        Dim dt As DataTable
        Dim DE As New General.CalendarioCentro
        '++++++++++++++++++++++++++++++++++++++++++++++++++++
        'CODIGO ORIGINAL
        'Dim fECHAfINAL As Date
        'fECHAfINAL = "01/" & cbxmes.Value & "/" & cbxAnyo.Value
        'Obtenemos el último dia del mes
        'fECHAfINAL = DateAdd(DateInterval.Month, 1, fECHAfINAL)
        'fECHAfINAL = DateAdd(DateInterval.Day, -1, fECHAfINAL)

        'sSQL = "(Fecha >='01/" & cbxmes.Value & "/" & cbxAnyo.Value & _
        '      "'  and Fecha <='" & Day(fECHAfINAL) & "/" & cbxmes.Value & "/" & cbxAnyo.Value & "')" & _
        '     " and idcentro='" & sObra & "' and TipoDia = 1"
        '+++++++++++++++++++++++++++++++++++++++++++++++

        f.Add("Fecha", FilterOperator.GreaterThanOrEqual, "21/" & MesAnterior(cbxMes.Value, cbxAnyo.Value))
        f.Add("Fecha", FilterOperator.LessThanOrEqual, "20/" & cbxMes.Value & "/" & cbxAnyo.Value)

        'David 12/11

        Dim drOperario As DataRow = advOperario1.SelectedRow()
        Dim sObra = Nz(drOperario("Obra_Predeterminada"), "")
        'Dim op As New Operario
        'sObra = op.devuelveIDTrabajo(sObra)

        'David 12/11
        f.Add("IDCentro", FilterOperator.Equal, sObra)
        f.Add("TipoDia", FilterOperator.Equal, 1)
        'sSQL = "(Fecha >='21/" & MesAnterior(cbxMes.Value, cbxAnyo.Value) & _
        '       "' and Fecha <='20/" & cbxMes.Value & "/" & cbxAnyo.Value & "')" & _
        '       " and idcentro='" & sObra & "' and TipoDia = 1"

        'rs = AdminData.GetData("tbcalendariocentro", , sSQL)
        'rs = AdminData.Filter("tbCalendarioCentro", , sSQL)
        dt = DE.Filter(f)
        'If rs.EOF = True Then
        If dt.Rows.Count = 0 Then
            Exit Function
        Else
            For Each dr As DataRow In dt.Rows

                Select Case Day(dr("Fecha"))
                    Case 21
                        lb1.ForeColor = System.Drawing.Color.Red
                    Case 22
                        lb2.ForeColor = System.Drawing.Color.Red
                    Case 23
                        lb3.ForeColor = System.Drawing.Color.Red
                    Case 24
                        lb4.ForeColor = System.Drawing.Color.Red
                    Case 25
                        lb5.ForeColor = System.Drawing.Color.Red
                    Case 26
                        lb6.ForeColor = System.Drawing.Color.Red
                    Case 27
                        lb7.ForeColor = System.Drawing.Color.Red
                    Case 28
                        lb8.ForeColor = System.Drawing.Color.Red
                    Case 29
                        lb9.ForeColor = System.Drawing.Color.Red
                    Case 30
                        lb10.ForeColor = System.Drawing.Color.Red
                    Case 31
                        lb11.ForeColor = System.Drawing.Color.Red
                    Case 1
                        lb12.ForeColor = System.Drawing.Color.Red
                    Case 2
                        lb13.ForeColor = System.Drawing.Color.Red
                    Case 3
                        lb14.ForeColor = System.Drawing.Color.Red
                    Case 4
                        lb15.ForeColor = System.Drawing.Color.Red
                    Case 5
                        lb16.ForeColor = System.Drawing.Color.Red
                    Case 6
                        lb17.ForeColor = System.Drawing.Color.Red
                    Case 7
                        lb18.ForeColor = System.Drawing.Color.Red
                    Case 8
                        lb19.ForeColor = System.Drawing.Color.Red
                    Case 9
                        lb20.ForeColor = System.Drawing.Color.Red
                    Case 10
                        lb21.ForeColor = System.Drawing.Color.Red
                    Case 11
                        lb22.ForeColor = System.Drawing.Color.Red
                    Case 12
                        lb23.ForeColor = System.Drawing.Color.Red
                    Case 13
                        lb24.ForeColor = System.Drawing.Color.Red
                    Case 14
                        lb25.ForeColor = System.Drawing.Color.Red
                    Case 15
                        lb26.ForeColor = System.Drawing.Color.Red
                    Case 16
                        lb27.ForeColor = System.Drawing.Color.Red
                    Case 17
                        lb28.ForeColor = System.Drawing.Color.Red
                    Case 18
                        lb29.ForeColor = System.Drawing.Color.Red
                    Case 19
                        lb30.ForeColor = System.Drawing.Color.Red
                    Case 20
                        lb31.ForeColor = System.Drawing.Color.Red
                End Select
                'Avanzo registro
            Next
        End If



        dt = Nothing
    End Function

    Function bValidarWhere() As String

        Dim bValidar As Integer

        Dim fil As New Filter
        If advOperario1.Text = "" Then
            bValidar = 0
            Exit Function
        End If
        If cbxMes.Text = "" Then
            bValidar = 0
            Exit Function
        End If
        If cbxAnyo.Text = "" Then
            bValidar = 0
            Exit Function
        End If

        Dim sSQL As String
        'Dim rs As New Recordset
        Dim dt As DataTable
        Dim DE As New Expertis.Engine.BE.DataEngine

        fil.Add("IDOperario", FilterOperator.Equal, advOperario1.Value)
        fil.Add("Mes", FilterOperator.Equal, cbxMes.Value)
        fil.Add("Anyo", FilterOperator.Equal, cbxAnyo.Value)

        sSQL = "idoperario ='" & advOperario1.Value & "' and" & _
               " mes = " & cbxMes.Value & _
               " and anyo = " & cbxAnyo.Value
        'rs = AdminData.GetData("tbOperarioCalendario", , sSQL)

        Return sSQL
    End Function

    Function bValidar() As Integer
        Dim fil As New Filter
        If advOperario1.Text = "" Then
            bValidar = 0
            Exit Function
        End If
        If cbxMes.Text = "" Then
            bValidar = 0
            Exit Function
        End If
        If cbxAnyo.Text = "" Then
            bValidar = 0
            Exit Function
        End If

        Dim sSQL As String
        'Dim rs As New Recordset
        Dim dt As DataTable
        Dim DE As New Expertis.Engine.BE.DataEngine

        fil.Add("IDOperario", FilterOperator.Equal, advOperario1.Value)
        fil.Add("Mes", FilterOperator.Equal, cbxMes.Value)
        fil.Add("Anyo", FilterOperator.Equal, cbxAnyo.Value)

        sSQL = "idoperario ='" & advOperario1.Value & "' and" & _
               " mes = " & cbxMes.Value & _
               " and anyo = " & cbxAnyo.Value
        'rs = AdminData.GetData("tbOperarioCalendario", , sSQL)

        Dim op As New OperarioCalendario
        Dim formObservaciones As New frmObservaciones
        dt = op.Filter(fil)

        'If rs.EOF = True Then
        If dt.Rows.Count > 0 Then
            bValidar = 2
        Else
            formObservaciones.filtroHistorico = fil
            bValidar = 1
        End If

        'libero memoria
        dt = Nothing
    End Function

    Function Guardar()
        'Dim SQLHistorico As String
        Dim txtSQL As String
        Dim dt As DataTable
        'Dim rs As New Recordset
        Dim sTexto As String = ""
        Dim DE As New Expertis.Engine.BE.DataEngine

        If CamposObligatorios() = False Then
            Exit Function
        End If

        If MsgBox("¿Desea realmente guardar los datos?", vbInformation + vbYesNo, "Guardar Datos") = vbNo Then
            Exit Function
        End If

        Dim FilHist As New Filter
        FilHist.Add("IDOPerario", FilterOperator.Equal, advOperario1.Value)
        FilHist.Add("Mes", FilterOperator.Equal, cbxMes.Value)
        FilHist.Add("anyo", FilterOperator.Equal, cbxAnyo.Value)

        txtSQL = "idoperario='" & advOperario1.Value & "' and Mes=" & cbxMes.Value & _
                 " and anyo=" & cbxAnyo.Value & ""
        dt = DE.Filter("tbOperarioCalendario", FilHist)

        Dim clsOP As New OperarioCalendario


        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(dt.Rows.Count - 1)
            If MsgBox("Ya hay datos guardados para este operario. ¿Desea sobreescribirlos?", vbQuestion + vbYesNo, "Sobreescribir") = vbNo Then
                Exit Function
            Else
                sTexto = Nz(dr("Comentario"), "")
                clsOP.BorrarDatos(txtSQL)
                'txtSQL = "delete from tbOperarioCalendario where " & txtSQL
                'AdminData.Execute(txtSQL)
            End If
        End If

        If txtFechaBaja.Text = "" Then
            txtFechaBaja.Text = "01/01/1900"
        End If

        clsOP.InsertarDatos(advOperario1.Value, cbxMes.Value, cbxAnyo.Value, txtNMensual.Text, txtNAnticipo.Text, txtSSocial.Text, txtACC.Text, txtACCImporte.Text, _
        txtACCTotal.Text, txtCC.Text, txtCCImporte.Text, txtCCTotal.Text, txtVACAV.Text, txtVACAVImporte.Text, txtVACAVTotal.Text, txtAnticipo.Text, txtVarios2.Text, clbACCDesde.Value, _
        clbACCHasta.Value, clbCCDesde.Value, clbCCHasta.Value, clbVACAVDesde.Value, clbVACAVHasta.Value, txtTotal.Text, txtDerechos.Text, txtIRPF.Text, txtHReg.Text, _
        sTexto, sObra, txtHorasReg.Text, txtImpReg.Text, txtTotalReg.Text, ntbDiasLab1.Text, ntbDiasTrab1.Text, ntbDiasTot1.Text, ntbDiasLab2.Text, ntbDiasTrab2.Text, _
        ntbDiasTot2.Text, ntbHorasTeoricas.Text, ntbDiasTot.Text, txtHorasRegC.Text, txtNHNormal.Text, txtNHNormalImporte.Text, txtNHNormalTotal.Text, _
        txtNHExtra.Text, txtNHExtraImporte.Text, txtNHExtraTotal.Text, txtNHEspecial.Text, txtNHEspecialImporte.Text, txtNHEspecialTotal.Text, txtFechaBaja.Text, txtSueldo.Text, _
        txtPlus.Text, txtDieta.Text, txtVarios.Text, ntbAmortizacion.Text, txtEmbargos.Text, txtAnticipoFijo.Text, txtAtrasos.Text, txtPagare.Text, txtPagaExtra.Text, _
        txtPrima.Text)


        'txtSQL = "insert into tbOperarioCalendario(IdOperario, Mes, Anyo, NMensual, NAnticipo, SSocial, ACC, ACCImporte, ACCTotal, CC, CCImporte, CCTotal, Vacav, VacavImporte, " & _
        '        "VacavTotal , Anticipo, Varios2, ACCDesde, ACCHasta, CCDesde, CCHasta, VACAVDesde, VACAVHasta, Total, Derechos, IRPF, HoraReg, Comentario, IdObraPredet, HorasReg, ImpReg, TotalReg, " & _
        '        "HorasLab0120, HorasTrab0120, Importe0120, HorasLab2131, HorasTrab2131, Importe2131, HorasTeoricas, IT, HorasRegularizadasC, " & _
        '        "HNCantidad, HNPrecio, HNImporte, HXCantidad, HXPrecio, HXImporte, HECantidad, HEPrecio, HEImporte, FBaja, Sueldo, Plus, Dieta, Varios, Amortizacion, Embargos, AnticipoFijo, Atrasos, Pagare, PagaExtra,Prima)" & _
        '         "Values ('" & advOperario1.Value & "', " & cbxMes.Value & ", " & cbxAnyo.Value & ", " & _
        '         Replace(Nz(txtNMensual.Text, 0), ",", ".") & ", " & Replace(Nz(txtNAnticipo.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtSSocial.Text, 0), ",", ".") & ", " & Replace(Nz(txtACC.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtACCImporte.Text, 0), ",", ".") & ", " & Replace(Nz(txtACCTotal.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtCC.Text, 0), ",", ".") & ", " & Replace(Nz(txtCCImporte.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtCCTotal.Text, 0), ",", ".") & ", " & Replace(Nz(txtVACAV.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtVACAVImporte.Text, 0), ",", ".") & ", " & Replace(Nz(txtVACAVTotal.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtAnticipo.Text, 0), ",", ".") & ", " & Replace(Nz(txtVarios2.Text, 0), ",", ".") & _
        '         ", '" & clbACCDesde.Value & "', '" & clbACCHasta.Value & "', '" & clbCCDesde.Value & "', '" & clbCCHasta.Value & _
        '         "', '" & clbVACAVDesde.Value & "', '" & clbVACAVHasta.Value & "', " & _
        '         Replace(Nz(txtTotal.Text, 0), ",", ".") & ", " & Replace(Nz(txtDerechos.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtIRPF.Text, 0), ",", ".") & ", " & Replace(Nz(txtHReg.Text, 0), ",", ".") & ", '" & _
        '         sTexto & "','" & sObra & "', " & _
        '         Replace(Nz(txtHorasReg.Text, 0), ",", ".") & ", " & Replace(Nz(txtImpReg.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtTotalReg.Text, 0), ",", ".") & ", " & Replace(Nz(ntbDiasLab1.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(ntbDiasTrab1.Text, 0), ",", ".") & ", " & Replace(Nz(ntbDiasTot1.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(ntbDiasLab2.Text, 0), ",", ".") & ", " & Replace(Nz(ntbDiasTrab2.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(ntbDiasTot2.Text, 0), ",", ".") & ", " & Replace(Nz(ntbHorasTeoricas.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(ntbDiasTot.Text, 0), ",", ".") & ", " & Replace(Nz(txtHorasRegC.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtNHNormal.Text, 0), ",", ".") & ", " & Replace(Nz(txtNHNormalImporte.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtNHNormalTotal.Text, 0), ",", ".") & ", " & Replace(Nz(txtNHExtra.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtNHExtraImporte.Text, 0), ",", ".") & ", " & Replace(Nz(txtNHExtraTotal.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtNHEspecial.Text, 0), ",", ".") & ", " & Replace(Nz(txtNHEspecialImporte.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtNHEspecialTotal.Text, 0), ",", ".") & ", '" & Nz(txtFechaBaja.Text, "01/01/1900") & "', " & _
        '         Replace(Nz(txtSueldo.Text, 0), ",", ".") & ", " & Replace(Nz(txtPlus.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtDieta.Text, 0), ",", ".") & ", " & Replace(Nz(txtVarios.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(ntbAmortizacion.Text, 0), ",", ".") & ", " & Replace(Nz(txtEmbargos.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtAnticipoFijo.Text, 0), ",", ".") & ", " & Replace(Nz(txtAtrasos.Text, 0), ",", ".") & ", " & _
        '         Replace(Nz(txtPagare.Text, 0), ",", ".") & ", " & Replace(Nz(txtPagaExtra.Text, 0), ",", ".") & ", " & Replace(Nz(txtPrima.Text, 0), ",", ".") & ")"

        'AdminData.Execute(txtSQL)
        '


        'SQLHistorico = "select * from tbHistoricoCalendario where " & _
        '"idoperario='" & advOperario1.Value & "' and Mes=" & cbxMes.Value & _
        '    " and anyo=" & cbxAnyo.Value
        Dim clsBE As New BE.DataEngine
       

        Dim dtHist As DataTable = clsBE.Filter("tbHistoricoCalendario", FilHist)

        'If rs.EOF = False Then
        Dim fwPOCI As New HistoricoCalendario
        Dim dtPer As New DataTable

        'Dim rsPer As New Recordset
        Dim f As New Filter
        'Dim i As Integer

        'While rs.EOF = False
        f.Clear()
        f.Add("IdOperario", advOperario1.Value)
        f.Add("Mes", cbxMes.Value)
        f.Add("Anyo", cbxAnyo.Value)

        'dtPer = fwPOCI.Filter(f)
        'Dim drPer As DataRow

        'dtPer.Columns.Add("id")
        'dtPer.Columns.Add("IDOperario")
        'dtPer.Columns.Add("Mes")
        'dtPer.Columns.Add("Anyo")
        'dtPer.Columns.Add("Sueldo")
        'dtPer.Columns.Add("C_H_N")
        'dtPer.Columns.Add("C_H_X")
        'dtPer.Columns.Add("C_H_E")
        'dtPer.Columns.Add("Plus")
        'dtPer.Columns.Add("Dieta")
        'dtPer.Columns.Add("Varios")
        'dtPer.Columns.Add("Embargo")
        'dtPer.Columns.Add("Anticipo_fijo")
        'dtPer.Columns.Add("Varios_izq")
        'dtPer.Columns.Add("NMensual")
        'dtPer.Columns.Add("NGastos")
        'dtPer.Columns.Add("NMesLiq")
        'dtPer.Columns.Add("NFiniquito")
        'dtPer.Columns.Add("Atrasos")
        'dtPer.Columns.Add("Pagare")
        'dtPer.Columns.Add("PagaExtra")

        Dim iSQL As String = "INSERT INTO tbHistoricoCalendario(IDOperario, Mes, Anyo, Incentivos, Sueldo, C_H_N, C_H_X, C_H_E, Plus, Dieta, Varios, Embargo, Anticipo_fijo, Varios_izq, NMensual, NGastos, NMesLiq, NFiniquito, Atrasos, Pagare, PagaExtra)" & _
        "VALUES ('" & advOperario1.Value & "','" & cbxMes.Value & "', '" & cbxAnyo.Value & "','" & chxIncentivo.Checked & "','" & txtSueldo.Value & "','" & txtNHNormalImporte.Value & "'," & _
        "'" & txtNHExtraImporte.Value & "','" & txtNHEspecialImporte.Value & "','" & txtPlus.Value & "','" & txtDieta.Value & "','" & txtVarios2.Value & "'," & _
        "'" & txtEmbargos.Value & "','" & txtAnticipoFijo.Value & "','" & txtVarios.Value & "','" & txtNMensual.Value & "','" & txtNAnticipo.Value & "'," & _
        "'" & txtSSocial.Value & "','" & txtDerechos.Value & "','" & txtAtrasos.Value & "','" & txtPagare.Value & "','" & txtPagaExtra.Value & "')"

        clsOP.Ejecutar(iSQL)

        'i = fwPOCI.Autonumerico()
        ''i = rsPer.Fields("id").Value
        ''dtPer = fwPOCI.AddNew()
        'drPer = dtPer.NewRow

        'drPer("id") = i
        'drPer("idOperario") = advOperario1.Value
        'drPer("Mes") = cbxMes.Value
        'drPer("Anyo") = cbxAnyo.Value
        ''rsPer.Fields("Incentivos").Value = chkIncentivo.Text
        'drPer("Sueldo") = txtSueldo.Value
        'drPer("C_H_N") = txtNHNormalImporte.Value
        'drPer("C_H_X") = txtNHExtraImporte.Value
        'drPer("C_H_E") = txtNHEspecialImporte.Value
        'drPer("Plus") = txtPlus.Value
        'drPer("Dieta") = txtDieta.Value
        'drPer("Varios") = txtVarios2.Value
        'drPer("Embargo") = txtEmbargos.Value
        'drPer("Anticipo_Fijo") = txtAnticipoFijo.Value
        'drPer("Varios_izq") = txtVarios.Value
        'drPer("NMensual") = txtNMensual.Value
        'drPer("NGastos") = txtNAnticipo.Value
        'drPer("NMesLiq") = txtSSocial.Value
        'drPer("NFiniquito") = txtDerechos.Value
        'drPer("Atrasos") = txtAtrasos.Value
        'drPer("Pagare") = txtPagare.Value
        'drPer("PagaExtra") = txtPagaExtra.Value
        'dtPer.Rows.Add(drPer)
        ''rsPer.Update()
        'fwPOCI.Update(dtPer)

        'rs.MoveNext()
        'End While
        'fwPOCI = Nothing

        'Libero memoria


        'alejandro AKI ES DONDE VOY A PONER LO DEL CALCULO DE HORAS POR OBRA E INTRODUCCIÓN DE ESOS DATOS EN LAS TABLAS DE C.I.

        'Dim fECHAfINAL As Date
        'fECHAfINAL = "01/" & cbxmes.Value & "/" & cbxAnyo.Value
        'Obtenemos el último dia del mes
        'fECHAfINAL = DateAdd(DateInterval.Month, 1, fECHAfINAL)
        'fECHAfINAL = DateAdd(DateInterval.Day, -1, fECHAfINAL)

        'txtSQL = "select idobra, idhora, idcategoria, sum(horasrealmod) as totalHoras from vHorasOperarioObra where " & _
        '"idoperario = '" & advIdOperario.Value & "' and fechainicio >= '01/" & cbxmes.Value & "/" & cbxAnyo.Value & _
        '"' and fechainicio <= '" & Day(fECHAfINAL) & "/" & cbxmes.Value & "/" & cbxAnyo.Value & "'" & "group by idobra, idhora, idcategoria"

        'rs = AdminData.GetData(txtSQL, False)
        'If rs.EOF = False Then
        'Dim fwPOCI As New PersonalIndiOperario
        'Dim rsPer As New Recordset
        'Dim f As New Filter
        'While rs.EOF = False
        'f.Clear()
        'f.Add("IdObra", rs.Fields("idObra").Value)
        'f.Add("IdOperario", advIdOperario.Value)
        'f.Add("Mes", cbxmes.Value)
        'f.Add("Anio", cbxAnyo.Value)
        'rsPer = fwPOCI.Filter(f)
        'compruebo si existen ya datos de este operario 
        'If rsPer.EOF = False Then
        'si es asi solo actualizo sus datos
        'rsPer.Fields("Importe").Value = CalcularImporte(rs.Fields("totalHoras").Value, rs.Fields("idHora").Value)
        'rsPer.Update()
        'fwPOCI.Update(rsPer)
        'Else
        'si no lo creo
        '   rsPer = fwPOCI.AddNew()
        '   rsPer.AddNew()
        '  rsPer.Fields("idPersonalInd").Value = AdminData.GetAutoNumeric
        ' rsPer.Fields("idOperario").Value = advIdOperario.Value
        ' rsPer.Fields("idCategoria").Value = rs.Fields("idCategoria").Value
        'rsPer.Fields("idObra").Value = rs.Fields("idObra").Value
        'rsPer.Fields("idHora").Value = rs.Fields("idHora").Value
        'rsPer.Fields("Mes").Value = cbxmes.Value
        'rsPer.Fields("Anio").Value = cbxAnyo.Value
        'rsPer.Fields("Importe").Value = CalcularImporte(rs.Fields("totalHoras").Value, rs.Fields("idHora").Value)
        'rsPer.Update()
        'fwPOCI.Update(rsPer)
        'End If

        'rs.MoveNext()
        'End While
        'fwPOCI = Nothing
        'End If
        'rs = Nothing

    End Function


    Private Function CalcularImporte(ByVal dNHoras As Double, ByVal sIdHora As String) As Double
        'EN ESTA FUNCION VOY A CALCULAR LO Q HA COSTADO ESTE TIO A LA OBRA.
        'DEPENDIENDO DEL TIPO DE HORA SE CALCULA SU PARTE PROPORCIONAL (SI
        'ES NORMAL) O SE MULTIPLICA EL TIEMPO POR SU COSTE HORA
        CalcularImporte = 0
        Dim dPHora As Double
        dPHora = 0
        Select Case sIdHora
            Case "HO"
                'primero averiguo cual es el precio de hora normal dividiendo el total de sueldo
                'entre el total de horas trabajadas en todas las obras

                dPHora = Nz(txtTotalSueldo.Text, 0) / Nz(lblTotHoras.Text, 1)
                If dPHora > 0 Then
                    dPHora = dPHora * dNHoras
                End If

            Case "HX"
                dPHora = Nz(txtNHExtraImporte.Text, 0) * dNHoras
            Case "HE"
                dPHora = Nz(txtNHEspecialImporte.Text, 0) * dNHoras
        End Select

        CalcularImporte = dPHora
    End Function

    Function CamposObligatorios() As Boolean
        CamposObligatorios = False

        If clbACCDesde.Value Is DBNull.Value Then
            MsgBox("Debes rellenar ACCDesde", vbExclamation + vbOKOnly)
            clbACCDesde.Focus()
            Exit Function
        End If

        If clbACCHasta.Value Is DBNull.Value Then
            MsgBox("Debes rellenar ACCHasta", vbExclamation + vbOKOnly)
            clbACCHasta.Focus()
            Exit Function
        End If

        If clbCCDesde.Value Is DBNull.Value Then
            MsgBox("Debes rellenar CCDesde", vbExclamation + vbOKOnly)
            clbCCDesde.Focus()
            Exit Function
        End If

        If clbCCHasta.Value Is DBNull.Value Then
            MsgBox("Debes rellenar CCHasta", vbExclamation + vbOKOnly)
            clbACCHasta.Focus()
            Exit Function
        End If

        If clbVACAVDesde.Value Is DBNull.Value Then
            MsgBox("Debes rellenar VACAVDesde", vbExclamation + vbOKOnly)
            clbVACAVDesde.Focus()
            Exit Function
        End If

        If clbVACAVHasta.Value Is DBNull.Value Then
            MsgBox("Debes rellenar VACAVHasta", vbExclamation + vbOKOnly)
            clbVACAVHasta.Focus()
            Exit Function
        End If

        If txtACC.Text = "" Then
            MsgBox("Debes rellenar ACC", vbExclamation + vbOKOnly)
            txtACC.Focus()
            Exit Function
        End If
        If txtACCImporte.Text = "" Then
            MsgBox("Debes rellenar ACC Importe", vbExclamation + vbOKOnly)
            txtACCImporte.Focus()
            Exit Function
        End If

        If txtCC.Text = "" Then
            MsgBox("Debes rellenar CC", vbExclamation + vbOKOnly)
            txtCC.Focus()
            Exit Function
        End If

        If txtCCImporte.Text = "" Then
            MsgBox("Debes rellenar CC", vbExclamation + vbOKOnly)
            txtCCImporte.Focus()
            Exit Function
        End If

        If txtVACAV.Text = "" Then
            MsgBox("Debes VACAV", vbExclamation + vbOKOnly)
            txtVACAV.Focus()
            Exit Function
        End If

        If txtVACAVImporte.Text = "" Then
            MsgBox("Debes VACAV", vbExclamation + vbOKOnly)
            txtVACAVImporte.Focus()
            Exit Function
        End If

        If txtAnticipo.Text = "" Then
            MsgBox("Debes rellenar Anticipo", vbExclamation + vbOKOnly)
            txtAnticipo.Focus()
            Exit Function
        End If

        If txtVarios2.Text = "" Then
            MsgBox("Debes rellenar Varios", vbExclamation + vbOKOnly)
            txtVarios2.Focus()
            Exit Function
        End If
        If txtAtrasos.Text = "" Then
            MsgBox("Debes rellenar Atrasos", vbExclamation + vbOKOnly)
            txtAtrasos.Focus()
            Exit Function
        End If
        If txtPagare.Text = "" Then
            MsgBox("Debes rellenar Pagare", vbExclamation + vbOKOnly)
            txtPagare.Focus()
            Exit Function
        End If
        If txtPagaExtra.Text = "" Then
            MsgBox("Debes rellenar la Paga Extra", vbExclamation + vbOKOnly)
            txtPagare.Focus()
            Exit Function
        End If
        If txtPrima.Text = "" Then
            MsgBox("Debes rellenar la Prima", vbExclamation + vbOKOnly)
            txtPrima.Focus()
            Exit Function
        End If
        CamposObligatorios = True
    End Function
    'Sergio Blanco - Tecozam 12/04/2016
    Private Sub aviso9m(ByVal idOperario As String, ByVal mes As Integer, ByVal anio As Integer)
        Dim dt As New DataTable
        Dim DE As New Expertis.Engine.BE.DataEngine
        Dim fechaAlta As Date
        Dim aviso As Boolean
        Dim strSQl As String = "select idoperario, fechaAlta,aviso9m from vMaestroOperarioCompleta where idOperario ='" & idOperario & "'"
        Dim faviso As New Filter
        faviso.Add("IDOperario", FilterOperator.Equal, idOperario)
        dt = DE.Filter("vMaestroOperarioCompleta", faviso, "idOperario, fechaAlta, aviso9m")

        'dt = AdminData.GetData(strSQl, False)
        Dim fechatemp As Date
        If mes = 12 Then
            fechatemp = "1/" & CStr(1) & "/" & CStr(anio + 1)
        Else
            fechatemp = "1/" & CStr(mes + 1) & "/" & CStr(anio)
        End If

        'fechatemp = "1/" & CStr(mes + 1) & "/" & CStr(anio)
        Dim ffinMes As Date = DateAdd(DateInterval.Day, -1, fechatemp)
        'MsgBox(ffinMes)
        For Each dr As DataRow In dt.Rows
            fechaAlta = dr("FechaAlta")
            aviso = Nz(dr("aviso9m"), False)

            'MsgBox(fechaAlta)
        Next
        'calculamos la perioricidad del aviso en este caso 10 meses
        If aviso Then

            Dim fechaAviso As Date = DateAdd(DateInterval.Month, 10, fechaAlta)
            'MsgBox(fechaAviso)
            If Month(fechaAviso) = mes And Year(fechaAviso) = anio Then
                MsgBox("El trabajador " & idOperario & " ha cumplido el ciclo de 9 meses")
            Else
                'MsgBox(DateDiff(DateInterval.Month, fechaAlta, ffinMes) & " meses")
                Dim diffMes As Integer = DateDiff(DateInterval.Month, fechaAlta, ffinMes)
                Dim Resto = (DateDiff(DateInterval.Month, fechaAlta, ffinMes) Mod 10)
                'MsgBox(Resto)
                If Resto = 0 And diffMes > 0 Then
                    MsgBox("El trabajador " & idOperario & " ha cumplido el ciclo de 9 meses")
                End If
            End If
        End If

    End Sub

    Private Sub btnCalcular_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalcular.Click
        ' If CamposObligatorios = True Then
        'txtTotal.Text = Round((CDbl(txtDiasTot1.Text) + CDbl(txtDiasTot2.Text) + txtPlus.Text + txtDieta.Text + txtVarios.Text + txtNHNormalTotal.Text + txtNHExtraTotal.Text + txtNHEspecialTotal.Text + txtACCTotal.Text + txtCCTotal.Text + txtVacacTotal.Text) - (txtAnticipo.Text + txtAmortizacion.Text + txtEmbargos.Text + txtVarios2.Text + txtAnticipoFijo.Text), 2)
        'txtTotal.Text = Round(CDbl(Nz(txtTotDias.Text, 0)) + CDbl(Nz(txtTotalReg.Text, 0)) + CDbl(Nz(txtNHNormalTotal.Text, 0)) + _
        'CDbl(Nz(txtNHExtraTotal.Text, 0)) + CDbl(Nz(txtNHEspecialTotal.Text, 0)) + CDbl(Nz(txtPlus.Text, 0)) + _
        'CDbl(Nz(txtDieta.Text, 0)) + CDbl(Nz(txtVarios.Text, 0)) + CDbl(Nz(txtACCTotal.Text, 0)) + CDbl(Nz(txtCCTotal.Text, 0)) + _
        'CDbl(Nz(txtVacacTotal.Text, 0)) + CDbl(Nz(txtAnticipo.Text, 0)) + CDbl(Nz(txtAmortizacion.Text, 0)) + _
        'CDbl(Nz(txtEmbargos.Text, 0)) + CDbl(Nz(txtVarios2.Text, 0)) + CDbl(Nz(txtAnticipoFijo.Text, 0)), 2)
        'End If
        If txtHorasReg.Text.Length <= 0 Then txtHorasReg.Text = 0
        If CDbl(txtHorasReg.Text) < 0 Then
            txtHorasReg.BackColor = System.Drawing.Color.Red
        End If
        If CamposObligatorios() = True Then
            'txtTotal.Text = Round((CDbl(txtDiasTot1.Text) + CDbl(txtDiasTot2.Text) + txtPlus.Text + txtDieta.Text + txtVarios.Text + txtNHNormalTotal.Text + txtNHExtraTotal.Text + txtNHEspecialTotal.Text + txtACCTotal.Text + txtCCTotal.Text + txtVacacTotal.Text) - (txtAnticipo.Text + txtAmortizacion.Text + txtEmbargos.Text + txtVarios2.Text + txtAnticipoFijo.Text), 2)
            txtTotal.Text = Round((CDbl(Nz(ntbDiasTot.Text, 0)) + CDbl(Nz(txtTotalReg.Text, 0)) + CDbl(Nz(txtTotalRegD.Text, 0)) + (Nz(txtNHNormalTotal.Text, 0)) + CDbl(Nz(txtNHExtraTotal.Text, 0)) + CDbl(Nz(txtNHEspecialTotal.Text, 0)) + Nz(txtPlus.Text, 0) + Nz(txtDieta.Text, 0) + Nz(txtVarios.Text, 0) + Nz(txtACCTotal.Text, 0) + Nz(txtCCTotal.Text, 0) + Nz(txtVACAVTotal.Text, 0)) + (CDbl(Nz(txtAnticipo.Text, 0)) + CDbl(Nz(ntbAmortizacion.Text, 0)) + CDbl(Nz(txtEmbargos.Text, 0)) + CDbl(Nz(txtVarios2.Text, 0)) + CDbl(Nz(txtAnticipoFijo.Text, 0) + CDbl(Nz(txtPrima.Text, 0)))), 2)
            'txtNAnticipo.Text = Round((CDbl(txtTotal.Text) - CDbl(txtNMensual.Text) - CDbl(txtSSocial.Text) - CDbl(txtDerechos.Text)), 2) 'LO NUEVO ES DESDE txtSSocial
            'Use 26/01/2009, pedido Fernando, que se controle por este campo
            txtNMensual.Text = Round((CDbl(Nz(txtTotal.Text, 0)) - CDbl(Nz(txtNAnticipo.Text, 0)) - CDbl(Nz(txtSSocial.Text, 0)) - CDbl(Nz(txtDerechos.Text, 0))), 2) 'LO NUEVO ES DESDE txtSSocial
        End If
    End Sub

    Private Sub btnHistoricoCalendario_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHistoricoCalendario.Click

        If advOperario1.Text = "" Then
            MessageBox.Show("Debe indicar un operario", "eXpertis 5.0", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim formHistorico As New frmHistoricoCalendario
            formHistorico.operario = advOperario1.Value
            formHistorico.mes = cbxMes.Value
            formHistorico.anio = cbxAnyo.Value
            formHistorico.Show()
        End If

    End Sub

    Private Sub btnGuardar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGuardar.Click
        btnGuardar.Enabled = False
        If txtNAnticipo.Text < 0 Then
            'solo avisa que el campo nomina anticipo es negativo pero guarda la valoracion
            Guardar()
            MsgBox("El Valor en Nomina Anticipo es Negativo", vbInformation + vbOKOnly)

        Else
            Guardar()
            'btnAceptar_Click(1, e)
        End If
        btnGuardar.Enabled = True
    End Sub

    Private Sub btnComentario_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnComentario.Click

        Select Case bValidar()
            Case 0
                MsgBox("Debes buscar los datos del operario", vbInformation + vbOKOnly)
            Case 1
                MsgBox("El registro todavia no ha sido guardado. Pulse el botón guardar antes de insertar Comentarios", vbExclamation)
            Case 2
                'David
                Dim formObservaciones As New frmObservaciones
                filtroHistorico.Clear()

                filtroHistorico.Add("idOperario", FilterOperator.Equal, advOperario1.Value)
                filtroHistorico.Add("mes", FilterOperator.Equal, cbxMes.Value)
                filtroHistorico.Add("anyo", FilterOperator.Equal, cbxAnyo.Value)


                formObservaciones.filtroHistorico = filtroHistorico

                formObservaciones.Params = filtroHistorico.Compose(New AdoFilterComposer)
                'formObservaciones.Params = sValor
                formObservaciones.Show()

        End Select

    End Sub

    Private Sub btnBorrarHr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBorrarHr.Click

        Dim frmBorrarHr As New frmBorrarHr
        Dim DE As New Expertis.Engine.BE.DataEngine

        Dim dt As New DataTable

        dt = DE.Filter("tbObraCabecera", "NObra", "DescObra='" & txtObraPredeterminada.Text & "'")

        frmBorrarHr.operario = advOperario1.Text
        'frmBorrarHr.obra = txtObraPredeterminada.Text
        frmBorrarHr.obra = dt.Rows(0)("NObra").ToString
        frmBorrarHr.Show()

    End Sub

    Private Sub btnEliminar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEliminar.Click

        Dim sSQL As String
        If MsgBox("¿Deseas realmente eliminar la valoracion?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If

        'bValidar()

        sValor = bValidarWhere()
        If sValor <> "" Then
            Dim op As New ClasesTecozam.OperarioCalendario
            op.BorrarDatos(sValor)
            'sSQL = "Delete tboperariocalendario where " & sValor
            'AdminData.Execute(sSQL)

        Else
            MsgBox("Este operario no tienen ninguna valoracion guardada con este mes y año", vbInformation + vbOKOnly)
        End If
        'Toolbar2_ButtonClick(Toolbar2.Buttons(1))
        btnAceptar_Click(1, e)

    End Sub
End Class