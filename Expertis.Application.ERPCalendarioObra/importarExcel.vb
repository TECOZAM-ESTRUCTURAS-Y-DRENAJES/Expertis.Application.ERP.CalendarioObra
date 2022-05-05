Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Business.Negocio
Imports Solmicro.Expertis.Business.Obra
Imports Solmicro.Expertis.Engine.UI
Imports Solmicro.Expertis.Engine.BE
Imports System.Math
Imports Microsoft.Office
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Business.General
Imports Solmicro.Expertis.Business.ClasesTecozam
Imports Solmicro.Expertis.Business

Public Class importarExcel

    Inherits Solmicro.Expertis.Engine.UI.FormBase
    Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

    Private Sub importarexcel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error GoTo TratarError

        'Funcion de descarga
        ' CopiarFichero()

        'Crear instancia de la clase que implementa al control embebido
        'mobjEmbeddedControl = New clsFwEmbeddedCtrl
        'Obtener referencia a la interface IEmbeddedControl del control embebido
        'mobjIEmbeddedControl = mobjEmbeddedControl
        'Si el OCX cumple con algún tipo de mantenimiento para el cual existe una clase, entonces:
        'Crear y establecer referencia a la objeto de mantenimiento
        'mobjMntoDoc = New clsFwMntoSimple
        'ReDim Preserve vIntervalo(31)
Fin:
        On Error Resume Next
        Exit Sub
TratarError:
        'GenerateError(Err.Number, Err.Description, "UserControl_Initialize", ExpertisApp.Title)
        Resume Fin
    End Sub

    Private Sub CmdUbicacion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUbicacion.Click
        CD.Filter = "Excel (*.xls)|*.xls"

        'CD.ShowOpen()
        CD.ShowDialog()

        If CD.FileName <> "" Then
            'lblRuta.Caption = CD.FileName
            lblRuta.Text = CD.FileName
        End If
    End Sub

    Public Function ObtenerDatosExcel(ByVal ruta As String, ByVal hoja As String, ByVal rango As String) As DataTable
        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim DtSet As System.Data.DataSet
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & ruta & "';Extended Properties='Excel 8.0;HDR=NO'")
        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & hoja & "$" & rango & "]", MyConnection)
        'MyCommand.TableMappings.Add("Table", "Net-informations.com")
        DtSet = New System.Data.DataSet
        MyCommand.Fill(DtSet)
        Dim dt As DataTable = DtSet.Tables(0)
        MyConnection.Close()

        Return dt

    End Function

    Sub importarExcel()

        'Dim my As System.Windows.Forms.Application

        Dim obraCab As New ObraCabecera

        Dim columna As Integer
        Dim ruta As String = lblRuta.Text
        Dim hoja As String = "Horas"
        Dim rango1 As String = "B1:B10"
        Dim rango2 As String = "A12:AG100"
        Dim rango3 As String = "A11:AG11"

        Dim empresa As String
        Dim estado As String
        Dim obra As String
        Dim trabajo As String
        Dim mes As String
        Dim numero As String
        Dim iRegistros As Integer
        Dim fecha As String
        Dim idOperario As String

        Dim hora As Double
        Dim tipoHora As String

        Dim sNombreUnicoGlobal As String
        Dim iSQL As String
        Dim sSQL As String

        Dim rsnobra As New DataTable
        Dim rs As New DataTable
        Dim dtHoras As New DataTable
        Dim dtDatos As New DataTable
        Dim dtFecha As New DataTable

        Dim f As New Filter
        dtDatos = ObtenerDatosExcel(ruta, hoja, rango1)
        dtHoras = ObtenerDatosExcel(ruta, hoja, rango2)
        dtFecha = ObtenerDatosExcel(ruta, hoja, rango3)


        empresa = dtDatos.Rows(0)(0)
        estado = dtDatos.Rows(1)(0)
        obra = dtDatos.Rows(2)(0)
        trabajo = dtDatos.Rows(3)(0)

        mes = dtDatos.Rows(9)(0)

        'iRegistros = dtHoras.Rows.Count

        'David Velasco 09/11
        'Recorro la tabla y en cuanto haya un Codigo de Operario vacio deja de leer.
        Dim cont As Integer = 0
        'Dim cont2 As Integer = 0

        For Each dr As DataRow In dtHoras.Rows
            If IsDBNull(dr(cont)) Then
                'MsgBox("El Excel tiene " & cont & " filas.")
                Exit For
            End If
            cont += 1
        Next
        'iRegistros = cont

        'David Velsaco 09/11

        iRegistros = dtHoras.Rows.Count - 1
        'MsgBox(iRegistros)

        sNombreUnicoGlobal = obra & " " & mes

        'If empresa <> ExpertisApp.EnterpriseName Then
        '    MsgBox("La empresa activa es: " & ExpertisApp.EnterpriseName & ". En la hoja excel esta intentando Importar: " & empresa & ". El proceso se cancelara", vbExclamation + vbOKOnly)
        '    
        'rs = Nothing
        'DeshacerTraspaso(sNombreUnicoGlobal)
        'If Err.Description <> "" Then
        'MsgBox("Proceso cancelado. Error: '" & Err.Description & "'", vbCritical + vbOKOnly)
        'End If
        'End If

        If estado <> "REVISADO" Then
            MsgBox("El estado del archivo es: " & estado & ". Para Importar debe ser 'Revisado'. El proceso se cancelara", vbExclamation + vbOKOnly)

            rs = Nothing
            DeshacerTraspaso(sNombreUnicoGlobal)
            If Err.Description <> "" Then
                MsgBox("Proceso cancelado. Error: '" & Err.Description & "'", vbCritical + vbOKOnly)
            End If
        End If

        f.Add("Nobra", FilterOperator.Equal, obra)

        iSQL = "Nobra= '" & obra & "'"

        rsnobra = obraCab.Filter(f, , "IDObra")

        numero = rsnobra(0)("IDObra")



        f.Clear()

        f.Add("IDObra", FilterOperator.Equal, numero)
        f.Add("CodTrabajo", FilterOperator.Equal, trabajo)

        sSQL = "IdObra=" & numero & " and Codtrabajo='" & trabajo & "'"

        Dim obraTrabajo As New ObraTrabajo

        rs = obraTrabajo.Filter(f)

        Dim idtrab As String
        idtrab = rs(0)("IDTrabajo").ToString

        If rs.Rows.Count <> 1 Then
            MsgBox("Ya hay datos insertados para este parte. Se cancela la importacion", vbCritical + vbOKOnly)
            sNombreUnicoGlobal = ""
            rs = Nothing
            DeshacerTraspaso(sNombreUnicoGlobal)
            If Err.Description <> "" Then
                MsgBox("Proceso cancelado. Error: '" & Err.Description & "'", vbCritical + vbOKOnly)
            End If

            Exit Sub

        End If

        PvProgreso.Value = 0
        PvProgreso.Maximum = dtFecha.Columns.Count - 1
        PvProgreso.Step = 1
        PvProgreso.Visible = True

        columna = 2
        'Dim cuenta As Integer = 1
        'RECORRE LAS COLUMNAS HASTA AG
        While columna < dtFecha.Columns.Count

            'MessageBox.Show("fecha: " & contador)
            Try
                fecha = dtFecha(0)(columna)
            Catch ex As Exception

            End Try


            For Each drHora As DataRow In dtHoras.Rows

                'MessageBox.Show("hora: " & cuenta)

                If Length(drHora(0)) > 0 Then
                    idOperario = drHora(0)
                    Windows.Forms.Application.DoEvents()
                    LProgreso.Text = "Importando : " & idOperario & " - " & fecha
                    Windows.Forms.Application.DoEvents()

                    If Length(drHora(columna)) > 0 Then

                        If IsNumeric(drHora(columna)) = True Then
                            hora = drHora(columna)
                            tipoHora = "HORAS"

                            Insertar(idOperario, numero, fecha, trabajo, tipoHora, hora, sNombreUnicoGlobal, numero, idtrab)
                            
                        Else
                            hora = 0
                            tipoHora = drHora(columna)
                            Insertar(idOperario, numero, fecha, trabajo, tipoHora, hora, sNombreUnicoGlobal, numero, idtrab)
                        End If
                        'cuenta = cuenta + 1
                    Else
                        'cuenta = cuenta + 1
                        Continue For
                    End If
                Else
                    Exit For
                End If

            Next
            columna = columna + 1

            If columna < dtFecha.Columns.Count Then
                PvProgreso.Value = columna
            End If


        End While
        If (PvProgreso.Value.Equals(dtFecha.Columns.Count - 1)) Then
            MsgBox("Se han insertado las " & cont & " filas correctamente.")
        End If

    End Sub

    '    Function Importar()
    '        On Error GoTo Depurar

    '        'Codigo compuesto por: lngObra + sMes
    '        Dim sNombreUnicoGlobal As String

    '        Dim sSQL As String
    '        Dim apExcel As New Excel.Application
    '        Dim iRegistros As Integer
    '        Dim I As Integer
    '        Dim x As Integer
    '        'Dim lngObra As Long 'Antes valia porque eran todos numericos
    '        Dim lngObra As String  ' Para que coja codigos tipo T073
    '        Dim sCodTrabajo As String
    '        Dim sMes As String
    '        Dim dFecha As Date
    '        Dim sOperario As String
    '        Dim dblHora As Double
    '        Dim sTipoHora As String
    '        Dim th As Threading.Thread
    '        Dim rs As DataTable
    '        Dim my As System.Windows.Forms.Application
    '        Dim iSQL As String
    '        Dim rsnobra As DataTable
    '        Dim numero As String

    '        'Dim clsAdmin As New AdminEjecutor

    '        'Declaro el objeto
    '        apExcel = CreateObject("Excel.Application")

    '        'Abro el objeto excel con la ruta del archivo que han seleccionado

    '        'apExcel.Workbooks.Open(lblRuta)
    '        apExcel.Workbooks.Open(lblRuta.Text)
    '        apExcel.Sheets(1).Select()

    '        apExcel.Range("C11").Select()

    '        'Cuento el nº de registros para la progressbar (desde A12 hasta el total de columna seleccionado con datos)
    '        iRegistros = apExcel.WorksheetFunction.CountA(apExcel.Range(apExcel.Selection, "A" & apExcel.Rows.End(Excel.XlDirection.xlToRight).Count))
    '        'apExcel.Range("C11").Select()
    '        'iRegistros = apExcel.Range(apExcel.Selection, apExcel.Selection.End(xlToRight).Count)

    '        '**********************************************
    '        'COLUMNA    CELDA           TIPO     FORMATO
    '        'EMPRESA    B1              (Texto)
    '        'ESTADO     B2              (Texto)
    '        'OBRA       B3              (Numeric)
    '        'TRABAJO    B4              (Texto)

    '        'MES:       B10             (Texto) (FORMATO: dd/mm/aaaa - dd/mm/aaaa)
    '        'DIA:       C11...RANGO(H)  (Fecha) (FORMATO: dd/mm/aaaa)
    '        'HORAS:     C12...RANGO(H)  (Texto) (FORMATO: HO (HX, HE), V, FI, FJ...)
    '        '**********************************************

    '        'apExcel.Cells(FILA, COLUMNA).Text

    '        'Compruebo la Empresa
    '        If apExcel.Cells(1, 2).text <> ExpertisApp.EnterpriseName Then
    '            MsgBox("La empresa activa es: " & ExpertisApp.EnterpriseName & ". En la hoja excel esta intentando Importar: " & apExcel.Cells(1, 2).Text & ". El proceso se cancelara", vbExclamation + vbOKOnly)
    '            GoTo Depurar
    '        End If

    '        'Compruebo el Estado
    '        If apExcel.Cells(2, 2).text <> "REVISADO" Then
    '            MsgBox("El estado del archivo es: " & apExcel.Cells(2, 2).Text & ". Para Importar debe ser 'Revisado'. El proceso se cancelara", vbExclamation + vbOKOnly)
    '            GoTo Depurar
    '        End If

    '        'Inicializo
    '        I = 3
    '        lngObra = apExcel.Cells(3, 2).text
    '        sCodTrabajo = apExcel.Cells(4, 2).text
    '        sMes = apExcel.Cells(10, 2).text

    '        sNombreUnicoGlobal = lngObra & " " & sMes

    '        'añadido para que busque el IDObra en funcion del numero de obra

    '        iSQL = "select IdObra from tbobracabecera where Nobra= '" & lngObra & "'"
    '        rsnobra = AdminData.GetData(iSQL)

    '        'If rsnobra.EOF = False Then
    '        numero = rsnobra.Rows("IdObra")(0)
    '        'End If

    '        'Antes de nada compruebo si existe el trabajo para la obra seleccionada

    '        sSQL = "IdObra=" & numero & " and Codtrabajo='" & sCodTrabajo & "'"
    '        'sSQL = "IdObra=" & lngObra & " and Codtrabajo='" & sCodTrabajo & "'"

    '        'rs = clsAdmin.Filtrar(, "tbObraTrabajo", sSQL)
    '        rs = AdminData.Filter("tbObraTrabajo", , sSQL)

    '        If rs.EOF = True Then
    '            MsgBox("El trabajo: '" & sCodTrabajo & "' no existe para la Obra: '" & lngObra & "'")
    '            GoTo Depurar
    '        End If

    '        'Compruebo si han importado esta Hoja a Real tbObraMODControl
    '        sSQL = "DescParte like '" & sNombreUnicoGlobal & "'"
    '        rs = AdminData.Filter("tbObraMODControl", , sSQL)

    '        If rs.EOF = False Then
    '            MsgBox("Ya hay datos insertados para este parte. Se cancela la importacion", vbCritical + vbOKOnly)
    '            sNombreUnicoGlobal = ""
    '            GoTo Depurar
    '        End If

    '        PvProgreso.Value = 0
    '        PvProgreso.Maximum = iRegistros
    '        PvProgreso.Step = 1
    '        PvProgreso.Visible = True

    '        'Recorro todo el fichero...
    '        'Vamos a recorrer tanto Horizonta (Las fechas, i)
    '        'como Vertical (Las horas, x)
    '        'Nos movemos por las fechas (i)
    '        Do While apExcel.Cells(11, I).text <> ""

    '            dFecha = apExcel.Cells(11, I).text
    '            'Cada vez que cogemos una Fecha tenemos que volver a recorrer todos
    '            'los Operarios
    '            x = 12
    '            'Nos movemos por el operario y las horas (x)
    '            Do While apExcel.Cells(x, 1).text <> ""
    '                sOperario = apExcel.Cells(x, 1).text
    '                my.DoEvents()
    '                'th.Sleep(100)
    '                LProgreso.Text = "Importando : " & sOperario & " - " & dFecha
    '                my.DoEvents()
    '                'th.Sleep(100)
    '                'Compruebo si no esta vacio, ya que si esta vacio no inserto nada
    '                If Len(apExcel.Cells(x, I).text) > 0 Then
    '                    'Pregunto si es numerico porque sino, es Vacaciones, falta, etc...
    '                    If IsNumeric(apExcel.Cells(x, I).text) = True Then
    '                        dblHora = apExcel.Cells(x, I).text
    '                        sTipoHora = "HORAS"
    '                        Insertar(sOperario, numero, dFecha, sCodTrabajo, sTipoHora, dblHora, sNombreUnicoGlobal, numero)
    '                    Else
    '                        dblHora = 0
    '                        sTipoHora = apExcel.Cells(x, I).text
    '                        Insertar(sOperario, numero, dFecha, sCodTrabajo, sTipoHora, dblHora, sNombreUnicoGlobal, numero)
    '                    End If
    '                End If
    '                x = x + 1
    '            Loop
    '            I = I + 1
    '            If I < iRegistros Then
    '                PvProgreso.Value = I + 1
    '            End If
    '        Loop

    '        'FINALIZA LA IMPORTACIÓN
    '        'pbImportarcion.Visible = False
    '        'Me.Caption = "Importación finalizada"

    '        apExcel.Cells(2, 2) = "LEIDO"
    '        apExcel.ActiveWorkbook.Save()
    '        my.DoEvents()
    '        'th.Sleep(100)
    '        my.DoEvents()
    '        'th.Sleep(100)
    '        my.DoEvents()

    '        'Libero memoria
    '        rs = Nothing

    '        apExcel.Workbooks.Close()
    '        apExcel.Quit()
    '        apExcel = Nothing

    '        MsgBox("Importacion Realizada con Exito", vbInformation + vbOKOnly)

    '        Exit Function
    'Depurar:
    '        'Libero memoria
    '        rs = Nothing
    '        apExcel.Workbooks.Close()
    '        apExcel.Quit()
    '        apExcel = Nothing
    '        DeshacerTraspaso(sNombreUnicoGlobal)
    '        If Err.Description <> "" Then
    '            MsgBox("Proceso cancelado. Error: '" & Err.Description & "'", vbCritical + vbOKOnly)
    '        End If
    '    End Function

    Sub Insertar(ByVal Operario As String, ByVal IdObra As String, ByVal Fecha As Date, ByVal cboTrabajo As String, ByVal sTipoHora As String, ByVal N_Horas As Double, ByVal sNombreUnico As String, ByVal numero As String, ByVal idtrab As String)

        Dim obj As New Operario
        Dim txtSQL As String
        Dim rs As New DataTable
        Dim rsTrabajo As New DataTable
        Dim rsOperario As New DataTable
        Dim rsCalendarioCentro As New DataTable
        ' Dim Conexion As ADODB.Connection
        'Dim Admin As New AdminEjecutor
        Dim IdOperacion As String
        Dim CodTrabajo As String
        Dim DescTrabajo As String
        'Dim DescTrabajo As Long
        'Dim IdTipoTrabajo As Long
        Dim IdTipoTrabajo As String
        Dim IdSubTipoTrabajo As Object
        'Dim IdSubTipoTrabajo As String
        Dim iVeces As Long
        Dim Coste_Hora As Double
        Dim Tipo_Hora As String
        Dim I As Long
        Dim IdAutonumerico As Long
        Dim HorasFacturables As Integer
        'Dim clsAdmin As New AdminEjecutor
        Dim IdTrabajo As Double
        'Dim IdTrabajo As String
        Dim HorasOrigen As Double
        'introducir fecha del sistema
        Dim dia As String
        dia = Date.Now.Date
        Dim f As New Filter
        'Antes de insertar compruebo si existe el Operario
        f.Add("IdOperario", FilterOperator.Equal, Operario)
        'rs = clsAdmin.Filtrar(, "tbMaestroOperario", txtSQL)
        rs = obj.Filter(f)

        If rs.Rows.Count = 0 Then
            MsgBox("El operario: '" & Operario & "' no existe en la BBDD. Todo el proceso se cancelara", vbExclamation + vbOKOnly)
            iVeces = "Error Provocado"
        End If

        rs = Nothing

        IdOperacion = "Guardar Datos"
        HorasOrigen = N_Horas
        Dim objTrabajo As New ObraTrabajo
        Dim filtro2 As New Filter
        Dim filtro3 As New Filter
        'Guardos los datos
        If IdOperacion = "Guardar Datos" Then

            'Conexion.Open(AdminData.GetConnectionString)

            'Obtengo datos de trabajo
            'txtSQL = "Select IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo from tbObraTrabajo where IdObra=" & IdObra & " and Codtrabajo='" & cboTrabajo & "'"
            txtSQL = "Select IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo from tbObraTrabajo where IdObra=" & numero & " and Codtrabajo='" & cboTrabajo & "'"
            'rsTrabajo = Conexion.Execute(txtSQL)
            filtro2.Add("IDObra", FilterOperator.Equal, numero)
            filtro2.Add("IdTrabajo", FilterOperator.Equal, idtrab)
            rsTrabajo = objTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

            If rsTrabajo.Rows.Count = 0 Then
                IdTrabajo = Nothing
                CodTrabajo = ""
                DescTrabajo = ""
                IdTipoTrabajo = Nothing
                IdSubTipoTrabajo = Nothing
            Else
                IdTrabajo = rsTrabajo.Rows(0)("IdTrabajo")
                CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo")
                IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo")
                IdSubTipoTrabajo = rsTrabajo.Rows(0)("IdSubtipotrabajo")
            End If

            'Obtengo datos del Operario
            txtSQL = "Select Jornada_Laboral, c_h_n, c_h_x, c_h_e from tbMaestroOperario where idoperario='" & Operario & "'"
            'rsOperario = Conexion.Execute(txtSQL)
            filtro3.Add("IDOperario", FilterOperator.Equal, Operario)
            rsOperario = obj.Filter(f, , "Jornada_Laboral, c_h_n, c_h_x, c_h_e")

            'Compruebo en el calendario
            txtSQL = "Select * from tbCalendarioCentro where idcentro='" & numero & "' and Fecha='" & Fecha & "' and tipodia=1"
            'rsCalendarioCentro = Conexion.Execute(txtSQL)
            Dim calendario As New General.CalendarioCentro
            Dim filtro As New Filter
            filtro.Add("Fecha", FilterOperator.Equal, Fecha)
            filtro.Add("IDCentro", FilterOperator.Equal, numero)
            filtro.Add("TipoDia", FilterOperator.Equal, 1)
            rsCalendarioCentro = calendario.Filter(filtro)

            'David 15/11/21 En vez de <>0 ponia "=0"
            'Si tiene datos es que es festivo
            If rsCalendarioCentro.Rows.Count <> 0 Then
                iVeces = 1
                N_Horas = N_Horas
                Coste_Hora = rsOperario.Rows(0)("c_h_e")
                Tipo_Hora = "HE"
            Else
                'Si no es festivo
                If rsOperario.Rows(0)("Jornada_Laboral") >= N_Horas Then
                    'Todas son horas normales
                    iVeces = 1
                    N_Horas = N_Horas
                    Coste_Hora = rsOperario.Rows(0)("c_h_n")
                    Tipo_Hora = "HO"
                Else
                    'Hay horas normales y horas extras, primero pongo las horas normales
                    iVeces = 2
                    Coste_Hora = rsOperario.Rows(0)("c_h_n")
                    N_Horas = rsOperario.Rows(0)("Jornada_Laboral")
                    Tipo_Hora = "HO"
                End If
            End If

            'Tipo de hora que se inserta
            If sTipoHora <> "HORAS" Then
                Tipo_Hora = sTipoHora
                iVeces = 1
            End If

            For I = 1 To iVeces

                'txtSQL = "execute xAutoNumericValueWEB 0"
                ''Conexion.Execute(txtSQL)
                'AdminData.Execute(txtSQL)

                'txtSQL = "select id from tbvalor"
                ''rs = Conexion.Execute(txtSQL)
                'rs = AdminData.GetData(txtSQL)

                ''Obtengo el IdLineaModControl
                'IdAutonumerico = rs.Fields("id").Value

                'IBIS. david. 25/08/2010
                Dim auto As New OperarioCalendario
                IdAutonumerico = auto.Autonumerico()

                'Horas Facturables
                If Trim(DescTrabajo) = "HORAS FACTURABLES" Then
                    HorasFacturables = 1
                Else
                    HorasFacturables = 0
                End If


                txtSQL = "Insert into tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                         "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                         "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IdTipoTurno) " & _
                         "Values(" & IdAutonumerico & ", " & IdTrabajo & ", " & IdObra & ", '" & _
                         CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                         IdSubTipoTrabajo & "', '" & Operario & "', 'PREDET', '" & _
                         Tipo_Hora & "', '" & Fecha & "', " & Replace(N_Horas, ",", ".") & _
                         ", " & Replace(Coste_Hora, ",", ".") & ", " & Replace(Round(CDbl(Coste_Hora) * CDbl(N_Horas), 2), ",", ".") & _
                         ", " & Replace(N_Horas, ",", ".") & ", " & Replace(Round(CDbl(Coste_Hora) * CDbl(N_Horas), 2), ",", ".") & _
                         ", '" & sNombreUnico & "', " & HorasFacturables & ", '" & dia & "', '" & dia & "', '" & ExpertisApp.UserName & "', 4)"
                'Inserto
                'Conexion.Execute(txtSQL)
                auto.Ejecutar(txtSQL)

                'Cambio valores, pongo las horas extras
                Coste_Hora = rsOperario.Rows(0)("c_h_x")
                N_Horas = CDbl(HorasOrigen) - CDbl(rsOperario.Rows(0)("Jornada_Laboral"))
                Tipo_Hora = "HX"
            Next
            'Libero memoria
            'Conexion = Nothing
            rs = Nothing
            rsTrabajo = Nothing
            rsOperario = Nothing
            rsCalendarioCentro = Nothing
        End If
    End Sub

    Sub DeshacerTraspaso(ByVal sNombreGlobal)
        'Dim clsAdmin As New AdminEjecutor

        If sNombreGlobal <> "" Then
            'clsAdmin.Ejecutar("Delete from tbObraMODControl where DescParte like '" & sNombreGlobal & "'", False)
            AdminData.Execute("Delete from tbObraMODControl where DescParte like '" & sNombreGlobal & "'")
        End If

        'Libero memoria
        'clsAdmin = Nothing

    End Sub

    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAceptar.Click
        importarExcel()
    End Sub

    Private Sub btnSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalir.Click
        Me.Close()
    End Sub

    Sub CopiarFichero()
        Dim Res As Long

        ' Este ejemplo copia el fichero AutoExec.Bat como NewExec.Bat
        ' y lo sobreescribe si existe
        Res = CopyFile("\\servidor\ExpertisNet\bin\ProgressBar.ocx", "c:\ExpertisNet\bin\ProgressBar.ocx", 0)

        Dim fs, f

        fs = CreateObject("Scripting.FileSystemObject")

        'f = fs.getfile("\\servidor\Expertis\bin\ProgressBar.ocx")
        f = fs.getfile("C:\ExpertisNet\Bin\ProgressBar.ocx")

        'If fs.FileExists("C:\Expertis\bin\ProgressBar.ocx") = False Then
        If fs.FileExists("C:\ExpertisNet\bin\ProgressBar.ocx") = False Then
            ' Para copiar el fichero con otro nombre o a otro directorio.
            'f.Copy("c:\Expertis\bin\ProgressBar.ocx", True)
            f.Copy("c:\ExpertisNet\bin\ProgressBar.ocx", True)

            'Para registrar
            'Call Shell("regsvr32 C:\Expertis\bin\ProgressBar.ocx", vbNormalFocus)
            Call Shell("regsvr32 C:\ExpertisNet\bin\ProgressBar.ocx", vbNormalFocus)
            'SendKeys("{ENTER}", False)
            SendKeys.Send("{ENTER}")
        Else
            'Para registrar
            'Call Shell("regsvr32 C:\Expertis\bin\ProgressBar.ocx", vbMaximizedFocus)
            Call Shell("regsvr32 C:\ExpertisNet\bin\ProgressBar.ocx", vbMaximizedFocus)
            'SendKeys("{ENTER}", False)
            SendKeys.Send("{ENTER}")
        End If

        'Liberar memoria
        f = Nothing
        fs = Nothing
    End Sub

    Private Sub borrahoras_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles borrahoras.Click
        Dim numobra As String
        Dim fechaDesde As String
        Dim fechaHasta As String
        Dim bSQL As String

        bSQL = Nothing

        numobra = InputBox("Introduzca Numero de obra", "Introduzca Datos:")
        fechaDesde = InputBox("Introduzca Fecha Inicio(dd/mm/aa).", "Introduzca Datos: ")
        fechaHasta = InputBox("Introduzca Fecha Fin(dd/mm/aa).", "Introduzca Datos:")

        If numobra = "" Or fechaDesde = "" Or fechaHasta = "" Then
            MsgBox("Introduzca Datos.")
        Else
            'bSQL = AdminData.Execute("SELECT COUNT (*) FROM tbObraMODControl WHERE DescParte LIKE '" & numobra & " " & fechaDesde & " - " & fechaHasta & "'")
            'MsgBox("Se van a borrar " & bSQL & " registros", vbYesNo, "Informacion")
            'If bSQL > 0 Then
            'If MsgBox("Se van a borrar " & bSQL & " registros", vbYesNo, "Informacion") = vbYes Then
            'AdminData.Execute("DELETE from tbObraMODControl WHERE DescParte LIKE = '" & numobra & " " & fechaDesde & " - " & fechaHasta & "'")

            'Comentado por David Velasco 21/3
            'Dim sql As String
            'sql = numobra & " " & fechaDesde & " - " & fechaHasta
            'Dim auto As New OperarioCalendario
            'auto.BorraDatosObraMODControl(sql)

            'Borrar por FechaInicio en vez de por DescParte
            Dim dt As New DataTable
            Dim filtro As New Filter
            filtro.Add("NObra", FilterOperator.Equal, numobra)
            dt = New BE.DataEngine().Filter("tbObraCabecera", filtro)
            Dim idobra As String

            idobra = dt.Rows(0)("IDObra").ToString()
            Dim auto As New OperarioCalendario
            auto.BorraDatosObraMODControlPorFechaInicio(idobra, fechaDesde, fechaHasta)

            'End If
            'Else
            'MsgBox("Datos incorrectos o el registro no existe. Ej sintaxis. num Obra:50032, Fecha Desde: 21/10/06, Fecha Hasta: 20/11/06")
        End If

        'End If
    End Sub

End Class
