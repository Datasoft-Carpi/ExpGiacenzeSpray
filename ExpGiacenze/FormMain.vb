Imports System.IO
Imports System.Threading
Imports Excel = Microsoft.Office.Interop.Excel

Public Class FormMain
    Dim rsGlobale As New ADODB.Recordset
    Dim PrimoGiro As Boolean = True
    Const MAX_ALTEZZA = 200

    Private Declare Auto Function GetPrivateProfileString Lib "kernel32" (ByVal lpAppName As String,
            ByVal lpKeyName As String,
            ByVal lpDefault As String,
            ByVal lpReturnedString As System.Text.StringBuilder,
            ByVal nSize As Integer,
            ByVal lpFileName As String) As Integer

    Const NUM_TAGLIE = 30
    Const idxStartTaglie = 4
    Public Const VER_35 = "3.5"
    Public Const VER_38 = "3.8"
    Public Const VER_41 = "4.1"

    Dim tmp_dir = "C:\DatasoftTmp\" 'Environment.GetFolderPath(Environment.SpecialFolder.Windows) & 

    Const MAX_ART = 1500

    Const TAG_TABLE = "<table border=1 cellspacing=1 cellpadding=0 width=""100%"">"
    Const TAG_TABLE_VUOTO = "<TABLE border=1>"
    Const TAG_TABLE_FIXED = "<TABLE class=""fixed"" border=""1"">"
    Const TAG_TABLE_END = "</TABLE>"
    Const TAG_ROW = "<TR>"
    Const TAG_ROW_END = "</TR>"

    Const TAG_CELL = "<TD align=""center"" style=""font-weight:bold"">"
    Const TAG_CELL_CALC = "<TD align=""center"" bgcolor=""#A5DFF2"" style=""font-weight:bold"">"
    Const TAG_CELL_LEFT = "<TD align=""left"" >"
    Const TAG_CELL_CALC_LEFT = "<TD align=""left"" bgcolor=""#A5DFF2""  >"
    Const TAG_CELL_END = "</TD>"

    Const TAG_CELL_DATAORA = "<TD align=""center"" bgcolor=""#FFFFAA"" style=""font-weight:bold"" >"
    Const TAG_CELL_ARTICOLO = "<TD align=""center"" bgcolor=""#AAAAAA"" style=""font-weight:bold"" >"
    Const TAG_CELL_PREZZO = "<TD align=""center"" bgcolor=""#DDDDDD"" style=""font-weight:bold"" >"
    Const TAG_CELL_TAGLIA = "<TD align=""center"" bgcolor=""#DDDDDD"" style=""font-weight:bold"" >"
    Const TAG_CELL_VARIANTE = "<TD align=""left"" bgcolor=""#DDDDDD"" style=""font-weight:bold"" >"
    Const TAG_CELL_MAGAZZINO = "<TD align=""left"" bgcolor=""#DDDDDD"" style=""font-weight:bold"" >"

    Dim TOP_LIMIT = ""
    Dim colore_disabilitato = Color.LightGray
    Dim colore_lettura_facilitata = Color.LightBlue
    Dim iniPath As String
    Dim bLoadImage As Boolean = True

    Dim Connessione As String = "Provider=sqloledb;Data Source=DSOFT06\SISTEMI;Initial Catalog=ESOLVER_SPRAY;User ID=sa;Password=Sistemi123;"
    Dim connSqlSrv As New ADODB.Connection
    Dim CurrentRowIdx As Integer = -1
    Dim loaded As Boolean = False
    Dim oldCellvalue As String

    Dim FLAG_ORDINATO = "1"
    Dim CODARTICOLO_DA = ""
    Dim CODARTICOLO_A = "ZZZZ"
    Dim CODMARCA_DA = ""
    Dim FAMIGLIA_DA = ""

    Dim TIPOART_DA = "0"
    Dim TIPOART_A = "6"

    Dim MACROFAM_DA = ""
    Dim STAGIONE_DA = ""
    Dim LINEA_DA = ""
    Dim CODFOR = "0"
    Dim CODAGE = "0"
    Dim CODZON = ""
    Dim CODMAGAZZINO = "('P006')"
    Dim LISTAMAGAZZINI = ""
    Dim CODMAGAZZINO1 = ""
    Dim CODMAGAZZINO2 = ""
    Dim CODMAGAZZINO3 = ""
    Dim CODMAGAZZINO4 = ""
    Dim CODMAGAZZINO5 = ""
    Dim CODGRUPPO = "AR"
    Dim CODUTEPERS = ""
    Dim codStatDA(20) As String
    Dim codStatA(20) As String
    Public versione As String = "4.2"
    Dim DETTAGLIO_MAG = 0

    Public FOLDER_IMG_VAR = ""
    Public EXTENSION_IMG_VAR = ""

    Public CODLIS_VEN = ""
    Dim DESLIS_VEN = ""

    Public CODLIS_ACQ = ""
    Dim DESLIS_ACQ = ""

    Dim currentUM As String = ""
    Dim currentDescEstesa As String = ""

    Structure Filtro
        Public Codice As String
        Public descrizione As String
        Public visibile As Boolean
    End Structure

    Public ListFiltriArticoli() As Filtro
    Public ListFiltriVarianti() As Filtro
    Public ListFiltriTaglie() As Filtro

    Structure Giacenze
        Public Taglie() As String
        Public Giacenze() As Integer
        Public Calcolato() As Integer
        Public DispTeorica() As Integer
    End Structure

    Structure Varianti
        Public CodArt As String
        Public DescArt As String
        Public DesEstesa As String
        Public CodStag As String
        Public DescStag As String
        Public Composizione As String
        Public Famiglia As String
        Public DescEstesa As String
        Public haTaglie As Boolean
        Public PrezzoLisVen As Double
        Public PrezzoLisAcq As Double
        Public UM As String
        Public CodMarca As String
        Public DesMarca As String
        Public CodNomenclatura As String
        Public MadeIn As String
        Public DescrInglese As String
        Public CodVariante() As String
        Public Varianti() As String
        Public TotaleGiac() As Integer
        Public TotaleCalc() As Integer
        Public Giacenze() As Giacenze
    End Structure

    Dim varTag() As Varianti

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim bConn As Boolean = True
        Dim strArg() As String
        Dim inifilname As String
        strArg = Command().Split(" ")

        inifilname = strArg(0)

        If inifilname = "" Then
            TOP_LIMIT = "top 50"
            inifilname = "c:\esmoda38\expGiacenze_spray.ini"
        End If

        iniPath = IO.Path.GetDirectoryName(inifilname) & "\"
        readIniFile(inifilname)

        Dim imginifilename = Path.GetDirectoryName(inifilname) & "\fileini\immaginivar.ini"
        If File.Exists(imginifilename) Then
            readIniFileImg(imginifilename)
        End If


        If Trim(CODLIS_ACQ) <> "" Then
            LabelLisAcq.Text = "[" & CODLIS_ACQ & "] - " & DESLIS_ACQ
            DataGridViewArticoli.Columns(4).Visible = True
        Else
            LabelLisAcq.Text = "Nessun listino di acquisto selezionato"

        End If

        If Trim(CODLIS_VEN) <> "" Then
            LabelLisVen.Text = "[" & CODLIS_VEN & "] - " & DESLIS_VEN
            DataGridViewArticoli.Columns(5).Visible = True
        Else
            LabelLisVen.Text = "Nessun listino di vendita selezionato"
        End If

        If Not System.IO.Directory.Exists(tmp_dir) Then
            System.IO.Directory.CreateDirectory(tmp_dir)
        End If

        ' collegamento al DB
        Try
            connSqlSrv.Open(Connessione)
        Catch
            bConn = False
            LabelDB.Text = "DATABASE non connesso - controllare"
            LabelDB.ForeColor = Color.Red

            ButtonExportHTML.Enabled = False
            ButtonExportEXCEL.Enabled = False

        End Try

        If bConn Then
            LabelDB.Text = "DATABASE connesso"
            LabelDB.ForeColor = Color.Black

        End If

    End Sub

    Private Function lookupStagione(codStag As String) As String
        Dim res As String = ""
        Dim rs As New ADODB.Recordset

        Try
            rs.Open("Select * from ModaTabellaStagioni where codiceStagione = '" & Replace(codStag, "'", "''") & "'", connSqlSrv, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox("Select * from ModaTabellaStagioni where codiceStagione = '" & Replace(codStag, "'", "''") & "'" & "    --- function lookupStagione")
        End Try

        If rs.RecordCount > 0 Then
            res = rs.Fields("DescrStagione_1").Value
        End If

        Return res
    End Function

    Private Function getDbNullStr(value As Object) As String

        Dim res As String
        If IsDBNull(value) = True Then
            res = ""
        Else
            res = value.ToString
        End If

        Return res

    End Function

    ' popola la griglia mastetr
    Public Sub PopolaGriglia(EseguiQuery As Boolean)
        Dim SQL As String
        'Dim idxrow As Integer = 0
        Dim srcfilename As String
        Dim bTrovato As Boolean
        Dim i As Integer
        Dim rowContatore As Integer = 0

        LabelDB.Text = "Caricamento articoli in corso ...."
        DataGridViewGiacenze.Visible = False
        ButtonExportHTML.Enabled = False
        ButtonExportEXCEL.Enabled = False
        BtnFiltri.Enabled = False
        loaded = False


        If EseguiQuery = True Then

            SQL = ""
            SQL = SQL & "SELECT "
            SQL = SQL & " ArtAnagrafica.CodArt AS [CodArticolo], "
            SQL = SQL & " ArtAnagrafica.StatoArt AS [StatoArt], "
            SQL = SQL & " ArtAnagrafica.MagUm AS [MagUm], "
            SQL = SQL & " ArtAnagrafica.RicercaAlternativa AS [DesArt], "
            SQL = SQL & " ArtAnagrafica.DesEstesa AS [DesEstesa], "
            SQL = SQL & " ArtAnagrafica.CodFamiglia AS [CodFamiglia], "
            SQL = SQL & " ArtAnagrafica.CodMarca AS [CodMarca], "
            SQL = SQL & " ModaArticoli.CodLinea AS [CodLinea], "
            SQL = SQL & " ArtAnagrafica.CodNomenclaturaComb AS [CodNomenclaturaComb], "
            SQL = SQL & " Marche.DesMarca AS [DesMarca], "
            SQL = SQL & " ModaTabellaLinee.DescrizioneLinea_1 AS [DesLinea], "
            SQL = SQL & " ModaArticoli.CodStagione AS [CodStagione], "
            SQL = SQL & " Famiglia.Des AS [DesFamiglia], "
            If versione <> VER_35 Then
                SQL = SQL & " ModaArticoli.ArtMadeIn AS [ArtMadeIn], "
            End If
            SQL = SQL & " ModaArticoli.CodStatistico1 AS [CodStatistico1], "
            If versione <> VER_35 Then
                SQL = SQL & " ModaArtComposEstesa.ComposizioneEstesa AS [ComposizioneEstesa], "
            End If

            SQL = SQL & " ModaArticoli.CodComposizione1 AS [Composizione1], "
            SQL = SQL & " ModaArticoli.CodComposizione2 AS [Composizione2], "
            SQL = SQL & " ModaTabComposizioni1.Descrizione_1 AS [DescrCompo1], "
            SQL = SQL & " ModaTabComposizioni2.Descrizione_1 AS [DescrCompo2], "
            SQL = SQL & " ArtDatiInLingua.DesEstesa AS [DescrizioneInglese], "

            SQL = SQL & " 1 AS [Contatore]"
            SQL = SQL & " FROM ArtAnagrafica "
            SQL = SQL & " LEFT OUTER JOIN ModaArticoli AS [ModaArticoli] ON (ModaArticoli.CodiceArticolo=ArtAnagrafica.CodArt AND (ModaArticoli.DBGruppo=ArtAnagrafica.DBGruppo))"
            SQL = SQL & " LEFT OUTER JOIN Famiglia AS [Famiglia] ON (Famiglia.CodFamiglia=ArtAnagrafica.CodFamiglia AND (Famiglia.DBGruppo=ArtAnagrafica.DBGruppo))"

            If versione <> VER_35 Then
                SQL = SQL & " LEFT OUTER JOIN ModaArtComposEstesa As [ModaArtComposEstesa] On (ModaArtComposEstesa.CodiceArticolo=ArtAnagrafica.CodArt And (ModaArtComposEstesa.DBGruppo=ArtAnagrafica.DBGruppo))"
            End If

            SQL = SQL & " LEFT OUTER JOIN ModaTabComposizioni As [ModaTabComposizioni1] On (ModaTabComposizioni1.CodiceComposizione=ModaArticoli.CodComposizione1 And (ModaTabComposizioni1.DBGruppo=ModaArticoli.DBGruppo))"
            SQL = SQL & " LEFT OUTER JOIN ModaTabComposizioni As [ModaTabComposizioni2] On (ModaTabComposizioni2.CodiceComposizione=ModaArticoli.CodComposizione2 And (ModaTabComposizioni2.DBGruppo=ModaArticoli.DBGruppo))"

            SQL = SQL & " LEFT OUTER JOIN ArtDatiInLingua As [ArtDatiInLingua] On (ArtDatiInLingua.CodArt=ArtAnagrafica.CodArt And ArtDatiInLingua.VarianteArt = '' And ArtDatiInLingua.CodLingua = 5 And (ArtDatiInLingua.DBGruppo=ArtAnagrafica.DBGruppo))"

            SQL = SQL & " LEFT OUTER JOIN Marche As [Marche] On (ArtAnagrafica.CodMarca = Marche.CodMarca And (Marche.DBGruppo=ArtAnagrafica.DBGruppo)) "
            SQL = SQL & " LEFT OUTER JOIN ModaTabellaLinee As [ModaTabellaLinee] On (ModaArticoli.CodLinea = ModaTabellaLinee.CodiceLinea And (ModaArticoli.DBGruppo=ModaTabellaLinee.DBGruppo)) "

            SQL = SQL & " WHERE (ArtAnagrafica.TipoAnagr = 1) AND ModaArticoli.codstagione <> '' "
            SQL = SQL & " And (ArtAnagrafica.CodArt BETWEEN '" & CODARTICOLO_DA & "' AND '" & CODARTICOLO_A & "') "

            If FAMIGLIA_DA <> "" Then
                SQL = SQL & " AND (ArtAnagrafica.CodFamiglia IN ('" & FAMIGLIA_DA & "')) "
            End If

            If CODMARCA_DA <> "" Then
                SQL = SQL & " AND (ArtAnagrafica.CodMarca IN ('" & CODMARCA_DA & "')) "
            End If

            If STAGIONE_DA <> "" Then
                SQL = SQL & " AND (ModaArticoli.codstagione IN ('" & STAGIONE_DA & "')) "
            End If

            If LINEA_DA <> "" Then
                SQL = SQL & " AND (ModaArticoli.codlinea IN ('" & LINEA_DA & "')) "
            End If

            If MACROFAM_DA <> "" Then
                SQL = SQL & " AND (Famiglia.codMacrofamiglia IN ('" & MACROFAM_DA & "')) "
            End If

            SQL = SQL & " AND (ArtAnagrafica.DBGruppo='" & CODGRUPPO & "') "
            SQL = SQL & " AND (ArtAnagrafica.TipoArt BETWEEN '" & TIPOART_DA & "' AND '" & TIPOART_A & "') "


            For i = 1 To 20
                If codStatDA(i) <> "" Then
                    SQL = SQL & " And (ModaArticoli.codStatistico" & i & " IN ('" & codStatDA(i) & "')) "
                End If
            Next

            If CODFOR <> 0 Then
                SQL = SQL & " AND (ArtAnagrafica.AcqCodForAbituale='" & CODFOR & "') "
            End If

            SQL = SQL & " ORDER BY [CodArticolo] ASC"


            Try
                rsGlobale.Open(SQL, connSqlSrv, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Catch ex As Exception
                Clipboard.SetText(SQL)
                MsgBox(ex.Message)
                'MsgBox(SQL & "   ----PopolaGriglia")
            End Try

        Else
            ' in caso sia già caricata
            rsGlobale.MoveFirst()
            DataGridViewArticoli.Rows.Clear()
            DataGridViewGiacenze.Rows.Clear()

        End If

        ProgressBarArticoli.Value = 0
        ProgressBarArticoli.Minimum = 0

        If rsGlobale.RecordCount > MAX_ART Then
            bLoadImage = False
            If MsgBox("Superata la soglia di numero articoli per la visualizzazione immagini (" & MAX_ART & ") - vuoi procedere senza immagini?", MsgBoxStyle.YesNo, "Conferma") = vbNo Then
                End
            End If

        End If

        ' azzera struttura che contiene le varianti
        ReDim varTag(0)


        rowContatore = -1
        ProgressBarArticoli.Maximum = rsGlobale.RecordCount
        While Not rsGlobale.EOF
            ProgressBarArticoli.Value = ProgressBarArticoli.Value + 1
            LabelDB.Text = "Caricamento articoli in corso (" & ProgressBarArticoli.Value & "/" & ProgressBarArticoli.Maximum & ")...."

            currentUM = rsGlobale.Fields("MagUM").Value

            currentDescEstesa = rsGlobale.Fields("CodArticolo").Value & " - " & getDbNullStr(rsGlobale.Fields("DesEstesa").Value) & vbCrLf
            If rsGlobale.Fields("CodStagione").Value <> "" Then
                currentDescEstesa = currentDescEstesa & "Stag: " & rsGlobale.Fields("CodStagione").Value & " - " & lookupStagione(rsGlobale.Fields("CodStagione").Value) & vbCrLf
            End If

            If getDbNullStr(rsGlobale.Fields("CodFamiglia").Value) <> "" Then
                currentDescEstesa = currentDescEstesa & "Fam: " & rsGlobale.Fields("CodFamiglia").Value & " - " & rsGlobale.Fields("DesFamiglia").Value & vbCrLf
            End If

            If versione <> VER_35 Then
                If getDbNullStr(rsGlobale.Fields("ComposizioneEstesa").Value) <> "" Then
                    currentDescEstesa = currentDescEstesa & "Comp: " & rsGlobale.Fields("ComposizioneEstesa").Value & vbCrLf
                Else
                    If getDbNullStr(rsGlobale.Fields("DescrCompo1").Value) <> "" Then
                        currentDescEstesa = currentDescEstesa & "Comp: " & rsGlobale.Fields("DescrCompo1").Value
                    End If

                    If getDbNullStr(rsGlobale.Fields("DescrCompo2").Value) <> "" Then
                        currentDescEstesa = currentDescEstesa & rsGlobale.Fields("DescrCompo2").Value
                    End If

                    currentDescEstesa = currentDescEstesa & vbCrLf
                End If
            End If

            'If getDbNullStr(rsGlobale.Fields("CodMarca").Value) <> "" Then
            '    currentDescEstesa = currentDescEstesa & "Marca: " & rsGlobale.Fields("CodMarca").Value & " - " & rsGlobale.Fields("DesMarca").Value & vbCrLf
            'End If

            If getDbNullStr(rsGlobale.Fields("CodLinea").Value) <> "" Then
                currentDescEstesa = currentDescEstesa & "Linea: " & rsGlobale.Fields("CodLinea").Value & " - " & rsGlobale.Fields("DesLinea").Value & vbCrLf
            End If

            If getDbNullStr(rsGlobale.Fields("CodNomenclaturaComb").Value) <> "" Then
                currentDescEstesa = currentDescEstesa & "Nomenclatura: " & rsGlobale.Fields("CodNomenclaturaComb").Value & vbCrLf
            End If

            If versione <> VER_35 Then
                If getDbNullStr(rsGlobale.Fields("ArtMadeIn").Value) <> "" Then
                    currentDescEstesa = currentDescEstesa & "MadeIn: " & rsGlobale.Fields("ArtMadeIn").Value & vbCrLf
                End If
            End If

            ' gestisce i filtri
            If checkVisibilitaArticolo(rsGlobale.Fields("CodArticolo").Value) = False Then
                rowContatore = rowContatore + 1
                DataGridViewArticoli.Rows.Add()
                DataGridViewArticoli.Rows(rowContatore).Cells(6).Value = rsGlobale.Fields("CodArticolo").Value
                DataGridViewArticoli.Rows(rowContatore).Visible = False
                rsGlobale.MoveNext()
                Application.DoEvents()
                Continue While
            End If

            bTrovato = PopolaStrutturaVariantiTaglie(rsGlobale.Fields("CodArticolo").Value)

            'MsgBox("EXIT From STRUCT")

            ' lo aggiunge in griglia solamente se ha trovato le condizioni di giacenza
            If bTrovato = True Then
                rowContatore = rowContatore + 1
                DataGridViewArticoli.Rows.Add()
                DataGridViewArticoli.Rows(rowContatore).Cells(6).Value = rsGlobale.Fields("CodArticolo").Value

                ' compila la struttura per i filtri al primo giro
                If PrimoGiro = True Then
                    Call AddArticoloInList(rsGlobale.Fields("CodArticolo").Value, getDbNullStr(rsGlobale.Fields("DesEstesa").Value))
                End If

                'MsgBox("LOAD IMAGE")

                If bLoadImage = True Then
                    srcfilename = getFilePath(rsGlobale.Fields("CodArticolo").Value)

                    If srcfilename <> "" Then
                        DataGridViewArticoli.Rows(rowContatore).Cells(0).Style.BackColor = Color.White
                        DataGridViewArticoli.Rows(rowContatore).Cells(0).Value = Image.FromFile(srcfilename)
                    Else
                        DataGridViewArticoli.Rows(rowContatore).Cells(0).Value = My.Resources.url
                    End If

                Else
                    DataGridViewArticoli.Columns(0).Visible = False
                End If

                'MsgBox("ENDING LOAD IMAGE")


                varTag(UBound(varTag) - 1).CodArt = rsGlobale.Fields("CodArticolo").Value
                varTag(UBound(varTag) - 1).DescArt = rsGlobale.Fields("DesArt").Value
                varTag(UBound(varTag) - 1).DescEstesa = rsGlobale.Fields("DesEstesa").Value
                varTag(UBound(varTag) - 1).CodStag = rsGlobale.Fields("CodStagione").Value
                varTag(UBound(varTag) - 1).DescStag = lookupStagione(rsGlobale.Fields("CodStagione").Value)
                varTag(UBound(varTag) - 1).Famiglia = rsGlobale.Fields("CodFamiglia").Value & " - " & rsGlobale.Fields("DesFamiglia").Value

                If getDbNullStr(rsGlobale.Fields("ComposizioneEstesa").Value) <> "" Then
                    varTag(UBound(varTag) - 1).Composizione = getDbNullStr(rsGlobale.Fields("ComposizioneEstesa").Value)
                Else
                    varTag(UBound(varTag) - 1).Composizione = getDbNullStr(rsGlobale.Fields("DescrCompo1").Value)

                    If getDbNullStr(rsGlobale.Fields("DescrCompo2").Value) <> "" Then
                        varTag(UBound(varTag) - 1).Composizione = varTag(UBound(varTag) - 1).Composizione & getDbNullStr(rsGlobale.Fields("DescrCompo2").Value)
                    End If
                End If

                varTag(UBound(varTag) - 1).CodMarca = getDbNullStr(rsGlobale.Fields("CodLinea").Value)
                varTag(UBound(varTag) - 1).DesMarca = getDbNullStr(rsGlobale.Fields("DesLinea").Value)
                varTag(UBound(varTag) - 1).CodNomenclatura = getDbNullStr(rsGlobale.Fields("CodNomenclaturaComb").Value)
                varTag(UBound(varTag) - 1).MadeIn = getDbNullStr(rsGlobale.Fields("ArtMadeIn").Value)
                varTag(UBound(varTag) - 1).DescrInglese = getDbNullStr(rsGlobale.Fields("DescrizioneInglese").Value)


                varTag(UBound(varTag) - 1).PrezzoLisAcq = 0
                If CODLIS_ACQ <> "" Then
                    varTag(UBound(varTag) - 1).PrezzoLisAcq = getPrezzoListini(CODLIS_ACQ, rsGlobale.Fields("CodArticolo").Value)
                End If

                varTag(UBound(varTag) - 1).PrezzoLisVen = 0
                If CODLIS_VEN <> "" Then
                    varTag(UBound(varTag) - 1).PrezzoLisVen = getPrezzoListini(CODLIS_VEN, rsGlobale.Fields("CodArticolo").Value)
                End If

                DataGridViewArticoli.Rows(rowContatore).Cells(1).Value = currentDescEstesa
                DataGridViewArticoli.Rows(rowContatore).Cells(2).Value = rsGlobale.Fields("CodArticolo").Value
                DataGridViewArticoli.Rows(rowContatore).Cells(3).Value = rowContatore
                DataGridViewArticoli.Rows(rowContatore).Cells(4).Value = varTag(UBound(varTag) - 1).PrezzoLisAcq
                DataGridViewArticoli.Rows(rowContatore).Cells(5).Value = varTag(UBound(varTag) - 1).PrezzoLisVen


                'idxrow = idxrow + 1
            End If


            rsGlobale.MoveNext()
            Application.DoEvents()
        End While

        'rs.Close()
        LabelArticoli.Text = "Articoli: " & rowContatore


        If DataGridViewArticoli.RowCount = 0 Then
            Try
                Clipboard.SetText(rsGlobale.Source.ToString)

            Catch ex As Exception

            End Try
            LabelDB.Text = "Caricamento articoli completato"
            MsgBox("Nessun articolo presente con questi parametri di ricerca")
            Exit Sub
        End If


        loaded = True
        DataGridViewGiacenze.Visible = True
        ButtonExportHTML.Enabled = True
        ButtonExportEXCEL.Enabled = True
        BtnFiltri.Enabled = True

        ' chiude il flag che aggiorna la struttura per i filtri il primo giro
        PrimoGiro = False
        LabelDB.Text = "Caricamento articoli completato"
        Application.DoEvents()

        ' imposta la prima riga

        'DataGridViewArticoli.CurrentCell = DataGridViewArticoli.Rows(0).Cells(1)
        Call PopolaGrigliaVariantiTaglie(DataGridViewGiacenze)
        CurrentRowIdx = DataGridViewArticoli.CurrentCell.RowIndex



    End Sub

    Private Function getPrezzoListini(codListino As String, codArt As String) As Double
        Dim strOggi As String = Format(Today, "yyyyMMdd")
        Dim rs As New ADODB.Recordset
        Dim SQL As String
        Dim res As Double = 0

        'SQL = "Select * from ListiniRigaArticolo "
        'SQL = SQL & " where 1=1"
        'SQL = SQL & " and DBGRUPPO = '" & CODGRUPPO & "'"
        'SQL = SQL & " and codListino = '" & codListino & "'"
        'SQL = SQL & " and dataInizioValidita <= '" & strOggi & "'"
        'SQL = SQL & " and ((dataFineValidita >= '" & strOggi & "') or (dataFineValidita = '18000101'))"
        'SQL = SQL & " and codArt = '" & codArt & "'"
        'SQL = SQL & " order by dataInizioValidita desc"

        SQL = "SELECT TOP 1 * FROM JSV_LIS_PRZVALDATA "
        SQL = SQL & " WHERE DbGruppo = '" & CODGRUPPO & "'"
        SQL = SQL & " AND CodListino = '" & codListino & "'"
        SQL = SQL & " AND DataIniVal <= '" & strOggi & "'"
        SQL = SQL & " AND ((DataFineVal >= '" & strOggi & "') or (DataFineVal = '18000101'))"
        SQL = SQL & " AND CodArticolo = '" & codArt & "'"
        SQL = SQL & " ORDER BY DataIniVal DESC"

        Try
            rs.Open(SQL, Connessione, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        Catch ex As Exception
            'MsgBox(ex.Message)
            MsgBox(SQL & "   ----getPrezzoListini")
        End Try

        If rs.RecordCount > 0 Then
            res = rs.Fields("PrezzoLis").Value
        Else
            res = 0
        End If
        rs.Close()


        Return res
    End Function

    Private Function getFileVarianti(codArt As String, codvar As String) As String
        Dim res As String
        res = FOLDER_IMG_VAR & codArt & "_" & codvar & "." & EXTENSION_IMG_VAR

        Return res
    End Function

    Private Function getFilePath(codArt As String) As String
        Dim rs As New ADODB.Recordset
        Dim path As String
        Dim extension As String
        Dim codice As String = Trim(codArt)
        Dim filename As String = ""
        Dim sql As String

        Try
            rs.Open("Select * from ModaTabellaLink where TipoAnagrafica = 'A' order by linkpredefinito desc", connSqlSrv, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox("qry ModatabellaLink   ---getFilePath")
        End Try

        While Not rs.EOF
            extension = rs.Fields("EstensioneLink").Value
            path = rs.Fields("PathLink").Value
            filename = path & codice & "." & extension

            ' trovato il file esce
            If System.IO.File.Exists(filename) Then
                Exit While
            Else
                filename = ""
            End If

            rs.MoveNext()
            Application.DoEvents()
        End While
        rs.Close()


        ' nel caso sia la 4.1 non controlla le immagini nella tabella link
        If versione = VER_41 Then
            Return filename
        End If

        ' sen non l'ha trovato controlla nell'altra tabella
        Dim orderby As String = ""

        If filename = "" Then

            If versione <> VER_35 Then
                orderby = " order by ImmagPreferenziale desc"
            End If

            sql = "Select * "
            sql = sql & "from ArtAllegatiAnnotaz LEFT OUTER JOIN Allegati ON ArtAllegatiAnnotaz.IdAllegato = Allegati.IdAllegato "
            sql = sql & "where ArtAllegatiAnnotaz.CodArt = '" & Replace(Trim(codArt), "'", "''") & "' and Allegati.FileImmagine = 1 "
            sql = sql & orderby

            rs.Open(sql, connSqlSrv, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

            While Not rs.EOF

                filename = rs.Fields("DirettivaAllegato").Value & rs.Fields("NomeFileAllegato").Value

                ' trovato il file esce
                If System.IO.File.Exists(filename) Then
                    Exit While
                Else
                    filename = ""
                End If

                rs.MoveNext()
                Application.DoEvents()
            End While

            rs.Close()
        End If


        Return filename
    End Function


    Private Function getSQLGiacenze(codArt As String) As String
        Dim SQL As String

        SQL = "      SELECT                                                     "
        SQL = SQL & "            ModaProgMagCor.GtgCart AS [GtgCart],                       "
        SQL = SQL & "            ModaProgMagCor.GtgVarart As [GtgVarart],                   "
        SQL = SQL & "            ArtConfigVariante.descrizione AS [descrizione],            "
        SQL = SQL & "            ArtDatiInLingua.DesEstesa AS [DescrInglese],               "
        SQL = SQL & "            Sum(MagProgrArticoli.qtaGiacUMMag) As [qtaGiacUMMag],      "
        SQL = SQL & "            Sum(ESV_IMPORD_AVMA_S.qtordum) As [qtOrdUM],               "
        SQL = SQL & "            sum(ESV_IMPORD_AVMA_S.qtimpum) As [qtimpUm],               "
        SQL = SQL & "            ModaArticoli.CodiceTabellaTaglie  AS [CodiceTabellaTaglie],"
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_1  As [CodiciTaglie_1],     "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_2  AS [CodiciTaglie_2],     "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_3  As [CodiciTaglie_3],     "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_4  AS [CodiciTaglie_4],     "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_5  As [CodiciTaglie_5],     "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_6  AS [CodiciTaglie_6],     "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_7  As [CodiciTaglie_7],     "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_8  AS [CodiciTaglie_8],     "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_9  As [CodiciTaglie_9],     "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_10  AS [CodiciTaglie_10],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_11  As [CodiciTaglie_11],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_12  AS [CodiciTaglie_12],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_13  As [CodiciTaglie_13],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_14  AS [CodiciTaglie_14],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_15  As [CodiciTaglie_15],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_16  AS [CodiciTaglie_16],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_17  As [CodiciTaglie_17],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_18  AS [CodiciTaglie_18],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_19  As [CodiciTaglie_19],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_20  AS [CodiciTaglie_20],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_21  As [CodiciTaglie_21],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_22  AS [CodiciTaglie_22],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_23  As [CodiciTaglie_23],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_24  AS [CodiciTaglie_24],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_25  As [CodiciTaglie_25],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_26  AS [CodiciTaglie_26],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_27  As [CodiciTaglie_27],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_28  AS [CodiciTaglie_28],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_29  As [CodiciTaglie_29],   "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_30  AS [CodiciTaglie_30],   "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp1) As [GtgQgp1],                  "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp2) As [GtgQgp2],                  "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp3) As [GtgQgp3],                  "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp4) As [GtgQgp4],                  "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp5) As [GtgQgp5],                  "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp6) As [GtgQgp6],                  "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp7) As [GtgQgp7],                  "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp8) As [GtgQgp8],                  "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp9) As [GtgQgp9],                  "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp10) As [GtgQgp10],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp11) As [GtgQgp11],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp12) As [GtgQgp12],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp13) As [GtgQgp13],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp14) As [GtgQgp14],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp15) As [GtgQgp15],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp16) As [GtgQgp16],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp17) As [GtgQgp17],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp18) As [GtgQgp18],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp19) As [GtgQgp19],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp20) As [GtgQgp20],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp21) As [GtgQgp21],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp22) As [GtgQgp22],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp23) As [GtgQgp23],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp24) As [GtgQgp24],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp25) As [GtgQgp25],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp26) As [GtgQgp26],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp27) As [GtgQgp27],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp28) As [GtgQgp28],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp29) As [GtgQgp29],                "
        SQL = SQL & "            sum(ModaProgMagCor.GtgQgp30) As [GtgQgp30],                "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf1) As [qtimpconf1],           "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf2) As [qtimpconf2],           "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf3) As [qtimpconf3],           "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf4) As [qtimpconf4],           "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf5) As [qtimpconf5],           "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf6) As [qtimpconf6],           "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf7) As [qtimpconf7],           "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf8) As [qtimpconf8],           "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf9) As [qtimpconf9],           "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf10) As [qtimpconf10],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf11) As [qtimpconf11],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf12) As [qtimpconf12],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf13) As [qtimpconf13],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf14) As [qtimpconf14],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf15) As [qtimpconf15],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf16) As [qtimpconf16],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf17) As [qtimpconf17],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf18) As [qtimpconf18],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf19) As [qtimpconf19],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf20) As [qtimpconf20],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf21) As [qtimpconf21],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf22) As [qtimpconf22],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf23) As [qtimpconf23],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf24) As [qtimpconf24],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf25) As [qtimpconf25],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf26) As [qtimpconf26],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf27) As [qtimpconf27],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf28) As [qtimpconf28],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf29) As [qtimpconf29],         "
        SQL = SQL & "            sum(JSV_IMPE_AVMA_D.qtimpconf30) As [qtimpconf30],         "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm1) As [QtOrdConfUm1],     "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm2) As [QtOrdConfUm2],     "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm3) As [QtOrdConfUm3],     "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm4) As [QtOrdConfUm4],     "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm5) As [QtOrdConfUm5],     "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm6) As [QtOrdConfUm6],     "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm7) As [QtOrdConfUm7],     "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm8) As [QtOrdConfUm8],     "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm9) As [QtOrdConfUm9],     "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm10) As [QtOrdConfUm10],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm11) As [QtOrdConfUm11],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm12) As [QtOrdConfUm12],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm13) As [QtOrdConfUm13],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm14) As [QtOrdConfUm14],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm15) As [QtOrdConfUm15],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm16) As [QtOrdConfUm16],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm17) As [QtOrdConfUm17],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm18) As [QtOrdConfUm18],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm19) As [QtOrdConfUm19],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm20) As [QtOrdConfUm20],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm21) As [QtOrdConfUm21],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm22) As [QtOrdConfUm22],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm23) As [QtOrdConfUm23],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm24) As [QtOrdConfUm24],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm25) As [QtOrdConfUm25],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm26) As [QtOrdConfUm26],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm27) As [QtOrdConfUm27],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm28) As [QtOrdConfUm28],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm29) As [QtOrdConfUm29],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtOrdConfUm30) As [QtOrdConfUm30],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm1) As [QtaODCUm1],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm2) As [QtaODCUm2],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm3) As [QtaODCUm3],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm4) As [QtaODCUm4],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm5) As [QtaODCUm5],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm6) As [QtaODCUm6],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm7) As [QtaODCUm7],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm8) As [QtaODCUm8],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm9) As [QtaODCUm9],   "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm10) As [QtaODCUm10], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm11) As [QtaODCUm11], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm12) As [QtaODCUm12], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm13) As [QtaODCUm13], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm14) As [QtaODCUm14], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm15) As [QtaODCUm15], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm16) As [QtaODCUm16], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm17) As [QtaODCUm17], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm18) As [QtaODCUm18], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm19) As [QtaODCUm19], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm20) As [QtaODCUm20], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm21) As [QtaODCUm21], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm22) As [QtaODCUm22], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm23) As [QtaODCUm23], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm24) As [QtaODCUm24], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm25) As [QtaODCUm25], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm26) As [QtaODCUm26], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm27) As [QtaODCUm27], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm28) As [QtaODCUm28], "
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm29) As [QtaODCUm29],"
        SQL = SQL & "            sum(JSV_IMPORD_AVMA_D.QtaODCUm30) As [QtaODCUm30]"
        SQL = SQL & "            FROM MagProgrArticoli As [MagProgrArticoli]"
        SQL = SQL & "            LEFT OUTER JOIN ModaProgMagCor As [ModaProgMagCor] On (ModaProgMagCor.GtgCart=MagProgrArticoli.CodArt And ModaProgMagCor.GtgVarart = MagProgrArticoli.varianteart And ModaProgMagCor.Gtgcmag= MagProgrArticoli.codmag And ModaProgMagCor.DBGruppo=MagProgrArticoli.DBGruppo and ModaProgMagCor.GtgCcom = MagProgrArticoli.CodAreaMag)"
        SQL = SQL & "            LEFT OUTER JOIN ModaArticoli As [ModaArticoli] On (ModaArticoli.CodiceArticolo=ModaProgMagCor.GtgCart And (ModaArticoli.DBGruppo=ModaProgMagCor.DBGruppo))"

        SQL = SQL & "            LEFT OUTER JOIN ESV_IMPORD_AVMA_S As [ESV_IMPORD_AVMA_S] On (ESV_IMPORD_AVMA_S.CodArt=MagProgrArticoli.codArt And ESV_IMPORD_AVMA_S.VarianteArt=MagProgrArticoli.VarianteArt And ESV_IMPORD_AVMA_S.DBGruppo=MagProgrArticoli.DBGruppo and ESV_IMPORD_AVMA_S.CodAreaMag = MagProgrArticoli.CodAreaMag and ESV_IMPORD_AVMA_S.codmag = MagProgrArticoli.codmag )"
        SQL = SQL & "            LEFT OUTER JOIN JSV_IMPE_AVMA_D As [JSV_IMPE_AVMA_D] On (JSV_IMPE_AVMA_D.CodArt=ModaProgMagCor.GtgCart And JSV_IMPE_AVMA_D.VarianteArt=ModaProgMagCor.GtgVarart And (JSV_IMPE_AVMA_D.DBGruppo=ModaProgMagCor.DBGruppo)  and JSV_IMPE_AVMA_D.CodMag = ModaProgMagCor.GtgCMag AND JSV_IMPE_AVMA_D.CodAreaMag = ModaProgMagCor.GtgCCom)"
        SQL = SQL & "            LEFT OUTER JOIN JSV_IMPORD_AVMA_D As [JSV_IMPORD_AVMA_D] On (JSV_IMPORD_AVMA_D.CodArt=ModaProgMagCor.GtgCart And JSV_IMPORD_AVMA_D.VarianteArt=ModaProgMagCor.GtgVarart  And (JSV_IMPORD_AVMA_D.DBGruppo=ModaProgMagCor.DBGruppo)  and JSV_IMPORD_AVMA_D.CodMag = ModaProgMagCor.GtgCMag and JSV_IMPORD_AVMA_D.CodAreaMag = ModaProgMagCor.GtgCCom)"

        SQL = SQL & "            LEFT OUTER JOIN ModaTabellaTaglie As [ModaTabellaTaglie] On (ModaTabellaTaglie.CodiceTabellaTaglie=ModaArticoli.CodiceTabellaTaglie And (ModaTabellaTaglie.DBGruppo=ModaArticoli.DBGruppo))"
        SQL = SQL & "            LEFT OUTER JOIN ArtConfigVariante As [ArtConfigVariante] On ((ArtConfigVariante.CodArt=ModaProgMagCor.GtgCart) And (ArtConfigVariante.VarianteArt=ModaProgMagCor.GtgVarart) And (ArtConfigVariante.DBGruppo=ModaArticoli.DBGruppo))"

        SQL = SQL & "            LEFT OUTER JOIN ArtDatiInLingua As [ArtDatiInLingua] On ((ArtDatiInLingua.CodArt=ArtConfigVariante.CodArt) And (ArtDatiInLingua.VarianteArt=ArtConfigVariante.VarianteArt) And ArtDatiInLingua.CodLingua=1 And (ArtDatiInLingua.DbGruppo=ArtConfigVariante.DBGruppo))"

        SQL = SQL & " WHERE (ModaProgMagCor.GtgCart = '" & Replace(Trim(codArt), "'", "''") & "') "
        SQL = SQL & " AND (ModaProgMagCor.Gtgcmag in " & CODMAGAZZINO & ")"
        SQL = SQL & " AND (ModaProgMagCor.GtgQgp1+ModaProgMagCor.GtgQgp2+ModaProgMagCor.GtgQgp3+ModaProgMagCor.GtgQgp4+ModaProgMagCor.GtgQgp5+ModaProgMagCor.GtgQgp6+ModaProgMagCor.GtgQgp7+ModaProgMagCor.GtgQgp8+ModaProgMagCor.GtgQgp9+ModaProgMagCor.GtgQgp10+"
        SQL = SQL & "ModaProgMagCor.GtgQgp11+ModaProgMagCor.GtgQgp12+ModaProgMagCor.GtgQgp13+ModaProgMagCor.GtgQgp14+ModaProgMagCor.GtgQgp15+ModaProgMagCor.GtgQgp16+ModaProgMagCor.GtgQgp17+ModaProgMagCor.GtgQgp18+ModaProgMagCor.GtgQgp19+ModaProgMagCor.GtgQgp20+"
        SQL = SQL & "ModaProgMagCor.GtgQgp21+ModaProgMagCor.GtgQgp22+ModaProgMagCor.GtgQgp23+ModaProgMagCor.GtgQgp24+ModaProgMagCor.GtgQgp25+ModaProgMagCor.GtgQgp26+ModaProgMagCor.GtgQgp27+ModaProgMagCor.GtgQgp28+ModaProgMagCor.GtgQgp29+ModaProgMagCor.GtgQgp30) > 0"
        SQL = SQL & " AND (ModaProgMagCor.DBGruppo='" & CODGRUPPO & "')"
        SQL = SQL & "            group by GtgCart, GtgVarart, ArtConfigVariante.descrizione, ArtDatiInLingua.DesEstesa, "
        SQL = SQL & "            ModaArticoli.CodiceTabellaTaglie,  "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_1,  "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_2,  "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_3,  "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_4,  "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_5,  "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_6,  "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_7,  "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_8,  "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_9,  "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_10, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_11, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_12, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_13, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_14, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_15, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_16, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_17, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_18, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_19, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_20, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_21, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_22, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_23, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_24, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_25, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_26, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_27, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_28, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_29, "
        SQL = SQL & "            ModaTabellaTaglie.CodiciTaglie_30  "



        Return SQL
    End Function


    Private Function getSQLImpegnato(codArt As String) As String
        Dim SQL As String

        SQL = " SELECT                                                         "
        SQL = SQL & " ESV_IMPORD_AVMA_S.qtordUm AS [qtordUm],                        "
        SQL = SQL & " ESV_IMPORD_AVMA_S.qtimpum As [qtimpum],                        "
        SQL = SQL & " JSV_IMPE_AVMA_D.codArt AS [GtgCart],                           "
        SQL = SQL & " JSV_IMPE_AVMA_D.VarianteArt As [GtgVarart],                    "
        SQL = SQL & " ArtConfigVariante.descrizione AS [descrizione],                "
        SQL = SQL & " ISNULL(ArtDatiInLingua.DesEstesa,'') AS [DescrInglese],               "
        SQL = SQL & " ModaArticoli.CodiceTabellaTaglie As [CodiceTabellaTaglie],    "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_1 AS [CodiciTaglie_1],          "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_2 As [CodiciTaglie_2],          "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_3 AS [CodiciTaglie_3],          "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_4 As [CodiciTaglie_4],          "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_5 AS [CodiciTaglie_5],          "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_6 As [CodiciTaglie_6],          "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_7 AS [CodiciTaglie_7],          "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_8 As [CodiciTaglie_8],          "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_9 AS [CodiciTaglie_9],          "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_10 As [CodiciTaglie_10],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_11 AS [CodiciTaglie_11],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_12 As [CodiciTaglie_12],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_13 AS [CodiciTaglie_13],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_14 As [CodiciTaglie_14],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_15 AS [CodiciTaglie_15],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_16 As [CodiciTaglie_16],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_17 AS [CodiciTaglie_17],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_18 As [CodiciTaglie_18],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_19 AS [CodiciTaglie_19],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_20 As [CodiciTaglie_20],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_21 AS [CodiciTaglie_21],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_22 As [CodiciTaglie_22],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_23 AS [CodiciTaglie_23],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_24 As [CodiciTaglie_24],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_25 AS [CodiciTaglie_25],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_26 As [CodiciTaglie_26],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_27 AS [CodiciTaglie_27],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_28 As [CodiciTaglie_28],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_29 AS [CodiciTaglie_29],        "
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_30 As [CodiciTaglie_30],        "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf1) As [qtimpconf1],               "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf2) As [qtimpconf2],               "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf3) As [qtimpconf3],               "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf4) As [qtimpconf4],               "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf5) As [qtimpconf5],               "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf6) As [qtimpconf6],               "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf7) As [qtimpconf7],               "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf8) As [qtimpconf8],               "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf9) As [qtimpconf9],               "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf10) As [qtimpconf10],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf11) As [qtimpconf11],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf12) As [qtimpconf12],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf13) As [qtimpconf13],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf14) As [qtimpconf14],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf15) As [qtimpconf15],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf16) As [qtimpconf16],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf17) As [qtimpconf17],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf18) As [qtimpconf18],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf19) As [qtimpconf19],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf20) As [qtimpconf20],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf21) As [qtimpconf21],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf22) As [qtimpconf22],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf23) As [qtimpconf23],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf24) As [qtimpconf24],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf25) As [qtimpconf25],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf26) As [qtimpconf26],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf27) As [qtimpconf27],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf28) As [qtimpconf28],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf29) As [qtimpconf29],             "
        SQL = SQL & " SUM(JSV_IMPE_AVMA_D.qtimpconf30) As [qtimpconf30],             "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm1) As [QtOrdConfUm1],         "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm2) As [QtOrdConfUm2],         "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm3) As [QtOrdConfUm3],         "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm4) As [QtOrdConfUm4],         "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm5) As [QtOrdConfUm5],         "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm6) As [QtOrdConfUm6],         "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm7) As [QtOrdConfUm7],         "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm8) As [QtOrdConfUm8],         "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm9) As [QtOrdConfUm9],         "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm10) As [QtOrdConfUm10],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm11) As [QtOrdConfUm11],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm12) As [QtOrdConfUm12],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm13) As [QtOrdConfUm13],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm14) As [QtOrdConfUm14],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm15) As [QtOrdConfUm15],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm16) As [QtOrdConfUm16],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm17) As [QtOrdConfUm17],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm18) As [QtOrdConfUm18],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm19) As [QtOrdConfUm19],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm20) As [QtOrdConfUm20],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm21) As [QtOrdConfUm21],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm22) As [QtOrdConfUm22],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm23) As [QtOrdConfUm23],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm24) As [QtOrdConfUm24],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm25) As [QtOrdConfUm25],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm26) As [QtOrdConfUm26],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm27) As [QtOrdConfUm27],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm28) As [QtOrdConfUm28],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm29) As [QtOrdConfUm29],       "
        SQL = SQL & " SUM(JSV_IMPORD_AVMA_D.QtOrdConfUm30) As [QtOrdConfUm30],       "

        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm1) As [QtaODCUm1],               "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm2) As [QtaODCUm2],               "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm3) As [QtaODCUm3],               "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm4) As [QtaODCUm4],               "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm5) As [QtaODCUm5],               "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm6) As [QtaODCUm6],               "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm7) As [QtaODCUm7],               "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm8) As [QtaODCUm8],               "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm9) As [QtaODCUm9],               "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm10) As [QtaODCUm10],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm11) As [QtaODCUm11],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm12) As [QtaODCUm12],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm13) As [QtaODCUm13],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm14) As [QtaODCUm14],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm15) As [QtaODCUm15],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm16) As [QtaODCUm16],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm17) As [QtaODCUm17],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm18) As [QtaODCUm18],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm19) As [QtaODCUm19],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm20) As [QtaODCUm20],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm21) As [QtaODCUm21],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm22) As [QtaODCUm22],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm23) As [QtaODCUm23],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm24) As [QtaODCUm24],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm25) As [QtaODCUm25],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm26) As [QtaODCUm26],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm27) As [QtaODCUm27],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm28) As [QtaODCUm28],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm29) As [QtaODCUm29],             "
        SQL = SQL & " sum(JSV_IMPORD_AVMA_D.QtaODCUm30) As [QtaODCUm30]              "



        SQL = SQL & " FROM ESV_IMPORD_AVMA_S As [ESV_IMPORD_AVMA_S]           "
        SQL = SQL & " LEFT OUTER JOIN JSV_IMPE_AVMA_D As [JSV_IMPE_AVMA_D] On (JSV_IMPE_AVMA_D.codArt=ESV_IMPORD_AVMA_S.codArt And JSV_IMPE_AVMA_D.VarianteArt = ESV_IMPORD_AVMA_S.VarianteArt And JSV_IMPE_AVMA_D.CodMag = ESV_IMPORD_AVMA_S.codMag And JSV_IMPE_AVMA_D.DBGruppo=ESV_IMPORD_AVMA_S.DBGruppo)"

        SQL = SQL & " LEFT OUTER JOIN ModaArticoli As [ModaArticoli] On (ModaArticoli.CodiceArticolo=JSV_IMPE_AVMA_D.codArt And (ModaArticoli.DBGruppo=JSV_IMPE_AVMA_D.DBGruppo))"

        SQL = SQL & " LEFT OUTER JOIN JSV_IMPORD_AVMA_D As [JSV_IMPORD_AVMA_D] On (JSV_IMPORD_AVMA_D.CodArt=ESV_IMPORD_AVMA_S.CodArt And JSV_IMPORD_AVMA_D.VarianteArt=ESV_IMPORD_AVMA_S.VarianteArt  And (JSV_IMPORD_AVMA_D.DBGruppo=ESV_IMPORD_AVMA_S.DBGruppo)  and JSV_IMPORD_AVMA_D.CodAreaMag = ESV_IMPORD_AVMA_S.CodAreaMag  and JSV_IMPORD_AVMA_D.CodMag = ESV_IMPORD_AVMA_S.CodMag)"

        SQL = SQL & " LEFT OUTER JOIN ModaTabellaTaglie As [ModaTabellaTaglie] On (ModaTabellaTaglie.CodiceTabellaTaglie=ModaArticoli.CodiceTabellaTaglie And (ModaTabellaTaglie.DBGruppo=ModaArticoli.DBGruppo))"
        SQL = SQL & " LEFT OUTER JOIN ArtConfigVariante As [ArtConfigVariante] On ((ArtConfigVariante.CodArt=JSV_IMPE_AVMA_D.CodArt) And (ArtConfigVariante.VarianteArt=JSV_IMPE_AVMA_D.VarianteArt) And (ArtConfigVariante.DBGruppo=ModaArticoli.DBGruppo))"
        SQL = SQL & " LEFT OUTER JOIN ArtDatiInLingua As [ArtDatiInLingua] On ((ArtDatiInLingua.CodArt=ArtConfigVariante.CodArt) And (ArtDatiInLingua.VarianteArt=ArtConfigVariante.VarianteArt) And ArtDatiInLingua.CodLingua=1 And (ArtDatiInLingua.DbGruppo=ArtConfigVariante.DBGruppo))"

        SQL = SQL & " WHERE (JSV_IMPE_AVMA_D.CodArt = '" & Replace(Trim(codArt), "'", "''") & "')"
        SQL = SQL & " AND (JSV_IMPE_AVMA_D.codmag in " & CODMAGAZZINO & ")"
        SQL = SQL & " AND (JSV_IMPE_AVMA_D.DBGruppo='" & CODGRUPPO & "')   "

        SQL = SQL & " group by ESV_IMPORD_AVMA_S.qtordUm, ESV_IMPORD_AVMA_S.qtimpum, JSV_IMPE_AVMA_D.codArt, JSV_IMPE_AVMA_D.VarianteArt, ArtConfigVariante.descrizione, ArtDatiInLingua.DesEstesa, "
        SQL = SQL & " ModaArticoli.CodiceTabellaTaglie, ModaTabellaTaglie.CodiciTaglie_1, ModaTabellaTaglie.CodiciTaglie_2, ModaTabellaTaglie.CodiciTaglie_3, ModaTabellaTaglie.CodiciTaglie_4, ModaTabellaTaglie.CodiciTaglie_5, ModaTabellaTaglie.CodiciTaglie_6, ModaTabellaTaglie.CodiciTaglie_7, ModaTabellaTaglie.CodiciTaglie_8, ModaTabellaTaglie.CodiciTaglie_9, ModaTabellaTaglie.CodiciTaglie_10,"
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_11, ModaTabellaTaglie.CodiciTaglie_12, ModaTabellaTaglie.CodiciTaglie_13, ModaTabellaTaglie.CodiciTaglie_14, ModaTabellaTaglie.CodiciTaglie_15, ModaTabellaTaglie.CodiciTaglie_16, ModaTabellaTaglie.CodiciTaglie_17, ModaTabellaTaglie.CodiciTaglie_18, ModaTabellaTaglie.CodiciTaglie_19, ModaTabellaTaglie.CodiciTaglie_20, ModaTabellaTaglie.CodiciTaglie_21,"
        SQL = SQL & " ModaTabellaTaglie.CodiciTaglie_22, ModaTabellaTaglie.CodiciTaglie_23, ModaTabellaTaglie.CodiciTaglie_24, ModaTabellaTaglie.CodiciTaglie_25, ModaTabellaTaglie.CodiciTaglie_26, ModaTabellaTaglie.CodiciTaglie_27, ModaTabellaTaglie.CodiciTaglie_28, ModaTabellaTaglie.CodiciTaglie_29, ModaTabellaTaglie.CodiciTaglie_30                                                                        "

        Return SQL
    End Function


    Private Function getSQLOrdinato(codArt As String) As String
        Dim SQL As String

        SQL = SQL & " SELECT       "
        SQL = SQL & "    JSV_IMPORD_AVMA_D.codArt AS [GtgCart],   "
        SQL = SQL & "    JSV_IMPORD_AVMA_D.VarianteArt AS [GtgVarart],    "
        SQL = SQL & "    JSV_IMPORD_AVMA_D.CodMag AS [Gtgcmag],    "
        SQL = SQL & "    ArtConfigVariante.descrizione AS [descrizione],  "
        SQL = SQL & "    ArtDatiInLingua.DesEstesa AS [DescrInglese],               "
        SQL = SQL & "    ModaArticoli.CodiceTabellaTaglie AS [CodiceTabellaTaglie],  "
        SQL = SQL & "    ModaTabellaTaglie.CodiciTaglie_1 AS [CodiciTaglie_1], ModaTabellaTaglie.CodiciTaglie_2 AS [CodiciTaglie_2], ModaTabellaTaglie.CodiciTaglie_3 AS [CodiciTaglie_3], ModaTabellaTaglie.CodiciTaglie_4 AS [CodiciTaglie_4], ModaTabellaTaglie.CodiciTaglie_5 AS [CodiciTaglie_5],"
        SQL = SQL & "    ModaTabellaTaglie.CodiciTaglie_6 AS [CodiciTaglie_6], ModaTabellaTaglie.CodiciTaglie_7 AS [CodiciTaglie_7], ModaTabellaTaglie.CodiciTaglie_8 AS [CodiciTaglie_8], ModaTabellaTaglie.CodiciTaglie_9 AS [CodiciTaglie_9], ModaTabellaTaglie.CodiciTaglie_10 AS [CodiciTaglie_10],"
        SQL = SQL & "    ModaTabellaTaglie.CodiciTaglie_11 AS [CodiciTaglie_11], ModaTabellaTaglie.CodiciTaglie_12 AS [CodiciTaglie_12], ModaTabellaTaglie.CodiciTaglie_13 AS [CodiciTaglie_13], ModaTabellaTaglie.CodiciTaglie_14 AS [CodiciTaglie_14], ModaTabellaTaglie.CodiciTaglie_15 AS [CodiciTaglie_15],"
        SQL = SQL & "    ModaTabellaTaglie.CodiciTaglie_16 AS [CodiciTaglie_16], ModaTabellaTaglie.CodiciTaglie_17 AS [CodiciTaglie_17], ModaTabellaTaglie.CodiciTaglie_18 AS [CodiciTaglie_18], ModaTabellaTaglie.CodiciTaglie_19 AS [CodiciTaglie_19], ModaTabellaTaglie.CodiciTaglie_20 AS [CodiciTaglie_20], "
        SQL = SQL & "    ModaTabellaTaglie.CodiciTaglie_21 AS [CodiciTaglie_21], ModaTabellaTaglie.CodiciTaglie_22 AS [CodiciTaglie_22], ModaTabellaTaglie.CodiciTaglie_23 AS [CodiciTaglie_23], ModaTabellaTaglie.CodiciTaglie_24 AS [CodiciTaglie_24], ModaTabellaTaglie.CodiciTaglie_25 AS [CodiciTaglie_25],  "
        SQL = SQL & "    ModaTabellaTaglie.CodiciTaglie_26 AS [CodiciTaglie_26], ModaTabellaTaglie.CodiciTaglie_27 AS [CodiciTaglie_27], ModaTabellaTaglie.CodiciTaglie_28 AS [CodiciTaglie_28], ModaTabellaTaglie.CodiciTaglie_29 AS [CodiciTaglie_29], ModaTabellaTaglie.CodiciTaglie_30 AS [CodiciTaglie_30],   "
        SQL = SQL & "    JSV_IMPORD_AVMA_D.QtOrdConfUm1 AS [QtOrdConfUm1], JSV_IMPORD_AVMA_D.QtOrdConfUm2 AS [QtOrdConfUm2], JSV_IMPORD_AVMA_D.QtOrdConfUm3 AS [QtOrdConfUm3], JSV_IMPORD_AVMA_D.QtOrdConfUm4 AS [QtOrdConfUm4], JSV_IMPORD_AVMA_D.QtOrdConfUm5 AS [QtOrdConfUm5],   "
        SQL = SQL & "    JSV_IMPORD_AVMA_D.QtOrdConfUm6 AS [QtOrdConfUm6], JSV_IMPORD_AVMA_D.QtOrdConfUm7 AS [QtOrdConfUm7], JSV_IMPORD_AVMA_D.QtOrdConfUm8 AS [QtOrdConfUm8], JSV_IMPORD_AVMA_D.QtOrdConfUm9 AS [QtOrdConfUm9], JSV_IMPORD_AVMA_D.QtOrdConfUm10 AS [QtOrdConfUm10],  "
        SQL = SQL & "    JSV_IMPORD_AVMA_D.QtOrdConfUm11 AS [QtOrdConfUm11], JSV_IMPORD_AVMA_D.QtOrdConfUm12 AS [QtOrdConfUm12], JSV_IMPORD_AVMA_D.QtOrdConfUm13 AS [QtOrdConfUm13], JSV_IMPORD_AVMA_D.QtOrdConfUm14 AS [QtOrdConfUm14], JSV_IMPORD_AVMA_D.QtOrdConfUm15 AS [QtOrdConfUm15],"
        SQL = SQL & "    JSV_IMPORD_AVMA_D.QtOrdConfUm16 AS [QtOrdConfUm16], JSV_IMPORD_AVMA_D.QtOrdConfUm17 AS [QtOrdConfUm17], JSV_IMPORD_AVMA_D.QtOrdConfUm18 AS [QtOrdConfUm18], JSV_IMPORD_AVMA_D.QtOrdConfUm19 AS [QtOrdConfUm19], JSV_IMPORD_AVMA_D.QtOrdConfUm20 AS [QtOrdConfUm20], "
        SQL = SQL & "    JSV_IMPORD_AVMA_D.QtOrdConfUm21 AS [QtOrdConfUm21], JSV_IMPORD_AVMA_D.QtOrdConfUm22 AS [QtOrdConfUm22], JSV_IMPORD_AVMA_D.QtOrdConfUm23 AS [QtOrdConfUm23], JSV_IMPORD_AVMA_D.QtOrdConfUm24 AS [QtOrdConfUm24], JSV_IMPORD_AVMA_D.QtOrdConfUm25 AS [QtOrdConfUm25],  "
        SQL = SQL & "    JSV_IMPORD_AVMA_D.QtOrdConfUm26 AS [QtOrdConfUm26], JSV_IMPORD_AVMA_D.QtOrdConfUm27 AS [QtOrdConfUm27], JSV_IMPORD_AVMA_D.QtOrdConfUm28 AS [QtOrdConfUm28], JSV_IMPORD_AVMA_D.QtOrdConfUm29 AS [QtOrdConfUm29], JSV_IMPORD_AVMA_D.QtOrdConfUm30 AS [QtOrdConfUm30]    "
        SQL = SQL & "    FROM ESV_IMPORD_AVMA_S AS [ESV_IMPORD_AVMA_S]"
        SQL = SQL & "    LEFT OUTER JOIN JSV_IMPORD_AVMA_D AS [JSV_IMPORD_AVMA_D] ON (JSV_IMPORD_AVMA_D.codArt=ESV_IMPORD_AVMA_S.codArt AND JSV_IMPORD_AVMA_D.VarianteArt = ESV_IMPORD_AVMA_S.VarianteArt AND JSV_IMPORD_AVMA_D.CodMag = ESV_IMPORD_AVMA_S.codMag AND JSV_IMPORD_AVMA_D.DBGruppo=ESV_IMPORD_AVMA_S.DBGruppo)"
        SQL = SQL & "    LEFT OUTER JOIN ModaArticoli AS [ModaArticoli] ON (ModaArticoli.CodiceArticolo=JSV_IMPORD_AVMA_D.codArt AND (ModaArticoli.DBGruppo=JSV_IMPORD_AVMA_D.DBGruppo))"
        SQL = SQL & "    LEFT OUTER JOIN ModaTabellaTaglie AS [ModaTabellaTaglie] ON (ModaTabellaTaglie.CodiceTabellaTaglie=ModaArticoli.CodiceTabellaTaglie AND (ModaTabellaTaglie.DBGruppo=ModaArticoli.DBGruppo))"
        SQL = SQL & "    LEFT OUTER JOIN ArtConfigVariante AS [ArtConfigVariante] ON ((ArtConfigVariante.CodArt=JSV_IMPORD_AVMA_D.CodArt) AND (ArtConfigVariante.VarianteArt=JSV_IMPORD_AVMA_D.VarianteArt) AND (ArtConfigVariante.DBGruppo=ModaArticoli.DBGruppo))"
        SQL = SQL & "    LEFT OUTER JOIN ArtDatiInLingua As [ArtDatiInLingua] On ((ArtDatiInLingua.CodArt=ArtConfigVariante.CodArt) And (ArtDatiInLingua.VarianteArt=ArtConfigVariante.VarianteArt) And ArtDatiInLingua.CodLingua=1 And (ArtDatiInLingua.DbGruppo=ArtConfigVariante.DBGruppo))"
        SQL = SQL & "    WHERE (JSV_IMPORD_AVMA_D.CodArt = '" & Replace(Trim(codArt), "'", "''") & "')"
        SQL = SQL & "    AND (JSV_IMPORD_AVMA_D.codmag in " & CODMAGAZZINO & ")"
        SQL = SQL & "    AND (JSV_IMPORD_AVMA_D.DBGruppo='" & CODGRUPPO & "')"


        Return SQL
    End Function

    Private Sub compilaStruttura(rs As ADODB.Recordset, isImpeganto As Boolean, isOrdinato As Boolean)
        Dim val1 As Integer
        Dim val2 As Integer
        Dim val3 As Integer
        Dim val4 As Integer
        Dim val5 As Integer

        Dim Assegnato(30) As Integer


        Dim mysize As Integer
        Dim mysize2 As Integer
        Dim notaglie As Boolean = False

        Dim totCalc As Integer = 0
        Dim totgiac As Integer = 0
        Dim contaTaglie As Integer = 1



        mysize = 0
        If Not varTag Is Nothing Then
            mysize = UBound(varTag)
        End If
        ReDim Preserve varTag(mysize + 1)


        If rs.RecordCount = 0 Then
            notaglie = True
        End If

        If notaglie = False Then

            If IsDBNull(rs.Fields("CodiceTabellaTaglie").Value) Or (Trim(rs.Fields("CodiceTabellaTaglie").Value.ToString) = "") Then
                notaglie = True
            End If
        End If

        ReDim varTag(UBound(varTag) - 1).CodVariante(0)
        ReDim varTag(UBound(varTag) - 1).Varianti(0)
        ReDim varTag(UBound(varTag) - 1).Giacenze(0)
        ReDim varTag(UBound(varTag) - 1).TotaleGiac(0)
        ReDim varTag(UBound(varTag) - 1).TotaleCalc(0)




        varTag(UBound(varTag) - 1).CodArt = rs.Fields("GtgCart").Value
        varTag(UBound(varTag) - 1).UM = currentUM
        varTag(UBound(varTag) - 1).DescEstesa = currentDescEstesa


        If notaglie = False Then
            varTag(UBound(varTag) - 1).haTaglie = True
            mysize2 = 0

            While Not rs.EOF

                ' compila la struttura per i filtri al primo giro
                If PrimoGiro = True Then
                    Call AddVarianteInList(rs.Fields("GtgVarart").Value)
                End If

                ' controlla se ignorare la variante
                If checkVisibilitaVariante(rs.Fields("GtgVarart").Value) = False Then
                    rs.MoveNext()
                    Application.DoEvents()
                    Continue While
                End If


                contaTaglie = 1
                totgiac = 0
                totCalc = 0
                ReDim Preserve varTag(UBound(varTag) - 1).CodVariante(mysize2 + 1)
                ReDim Preserve varTag(UBound(varTag) - 1).Varianti(mysize2 + 1)
                ReDim Preserve varTag(UBound(varTag) - 1).Giacenze(mysize2 + 1)
                ReDim Preserve varTag(UBound(varTag) - 1).TotaleGiac(mysize2 + 1)
                ReDim Preserve varTag(UBound(varTag) - 1).TotaleCalc(mysize2 + 1)

                'MsgBox("TEST")
                ' compila l'assegnato
                Call getAssegnatoTaglie(rs.Fields("GtgCart").Value, rs.Fields("GtgVarart").Value, Assegnato)
                'MsgBox("TEST POST-ASS")

                For idx = 1 To NUM_TAGLIE
                    If Not IsDBNull(rs.Fields("CodiciTaglie_" & idx).Value) Then
                        If rs.Fields("CodiciTaglie_" & idx).Value <> "" Then

                            'MsgBox("TEST DENTRO IF")

                            ' compila la struttura per i filtri al primo giro
                            If PrimoGiro = True Then
                                Call AddTagliaInList(rs.Fields("CodiciTaglie_" & idx).Value)
                            End If

                            'MsgBox(" -- TEST -- ")


                            ReDim Preserve varTag(UBound(varTag) - 1).Giacenze(mysize2).Taglie(contaTaglie)
                            ReDim Preserve varTag(UBound(varTag) - 1).Giacenze(mysize2).Giacenze(contaTaglie)
                            ReDim Preserve varTag(UBound(varTag) - 1).Giacenze(mysize2).Calcolato(contaTaglie)
                            '#MATTIA
                            ReDim Preserve varTag(UBound(varTag) - 1).Giacenze(mysize2).DispTeorica(contaTaglie)



                            val1 = 0    'giacenza
                            If isImpeganto = False And isOrdinato = False Then
                                If IsDBNull(rs.Fields("GtgQgp" & idx).Value) = False Then
                                    val1 = rs.Fields("GtgQgp" & idx).Value
                                End If
                            End If

                            ' QTA PRELEVATA (-)
                            val2 = Assegnato(idx)

                            val3 = 0    'impegnato confermato
                            If isOrdinato = False Then
                                If Not IsDBNull(rs.Fields("qtimpconf" & idx).Value) Then
                                    val3 = rs.Fields("qtimpconf" & idx).Value
                                End If
                            End If

                            val4 = 0
                            If Not IsDBNull(rs.Fields("QtaODCUm" & idx).Value) Then
                                val4 = rs.Fields("QtaODCUm" & idx).Value
                            End If

                            val5 = 0 'ordinato confermato
                            If Not IsDBNull(rs.Fields("QtOrdConfUm" & idx).Value) Then
                                val5 = rs.Fields("QtOrdConfUm" & idx).Value
                            End If

                            'MsgBox(" -- TEST 2 -- ")

                            varTag(UBound(varTag) - 1).CodVariante(mysize2) = rs.Fields("GtgVarart").Value

                            If FormExport.CheckBoxLingua.Checked = True Then
                                varTag(UBound(varTag) - 1).Varianti(mysize2) = rs.Fields("GtgVarart").Value & " - " & rs.Fields("DescrInglese").Value
                            Else
                                varTag(UBound(varTag) - 1).Varianti(mysize2) = rs.Fields("GtgVarart").Value & " - " & rs.Fields("descrizione").Value
                            End If

                            'MsgBox(" -- TEST 3 -- ")

                            ' codice taglia
                            varTag(UBound(varTag) - 1).Giacenze(mysize2).Taglie(contaTaglie - 1) = rs.Fields("CodiciTaglie_" & idx).Value

                            If FLAG_ORDINATO = "1" Then
                                '    ' CALCOLO DELLA GIACENZA MAGAZZINO = GIACENZA
                                varTag(UBound(varTag) - 1).Giacenze(mysize2).Giacenze(contaTaglie - 1) = val1 ' - val2
                                totgiac = totgiac + varTag(UBound(varTag) - 1).Giacenze(mysize2).Giacenze(contaTaglie - 1)

                                '    'calcolato
                                varTag(UBound(varTag) - 1).Giacenze(mysize2).Calcolato(contaTaglie - 1) = val1 + val5 - val3

                            Else

                                '    ' giacenza
                                varTag(UBound(varTag) - 1).Giacenze(mysize2).Giacenze(contaTaglie - 1) = val1
                                    totgiac = totgiac + varTag(UBound(varTag) - 1).Giacenze(mysize2).Giacenze(contaTaglie - 1)

                                    '    'calcolato
                                    varTag(UBound(varTag) - 1).Giacenze(mysize2).Calcolato(contaTaglie - 1) = val1 + val5 - val3
                                End If

                                totCalc = totCalc + varTag(UBound(varTag) - 1).Giacenze(mysize2).Calcolato(contaTaglie - 1)

                                contaTaglie = contaTaglie + 1
                            End If
                        End If


                Next
                varTag(UBound(varTag) - 1).TotaleCalc(mysize2) = totCalc
                varTag(UBound(varTag) - 1).TotaleGiac(mysize2) = totgiac



                mysize2 = mysize2 + 1
                rs.MoveNext()
                Application.DoEvents()
            End While
        Else


            ' articolo non a taglie
            varTag(UBound(varTag) - 1).haTaglie = False
            mysize2 = 0
            While Not rs.EOF
                ' compila la struttura per i filtri al primo giro
                If PrimoGiro = True Then
                    Call AddVarianteInList(rs.Fields("GtgVarart").Value)
                End If

                ' controlla se ignorare la variante
                If checkVisibilitaVariante(rs.Fields("GtgVarart").Value) = False Then
                    rs.MoveNext()
                    Application.DoEvents()
                    Continue While
                End If

                totgiac = 0
                totCalc = 0

                ReDim Preserve varTag(UBound(varTag) - 1).CodVariante(mysize2 + 1)
                ReDim Preserve varTag(UBound(varTag) - 1).Varianti(mysize2 + 1)
                ReDim Preserve varTag(UBound(varTag) - 1).Giacenze(mysize2 + 1)
                ReDim Preserve varTag(UBound(varTag) - 1).TotaleGiac(mysize2 + 1)
                ReDim Preserve varTag(UBound(varTag) - 1).TotaleCalc(mysize2 + 1)

                varTag(UBound(varTag) - 1).CodVariante(mysize2) = rs.Fields("GtgVarart").Value
                varTag(UBound(varTag) - 1).Varianti(mysize2) = rs.Fields("GtgVarart").Value & " - " & rs.Fields("DescrInglese").Value

                val1 = 0
                If isImpeganto = False And isOrdinato = False Then
                    If IsDBNull(rs.Fields("qtaGiacUMMag").Value) = False Then
                        val1 = rs.Fields("qtaGiacUMMag").Value
                    End If
                End If

                val2 = 0
                If Not IsDBNull(rs.Fields("qtOrdUM").Value) Then
                    val2 = rs.Fields("qtOrdUM").Value
                End If

                val3 = 0
                If isOrdinato = False Then
                    If Not IsDBNull(rs.Fields("qtimpUm").Value) Then
                        val3 = rs.Fields("qtimpUm").Value
                    End If
                End If

                'val4 = 0
                'If isOrdinato = False Then
                '    If Not IsDBNull(rs.Fields("qtOrdDaConfUM").Value) Then
                '        val3 = rs.Fields("qtOrdDaConfUM").Value
                '    End If
                'End If

                'val5 = 0 'ordinato confermato
                'If Not IsDBNull(rs.Fields("QtOrdConfUm").Value) Then
                '    val5 = rs.Fields("QtOrdConfUm").Value
                'End If


                If FLAG_ORDINATO = "1" Then
                    varTag(UBound(varTag) - 1).TotaleGiac(mysize2) = val1 - val3
                    varTag(UBound(varTag) - 1).TotaleCalc(mysize2) = val1 + val2 - val3
                Else
                    varTag(UBound(varTag) - 1).TotaleGiac(mysize2) = val1
                    varTag(UBound(varTag) - 1).TotaleCalc(mysize2) = val1 + val2 - val3
                End If



                rs.MoveNext()
                Application.DoEvents()
            End While
        End If



    End Sub

    Private Function PopolaStrutturaVariantiTaglie(codArt As String) As Boolean
        Dim SQL As String
        Dim rs As New ADODB.Recordset
        Dim idxrow As Integer = 0
        Dim compilato As Boolean = False


        ' torna l'SQL delle giacenze
        SQL = getSQLGiacenze(codArt)

        Try
            rs.Open(SQL, connSqlSrv, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(SQL & "    ---SQLGIACENZE IN POPOLASTRUTTURAVARIANTITAGLIE")
        End Try

        'MsgBox("1.." & rs.RecordCount)
        If rs.RecordCount > 0 Then
            compilaStruttura(rs, False, False)
            compilato = True
        End If
        rs.Close()

        If compilato = False Then

            SQL = getSQLImpegnato(codArt)
            Try
                rs.Open(SQL, connSqlSrv, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                'Clipboard.SetText(SQL)
            Catch ex As Exception
                MsgBox(ex.Message)
                MsgBox(SQL & "    ---SQLIMPEGNATO IN POPOLASTRUTTURAVARIANTITAGLIE")
            End Try


            'MsgBox("2.." & rs.RecordCount)

            If rs.RecordCount > 0 Then
                compilaStruttura(rs, True, False)
                compilato = True
            End If

            rs.Close()
        End If

        If compilato = False Then

            SQL = getSQLOrdinato(codArt)
            Try
                rs.Open(SQL, connSqlSrv, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Catch ex As Exception
                MsgBox(ex.Message)
                MsgBox(SQL & "    ---SQLORDINATO IN POPOLASTRUTTURAVARIANTITAGLIE")
            End Try



            'MsgBox("3.." & rs.RecordCount)
            If rs.RecordCount > 0 Then
                compilaStruttura(rs, True, True)
                compilato = True
            End If

            rs.Close()
        End If


        ' controlla che se a causa dei filtri non ha popolato varianti non mostra l'articolo
        If compilato = True Then
            If UBound(varTag(UBound(varTag) - 1).Giacenze) = 0 Then
                compilato = False
            End If
        End If


        Return compilato




    End Function



    ' popola la griglia details giacenze
    Private Sub PopolaGrigliaVariantiTaglie(myGrid As DataGridView)
        Dim idxTaglia As Integer
        Dim idxVariante As Integer
        Dim idxArticolo As Integer = DataGridViewArticoli.CurrentCell.RowIndex
        Dim idxVar1 As Integer
        Dim idxVar2 As Integer
        Dim i As Integer


        For i = 0 To UBound(varTag) - 1
            If varTag(i).CodArt = DataGridViewArticoli.Rows(DataGridViewArticoli.CurrentCell.RowIndex).Cells(6).Value Then
                idxArticolo = i
                Exit For
            End If
        Next



        If loaded = False Then
            Exit Sub
        End If

        myGrid.Columns.Clear()
        myGrid.Rows.Clear()

        If varTag Is Nothing Then
            Exit Sub
        End If


        'popola la griglia se l'articolo ha taglie
        If varTag(idxArticolo).haTaglie = True Then
            If varTag(idxArticolo).Giacenze Is Nothing Then
                Exit Sub
            End If




            myGrid.Columns.Add("Variante", "Variante")
            myGrid.Columns(0).ReadOnly = True
            myGrid.Columns(0).Width = 150
            myGrid.Columns(0).DefaultCellStyle.Font = New Font("Microsoft Sans Serif", 7, FontStyle.Bold)
            myGrid.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

            myGrid.Columns.Add("Tipologia", "")
            myGrid.Columns(1).ReadOnly = True
            myGrid.Columns(1).Width = 75
            myGrid.Columns(1).DefaultCellStyle.Font = New Font("Microsoft Sans Serif", 7, FontStyle.Bold)
            myGrid.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable

            myGrid.Columns.Add("UM", "UM")
            myGrid.Columns(2).ReadOnly = True
            myGrid.Columns(2).Width = 40
            myGrid.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable


            myGrid.Columns.Add("Totale", "Totale")
            myGrid.Columns(3).ReadOnly = True
            myGrid.Columns(3).Width = 40
            myGrid.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable

            Try
                For idxTaglia = 0 To UBound(varTag(idxArticolo).Giacenze(0).Taglie) - 1
                    myGrid.Columns.Add(varTag(idxArticolo).Giacenze(0).Taglie(idxTaglia), varTag(idxArticolo).Giacenze(0).Taglie(idxTaglia))
                    myGrid.Columns(idxTaglia + idxStartTaglie).Width = 40
                    myGrid.Columns(idxTaglia + idxStartTaglie).SortMode = DataGridViewColumnSortMode.NotSortable
                Next

            Catch ex As Exception
                Exit Sub
            End Try


            For idxVariante = 0 To UBound(varTag(idxArticolo).Varianti) - 1

                ' si popola a riga si riga no

                idxVar1 = idxVariante * 2 ' indice giacenze
                idxVar2 = (idxVariante * 2) + 1 ' indice calcolato

                ' giacenze  --> giacenza scaffale
                myGrid.Rows.Add()
                myGrid.Rows(idxVar1).Cells(0).Style.BackColor = colore_disabilitato
                myGrid.Rows(idxVar1).Cells(0).Value = varTag(idxArticolo).Varianti(idxVariante)

                myGrid.Rows(idxVar1).Cells(1).Style.BackColor = colore_disabilitato
                myGrid.Rows(idxVar1).Cells(1).Value = "Giacenza magazzino"

                myGrid.Rows(idxVar1).Cells(2).Style.BackColor = colore_disabilitato
                myGrid.Rows(idxVar1).Cells(2).Value = varTag(idxArticolo).UM

                myGrid.Rows(idxVar1).Cells(3).Style.BackColor = colore_disabilitato
                myGrid.Rows(idxVar1).Cells(3).Value = varTag(idxArticolo).TotaleGiac(idxVariante)


                For idxTaglia = 0 To UBound(varTag(idxArticolo).Giacenze(0).Giacenze) - 1
                    myGrid.Rows(idxVar1).Cells(idxTaglia + idxStartTaglie).Value = varTag(idxArticolo).Giacenze(idxVariante).Giacenze(idxTaglia)
                Next


                ' calcolato  --> disponibilità teorica
                myGrid.Rows.Add()
                myGrid.Rows(idxVar2).Cells(0).Style.BackColor = colore_disabilitato

                myGrid.Rows(idxVar2).Cells(1).Style.BackColor = colore_disabilitato
                myGrid.Rows(idxVar2).Cells(1).Value = "Disp. Teorica"

                myGrid.Rows(idxVar2).Cells(2).Style.BackColor = colore_disabilitato
                myGrid.Rows(idxVar2).Cells(2).Value = varTag(idxArticolo).UM


                myGrid.Rows(idxVar2).Cells(3).Style.BackColor = colore_disabilitato
                myGrid.Rows(idxVar2).Cells(3).Value = varTag(idxArticolo).TotaleCalc(idxVariante)



                For idxTaglia = 0 To UBound(varTag(idxArticolo).Giacenze(0).Giacenze) - 1
                    myGrid.Rows(idxVar2).Cells(idxTaglia + idxStartTaglie).Value = varTag(idxArticolo).Giacenze(idxVariante).Calcolato(idxTaglia)
                    myGrid.Rows(idxVar2).Cells(idxTaglia + idxStartTaglie).Style.BackColor = colore_lettura_facilitata
                Next


            Next

        Else
            myGrid.Columns.Add("Variante", "Variante")
            myGrid.Columns(0).ReadOnly = True
            myGrid.Columns(0).Width = 150
            myGrid.Columns(0).DefaultCellStyle.Font = New Font("Microsoft Sans Serif", 7, FontStyle.Bold)
            myGrid.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable


            myGrid.Columns.Add("Tipologia", "")
            myGrid.Columns(1).ReadOnly = True
            myGrid.Columns(1).Width = 75
            myGrid.Columns(1).DefaultCellStyle.Font = New Font("Microsoft Sans Serif", 7, FontStyle.Bold)
            myGrid.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable

            myGrid.Columns.Add("UM", "UM")
            myGrid.Columns(2).ReadOnly = True
            myGrid.Columns(2).Width = 40
            myGrid.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable


            myGrid.Columns.Add("Totale", "Totale")
            myGrid.Columns(3).ReadOnly = False
            myGrid.Columns(3).Width = 40
            myGrid.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable

            If Not varTag(idxArticolo).Varianti Is Nothing Then
                For idxVariante = 0 To UBound(varTag(idxArticolo).Varianti) - 1

                    ' si popola a riga si riga no

                    idxVar1 = idxVariante * 2 ' indice giacenze
                    idxVar2 = (idxVariante * 2) + 1 ' indice calcolato


                    ' giacenze
                    myGrid.Rows.Add()
                    myGrid.Rows(idxVar1).Cells(0).Style.BackColor = colore_disabilitato
                    myGrid.Rows(idxVar1).Cells(0).Value = varTag(idxArticolo).Varianti(idxVariante)

                    myGrid.Rows(idxVar1).Cells(1).Style.BackColor = colore_disabilitato
                    myGrid.Rows(idxVar1).Cells(1).Value = "Giacenza magazzino"

                    myGrid.Rows(idxVar1).Cells(2).Style.BackColor = colore_disabilitato
                    myGrid.Rows(idxVar1).Cells(2).Value = varTag(idxArticolo).UM

                    myGrid.Rows(idxVar1).Cells(3).Value = varTag(idxArticolo).TotaleGiac(idxVariante)
                    myGrid.Rows(idxVar1).Cells(3).Style.BackColor = Color.White



                    ' calcolato
                    myGrid.Rows.Add()
                    myGrid.Rows(idxVar2).Cells(0).Style.BackColor = colore_disabilitato

                    myGrid.Rows(idxVar2).Cells(1).Style.BackColor = colore_disabilitato

                    myGrid.Rows(idxVar2).Cells(1).Value = "Disp. Teorica"

                    myGrid.Rows(idxVar2).Cells(2).Style.BackColor = colore_disabilitato
                    myGrid.Rows(idxVar2).Cells(2).Value = varTag(idxArticolo).UM

                    myGrid.Rows(idxVar2).Cells(3).Style.BackColor = Color.White
                    myGrid.Rows(idxVar2).Cells(3).Value = varTag(idxArticolo).TotaleCalc(idxVariante)


                Next
            End If




        End If



    End Sub




    Private Sub DataGridViewGiacenze_CurrentCellChanged(sender As System.Object, e As System.EventArgs) Handles DataGridViewArticoli.CurrentCellChanged
        ' ripopola la griglia in base alla riga selezionata
        If Not DataGridViewArticoli.CurrentCell Is Nothing Then
            If CurrentRowIdx <> DataGridViewArticoli.CurrentCell.RowIndex Then
                Call PopolaGrigliaVariantiTaglie(DataGridViewGiacenze)
                CurrentRowIdx = DataGridViewArticoli.CurrentCell.RowIndex
            End If

        End If
    End Sub

    Private Sub FormMain_Shown(sender As System.Object, e As System.EventArgs) Handles MyBase.Shown

        Call PopolaGriglia(True)

    End Sub




    Private Sub DataGridViewGiacenze_CellEndEdit(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewGiacenze.CellEndEdit
        Dim idxRowVariante As Integer
        Dim idxColumVariante As Integer
        Dim idxArticolo As Integer
        Dim valAsInt As Integer

        If Not IsNumeric(DataGridViewGiacenze.CurrentCell.Value) Then
            MsgBox("inserire un valore numerico")
            DataGridViewGiacenze.CurrentCell.Value = oldCellvalue

        End If


        ' se non è cambiato nulla non faccio niente
        If oldCellvalue = DataGridViewGiacenze.CurrentCell.Value Then
            Exit Sub
        End If



        If Not Integer.TryParse(DataGridViewGiacenze.CurrentCell.Value, valAsInt) Then
            MsgBox("inserire un valore numerico senza la virgola")
            DataGridViewGiacenze.CurrentCell.Value = oldCellvalue
        End If


        idxArticolo = DataGridViewArticoli.CurrentCell.RowIndex

        idxRowVariante = DataGridViewGiacenze.CurrentCell.RowIndex \ 2
        idxColumVariante = DataGridViewGiacenze.CurrentCell.ColumnIndex

        ' controlla se aggiornare il calcolato o la giacenza
        If (DataGridViewGiacenze.CurrentCell.RowIndex Mod 2) = 0 Then
            'giacenza
            If varTag(idxArticolo).haTaglie Then
                ' x taglia
                varTag(idxArticolo).Giacenze(idxRowVariante).Giacenze(idxColumVariante - idxStartTaglie) = DataGridViewGiacenze.CurrentCell.Value
                RicalcalcolaTotaleGiacenze(idxArticolo, idxRowVariante)
                DataGridViewGiacenze.Rows(DataGridViewGiacenze.CurrentCell.RowIndex).Cells(idxStartTaglie - 1).Value = varTag(idxArticolo).TotaleGiac(idxRowVariante)
            Else
                ' totale
                varTag(idxArticolo).TotaleGiac(idxRowVariante) = DataGridViewGiacenze.CurrentCell.Value
            End If
        Else
            'calcolato
            If varTag(idxArticolo).haTaglie Then
                ' x taglie
                varTag(idxArticolo).Giacenze(idxRowVariante).Calcolato(idxColumVariante - idxStartTaglie) = DataGridViewGiacenze.CurrentCell.Value
                RicalcalcolaTotaleCalcolato(idxArticolo, idxRowVariante)
                DataGridViewGiacenze.Rows(DataGridViewGiacenze.CurrentCell.RowIndex).Cells(idxStartTaglie - 1).Value = varTag(idxArticolo).TotaleCalc(idxRowVariante)
            Else
                ' totale
                varTag(idxArticolo).TotaleCalc(idxRowVariante) = DataGridViewGiacenze.CurrentCell.Value
            End If

        End If

    End Sub

    Private Sub RicalcalcolaTotaleGiacenze(idxArticolo As Integer, idxVariante As Integer)
        Dim idx As Integer
        Dim tot As Integer = 0
        For idx = 0 To UBound(varTag(idxArticolo).Giacenze(idxVariante).Giacenze) - 1
            tot = tot + varTag(idxArticolo).Giacenze(idxVariante).Giacenze(idx)
        Next
        varTag(idxArticolo).TotaleGiac(idxVariante) = tot

    End Sub

    Private Sub RicalcalcolaTotaleCalcolato(idxArticolo As Integer, idxVariante As Integer)
        Dim idx As Integer
        Dim tot As Integer = 0
        For idx = 0 To UBound(varTag(idxArticolo).Giacenze(idxVariante).Calcolato) - 1
            tot = tot + varTag(idxArticolo).Giacenze(idxVariante).Calcolato(idx)
        Next
        varTag(idxArticolo).TotaleCalc(idxVariante) = tot

    End Sub


    Private Sub DataGridViewGiacenze_CellLeave(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridViewGiacenze.CellLeave
        ' salva il valore del dato della current cella prima della modifica per ripristinarlo in caso di errore
        oldCellvalue = DataGridViewGiacenze.CurrentCell.Value
    End Sub

    Private Sub ButtonExportHTML_Click(sender As System.Object, e As System.EventArgs) Handles ButtonExportHTML.Click
        If FormExport.ShowDialog = Windows.Forms.DialogResult.OK Then
            Call exportHTML(FormExport.CheckBoxGiac.Checked, FormExport.CheckBoxCalcolato.Checked, FormExport.CheckBoxImgArticolo.Checked, FormExport.TextBoxGiacenza.Text, FormExport.TextBoxDispTeorica.Text)
        End If

    End Sub


    ' scrive 
    Public Sub AppendHTML(filename As String, ByRef strExp As String)

        Dim fileExists As Boolean = File.Exists(filename)
        File.AppendAllText(filename, strExp & vbCrLf)
        strExp = ""
    End Sub



    ' esporta l'html delle giacenze
    Private Sub exportHTML(showGiacenze As Boolean, showCalcolato As Boolean, showImage As Boolean, LabelGiac As String, labelCalc As String)
        Dim indexcont As Integer = 0
        Dim stroutput As String
        Dim idxArt As Integer = 0
        Dim idxVariante As Integer
        Dim idxTaglia As Integer
        Dim html As String = ""
        Dim strPathTmp As String = "C:\DatasoftTmp\" 'Environment.GetFolderPath(Environment.SpecialFolder.Windows) & 
        Dim tmpHTML As String
        stroutput = ""
        Dim srcFilename As String
        Dim exportGiacZero As Boolean
        Dim exportGiacMinZero As Boolean
        Dim bEsportaVariante As Boolean = False
        Dim saveHtml As String
        Dim bEsportaArt As Boolean = False
        Dim saveHtmlArt As String
        Dim bTaglieSempre As Boolean
        Dim bShowImgVariante As Boolean
        Dim indexPrimaRigaTaglia As Integer = 0
        Dim strOutVal As String = ""

        Dim escludiVar() As String
        Dim escludiTg() As String
        Dim idxArticolo As Integer = 1


        strPathTmp = tmp_dir


        bShowImgVariante = FormExport.CheckBoxImgVariante.Checked

        escludiVar = Split(FormExport.TextBoxEscludiVarianti.Text, ",")
        escludiTg = Split(FormExport.TextBoxEscludiTaglie.Text, ",")

        exportGiacZero = FormExport.CheckBoxTotaliZero.Checked
        bTaglieSempre = FormExport.CheckBoxTaglie.Checked
        exportGiacMinZero = FormExport.CheckBoxTotaliMinZero.Checked

        If Not System.IO.Directory.Exists(strPathTmp) Then
            System.IO.Directory.CreateDirectory(strPathTmp)
        End If

        tmpHTML = strPathTmp & "tmpGiac.html"

        SaveFileDialog1.Filter = "JPeg Image|*.html"
        SaveFileDialog1.Title = "Salav un file HTML"
        SaveFileDialog1.ShowDialog()

        If SaveFileDialog1.FileName <> "" Then
            If File.Exists(tmpHTML) Then
                Kill(tmpHTML)
            End If

            LabelDB.Text = "Esportazione HTML in corso ...."


            html = "<table border=1 cellspacing=1 cellpadding=0 width=""100%"">" + vbNewLine

            ' data ora esportazione
            html += vbTab + TAG_TABLE_VUOTO + vbNewLine
            html += vbTab + TAG_ROW + vbNewLine
            html += vbTab + TAG_CELL_DATAORA + "Data ora esportazione:" + TAG_CELL_END + vbNewLine
            html += vbTab + TAG_CELL_DATAORA + Now + TAG_CELL_END + vbNewLine
            html += vbTab + TAG_ROW_END + vbNewLine
            html += vbTab + TAG_ROW + vbNewLine
            html += vbTab + TAG_CELL_DATAORA + "Codici Magazzini: " + TAG_CELL_END + vbNewLine
            html += vbTab + TAG_CELL_DATAORA + LISTAMAGAZZINI + TAG_CELL_END + vbNewLine
            html += vbTab + TAG_TABLE_END + vbNewLine

            AppendHTML(tmpHTML, html)

            'Try


            ProgressBarArticoli.Minimum = 0
            ProgressBarArticoli.Maximum = DataGridViewArticoli.Rows.Count
            ProgressBarArticoli.Value = 0
            For Each dr As DataGridViewRow In DataGridViewArticoli.Rows

                If dr.Visible = False Then
                    Continue For
                End If

                For i = 0 To UBound(varTag) - 1
                    If varTag(i).CodArt = dr.Cells(6).Value Then
                        idxArt = i
                        Exit For
                    End If
                Next

                ProgressBarArticoli.Value = ProgressBarArticoli.Value + 1
                indexPrimaRigaTaglia = 0

                saveHtmlArt = html
                bEsportaArt = False

                ' riga articolo
                html += vbTab + TAG_ROW + vbNewLine


                '--------------------
                ' articolo
                'html += vbTab + TAG_TABLE_VUOTO + vbNewLine
                html += vbTab + TAG_TABLE_FIXED + vbNewLine
                html += vbTab + "<col width=""100px"" />" + vbNewLine
                html += vbTab + "<col width=""600px"" />" + vbNewLine
                html += vbTab + "<col width=""100px"" />" + vbNewLine
                html += vbTab + TAG_ROW + vbNewLine
                For Each cel As DataGridViewCell In dr.Cells
                    Select Case cel.ColumnIndex
                        Case 0
                            If showImage = True Then
                                ' immagine
                                html += vbTab + vbTab + TAG_CELL_ARTICOLO

                                If bLoadImage = True Then
                                    'carica immagini da cella
                                    If FormExport.CheckBoxFormatoImgFisso.Checked = True Then
                                        html += "<img style='display:block; width:100px;height:100px;' id='base64image' src='data:image/jpeg;base64, " & convertBmpTobase64(cel.Value) & "' />"

                                    Else
                                        ' ho tolto la forzatura alla dimensione 100x100 dell'immagine
                                        html += "<img style='display:block; ' id='base64image' src='data:image/jpeg;base64, " & convertBmpTobase64(cel.Value) & "' />"
                                    End If
                                Else
                                    srcFilename = getFilePath(dr.Cells(2).Value)
                                    If srcFilename = "" Then
                                        ' carica tappo
                                        If FormExport.CheckBoxFormatoImgFisso.Checked = True Then
                                            html += "<img style='display:block; width:100px;height:100px;' id='base64image' src='data:image/jpeg;base64, " & convertBmpTobase64(My.Resources.url) & "' />"
                                        Else
                                            ' ho tolto la forzatura alla dimensione 100x100 dell'immagine
                                            html += "<img style='display:block;' id='base64image' src='data:image/jpeg;base64, " & convertBmpTobase64(My.Resources.url) & "' />"
                                        End If


                                    Else
                                        If FormExport.CheckBoxFormatoImgFisso.Checked = True Then
                                            ' carica immagine da file
                                            html += "<img style='display:block; width:100px;height:100px;' id='base64image' src='data:image/jpeg;base64, " & convertFileToTobase64(srcFilename) & "' />"
                                        Else
                                            ' ho tolto la forzatura alla dimensione 100x100 dell'immagine
                                            html += "<img style='display:block;' id='base64image' src='data:image/jpeg;base64, " & convertFileToTobase64(srcFilename) & "' />"
                                        End If
                                    End If
                                End If


                                html += TAG_CELL_END
                            End If
                        Case 1
                            'codice
                            html += vbTab + vbTab + TAG_CELL_ARTICOLO
                            If FormExport.CheckBoxArtCodice.Checked = True Then
                                'html += vbTab + vbTab + TAG_CELL_ARTICOLO + Replace(cel.Value, vbCrLf, "<BR>") + TAG_CELL_END
                                html += vbTab + vbTab + "ART: " + varTag(idxArt).CodArt + " - " + varTag(idxArt).DescArt + " - " + varTag(idxArt).DescrInglese & "<BR>"


                            End If

                            'Stagione
                            If FormExport.CheckBoxArtStagione.Checked = True Then
                                html += vbTab + vbTab + "STAG: " + varTag(idxArt).CodStag & " - " & varTag(idxArt).DescStag & "<BR>"

                            End If

                            'Famiglia
                            If FormExport.CheckBoxArtFamiglia.Checked = True Then
                                html += vbTab + vbTab + "FAM: " + varTag(idxArt).Famiglia & "<BR>"

                            End If

                            'Famiglia
                            If FormExport.CheckBoxArtComposizione.Checked = True Then
                                html += vbTab + vbTab + "COMP: " + varTag(idxArt).Composizione & "<BR>"

                            End If

                            'Marca
                            If FormExport.CheckBoxMarca.Checked = True Then
                                html += vbTab + vbTab + "MARCA: " + varTag(idxArt).CodMarca + " - " + varTag(idxArt).DesMarca & "<BR>"
                            End If

                            'Nomenclatura
                            If FormExport.CheckBoxCodNomenclatura.Checked = True Then
                                html += vbTab + vbTab + "COD NOMENCLATURA: " + varTag(idxArt).CodNomenclatura & "<BR>"

                            End If

                            'MadeIn
                            If versione <> VER_35 Then
                                If FormExport.CheckBoxMadeIn.Checked = True Then
                                    html += vbTab + vbTab + "MADE IN: " + varTag(idxArt).MadeIn & "<BR>"
                                End If
                            End If


                            html += vbTab + vbTab + TAG_CELL_END
                        Case 4
                            If (CODLIS_ACQ = "" Or FormExport.CheckBoxExpLisAcquisto.Checked = False) And (CODLIS_VEN = "" Or FormExport.CheckBoxExpLisVendita.Checked = False) Then
                                Continue For
                            End If

                            html += vbTab + vbTab + TAG_CELL_PREZZO
                            If CODLIS_ACQ <> "" Then
                                If FormExport.CheckBoxExpLisAcquisto.Checked = True Then
                                    html += vbTab + vbTab + "Costo: " + "<BR>" + dr.Cells(4).Value.ToString + " €" & "<BR>" + "<BR>"
                                End If
                            End If
                            If CODLIS_VEN <> "" Then
                                If FormExport.CheckBoxExpLisVendita.Checked = True Then
                                    html += vbTab + vbTab + "Prezzo: " + "<BR>" + dr.Cells(5).Value.ToString & " €" + "<BR>" + "<BR>"
                                End If
                            End If
                            html += vbTab + vbTab + TAG_CELL_END


                    End Select
                Next
                html += vbTab + TAG_ROW_END + vbNewLine
                html += vbTab + TAG_TABLE_END + vbNewLine

                '--------------------
                ' Varianti
                'html += vbTab + TAG_TABLE_VUOTO + vbNewLine
                html += vbTab + TAG_TABLE_FIXED + vbNewLine
                If bShowImgVariante = True Then
                    html += vbTab + "<col width=""100px"" />" + vbNewLine
                End If
                html += vbTab + "<col width=""300px"" />" + vbNewLine
                For i = 1 To 100
                    html += vbTab + "<col width=""30px"" />" + vbNewLine

                Next

                html += vbTab + TAG_ROW
                html += vbTab + TAG_CELL

                If Not varTag(idxArt).Varianti Is Nothing Then


                    For idxVariante = 0 To UBound(varTag(idxArt).Varianti) - 1

                        ' controllo per escludere la variante che deve essere ignorata
                        If Array.IndexOf(escludiVar, varTag(idxArt).CodVariante(idxVariante)) >= 0 Then
                            Continue For
                        End If

                        saveHtml = html
                        bEsportaVariante = False

                        html += vbTab + TAG_ROW

                        ' controlla se mostrare l'immagine variante
                        If bShowImgVariante = True Then
                            srcFilename = getFileVarianti(varTag(idxArt).CodArt, varTag(idxArt).CodVariante(idxVariante))
                            If File.Exists(srcFilename) = True Then
                                ' carica immagine da file
                                'html += vbTab + TAG_CELL_VARIANTE + "<img style='display:block; width:100px;height:100px;' id='base64image' src='data:image/jpeg;base64, " & convertFileToTobase64(srcFilename) & "' />" + TAG_CELL_END

                                ' ho tolto la forzatura alla dimensione 100x100 dell'immagine
                                html += vbTab + TAG_CELL_VARIANTE + "<img style='display:block;' id='base64image' src='data:image/jpeg;base64, " & convertFileToTobase64(srcFilename) & "' />" + TAG_CELL_END
                            Else
                                html += vbTab + TAG_CELL_VARIANTE + "" + TAG_CELL_END
                            End If

                        End If

                        ' la riga commentata esportava solo il codice variante quella scommentata esporta codice + descrizione
                        ' html += vbTab + TAG_CELL_VARIANTE + "<font size=""3"">" + varTag(idxArt).CodVariante(idxVariante) + "</font>" + TAG_CELL_END
                        html += vbTab + TAG_CELL_VARIANTE + "<font size=""3"">" + varTag(idxArt).Varianti(idxVariante) + "</font>" + TAG_CELL_END

                        If bTaglieSempre = True Or indexPrimaRigaTaglia = 0 Then
                            ' controlla se mostrare l'immagine variante
                            html += vbTab + TAG_CELL_TAGLIA + "UM" + TAG_CELL_END
                            html += vbTab + TAG_CELL_TAGLIA + "TOT" + TAG_CELL_END
                        Else
                            html += vbTab + TAG_CELL_TAGLIA + "" + TAG_CELL_END
                            html += vbTab + TAG_CELL_TAGLIA + "" + TAG_CELL_END

                        End If


                        If varTag(idxArt).haTaglie = True Then
                            If Not varTag(idxArt).Giacenze(0).Giacenze Is Nothing Then

                                'intestazione taglie
                                For idxTaglia = 0 To UBound(varTag(idxArt).Giacenze(0).Taglie) - 1
                                    If Array.IndexOf(escludiTg, varTag(idxArt).Giacenze(idxVariante).Taglie(idxTaglia).ToString) >= 0 Then
                                        Continue For
                                    End If

                                    If bTaglieSempre = True Or indexPrimaRigaTaglia = 0 Then
                                        html += vbTab + TAG_CELL_TAGLIA + varTag(idxArt).Giacenze(idxVariante).Taglie(idxTaglia).ToString + TAG_CELL_END
                                    Else
                                        html += vbTab + TAG_CELL_TAGLIA + "" + TAG_CELL_END
                                    End If
                                    Application.DoEvents()
                                Next
                                indexPrimaRigaTaglia = indexPrimaRigaTaglia + 1
                                html += vbTab + TAG_ROW_END

                                If showGiacenze = True Then
                                    'If (exportGiacZero = False) And (varTag(idxArt).TotaleGiac(idxVariante) = 0) Then
                                    If (exportGiacZero = False) And (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) = 0) Then

                                    Else
                                        'If (exportGiacMinZero = False) And (varTag(idxArt).TotaleGiac(idxVariante) < 0) Then
                                        If (exportGiacMinZero = False) And (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) < 0) Then

                                        Else
                                            bEsportaVariante = True
                                            ' giacenze
                                            html += vbTab + TAG_ROW
                                            If bShowImgVariante = True Then
                                                html += vbTab + TAG_CELL_VARIANTE + "" + TAG_CELL_END
                                            End If
                                            html += vbTab + TAG_CELL_LEFT + LabelGiac + TAG_CELL_END
                                            html += vbTab + TAG_CELL + varTag(idxArt).UM + TAG_CELL_END
                                            'html += vbTab + TAG_CELL + varTag(idxArt).TotaleGiac(idxVariante).ToString + TAG_CELL_END
                                            html += vbTab + TAG_CELL + ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) + TAG_CELL_END
                                            For idxTaglia = 0 To UBound(varTag(idxArt).Giacenze(0).Giacenze) - 1
                                                If Array.IndexOf(escludiTg, varTag(idxArt).Giacenze(idxVariante).Taglie(idxTaglia).ToString) >= 0 Then
                                                    Continue For
                                                End If
                                                strOutVal = ""
                                                If varTag(idxArt).Giacenze(idxVariante).Giacenze(idxTaglia) <> 0 Then
                                                    strOutVal = varTag(idxArt).Giacenze(idxVariante).Giacenze(idxTaglia).ToString
                                                End If

                                                'controlla se deve esportare un massimo
                                                If FormExport.CheckBoxMaxQta.Checked = True Then
                                                    If strOutVal <> "" Then
                                                        If Int(strOutVal) >= FormExport.NumericUpDownMaxQta.Value Then
                                                            strOutVal = FormExport.NumericUpDownMaxQta.Value.ToString
                                                        End If
                                                    End If
                                                End If



                                                html += vbTab + TAG_CELL + strOutVal + TAG_CELL_END
                                                Application.DoEvents()
                                            Next
                                            html += vbTab + TAG_ROW_END
                                        End If

                                    End If
                                End If


                                If showCalcolato = True Then
                                    If (exportGiacZero = False) And (varTag(idxArt).TotaleCalc(idxVariante) = 0) Then

                                    Else
                                        'If (exportGiacMinZero = False) And (varTag(idxArt).TotaleCalc(idxVariante) < 0) Then
                                        If (exportGiacMinZero = False) And (ricalcolatotaleCalcPerEsportazione(idxArt, idxVariante) < 0) Then

                                        Else
                                            bEsportaVariante = True

                                            ' calcolato
                                            html += vbTab + TAG_ROW
                                            If bShowImgVariante = True Then
                                                html += vbTab + TAG_CELL_VARIANTE + "" + TAG_CELL_END
                                            End If
                                            html += vbTab + TAG_CELL_CALC_LEFT + labelCalc + TAG_CELL_END
                                            html += vbTab + TAG_CELL_CALC + varTag(idxArt).UM + TAG_CELL_END
                                            'html += vbTab + TAG_CELL + varTag(idxArt).TotaleCalc(idxVariante).ToString + TAG_CELL_END
                                            html += vbTab + TAG_CELL_CALC + ricalcolatotaleCalcPerEsportazione(idxArt, idxVariante) + TAG_CELL_END
                                            For idxTaglia = 0 To UBound(varTag(idxArt).Giacenze(0).Calcolato) - 1
                                                If Array.IndexOf(escludiTg, varTag(idxArt).Giacenze(idxVariante).Taglie(idxTaglia).ToString) >= 0 Then
                                                    Continue For
                                                End If
                                                strOutVal = ""
                                                If varTag(idxArt).Giacenze(idxVariante).Calcolato(idxTaglia) <> 0 Then
                                                    strOutVal = varTag(idxArt).Giacenze(idxVariante).Calcolato(idxTaglia).ToString
                                                End If

                                                'controlla se deve esportare un massimo
                                                If FormExport.CheckBoxMaxQta.Checked = True Then
                                                    If strOutVal <> "" Then
                                                        If Int(strOutVal) >= FormExport.NumericUpDownMaxQta.Value Then
                                                            strOutVal = FormExport.NumericUpDownMaxQta.Value.ToString
                                                        End If
                                                    End If
                                                End If

                                                html += vbTab + TAG_CELL_CALC + strOutVal + TAG_CELL_END
                                                Application.DoEvents()
                                            Next
                                            html += vbTab + TAG_ROW_END
                                        End If
                                    End If
                                End If

                            Else
                                html += vbTab + TAG_CELL + TAG_CELL_END

                            End If
                        Else
                            If showGiacenze = True Then
                                'If (exportGiacZero = False) And (varTag(idxArt).TotaleGiac(idxVariante) = 0) Then
                                If (exportGiacZero = False) And (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) = 0) Then

                                Else
                                    'If (exportGiacMinZero = False) And (varTag(idxArt).TotaleGiac(idxVariante) < 0) Then
                                    If (exportGiacMinZero = False) And (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) < 0) Then

                                    Else
                                        bEsportaVariante = True
                                        ' giacenze
                                        html += vbTab + TAG_ROW
                                        If bShowImgVariante = True Then
                                            html += vbTab + TAG_CELL_VARIANTE + "" + TAG_CELL_END
                                        End If
                                        html += vbTab + TAG_CELL_LEFT + LabelGiac + TAG_CELL_END
                                        html += vbTab + TAG_CELL + varTag(idxArt).UM + TAG_CELL_END
                                        'html += vbTab + TAG_CELL + varTag(idxArt).TotaleGiac(idxVariante).ToString + TAG_CELL_END
                                        html += vbTab + TAG_CELL + ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) + TAG_CELL_END

                                        html += vbTab + TAG_ROW_END
                                    End If
                                End If
                            End If

                            If showCalcolato = True Then
                                If (exportGiacZero = False) And (varTag(idxArt).TotaleCalc(idxVariante) = 0) Then

                                Else
                                    'If (exportGiacMinZero = False) And (varTag(idxArt).TotaleCalc(idxVariante) < 0) Then
                                    If (exportGiacMinZero = False) And (ricalcolatotaleCalcPerEsportazione(idxArt, idxVariante) < 0) Then

                                    Else
                                        bEsportaVariante = True
                                        ' calcolato
                                        html += vbTab + TAG_ROW
                                        If bShowImgVariante = True Then
                                            html += vbTab + TAG_CELL_VARIANTE + "" + TAG_CELL_END
                                        End If
                                        html += vbTab + TAG_CELL_CALC_LEFT + labelCalc + TAG_CELL_END
                                        html += vbTab + TAG_CELL_CALC + varTag(idxArt).UM + TAG_CELL_END
                                        'html += vbTab + TAG_CELL + varTag(idxArt).TotaleCalc(idxVariante).ToString + TAG_CELL_END
                                        html += vbTab + TAG_CELL_CALC + ricalcolatotaleCalcPerEsportazione(idxArt, idxVariante) + TAG_CELL_END
                                        html += vbTab + TAG_ROW_END
                                    End If
                                End If
                            End If
                            indexPrimaRigaTaglia = indexPrimaRigaTaglia + 1

                        End If
                        html += vbTab + TAG_ROW_END

                        If bEsportaVariante = False Then
                            ' ripristino il vecchio valore in quanto non è stata esportata nessuna riga
                            html = saveHtml
                            indexPrimaRigaTaglia = indexPrimaRigaTaglia - 1
                        Else
                            ' se ha esportato almeno una volta lascia l'articolo
                            bEsportaArt = True
                        End If
                        Application.DoEvents()



                    Next
                Else
                    'cella vuota
                    html += vbTab + TAG_CELL + TAG_CELL_END

                End If

                html += vbTab + TAG_CELL_END
                html += vbTab + TAG_ROW_END


                html += vbTab + TAG_ROW_END + vbNewLine
                html += vbTab + TAG_TABLE_END + vbNewLine
                html += vbTab + "<BR>" + vbNewLine
                html += vbTab + "<BR>" + vbNewLine
                html += vbTab + "<BR>" + vbNewLine

                ' controlla se deve gestire l'interruzione pagina
                If FormExport.CheckBoxInterPag.Checked = True Then
                    If idxArticolo = FormExport.NumericUpDownInterrPag.Value Then
                        html += vbTab + "<p style=""page-break-after:always;""></p>"
                        idxArticolo = 1
                    Else
                        idxArticolo = idxArticolo + 1
                    End If
                End If

                If bEsportaArt = False Then
                    html = saveHtmlArt
                End If

                AppendHTML(tmpHTML, html)


                'idxArt = idxArt + 1
                Application.DoEvents()
            Next

            html += "</table>"
            AppendHTML(tmpHTML, html)




            LabelDB.Text = "Creazione file in corso ..."
            Application.DoEvents()

            Dim value As String = "" 'html
            Dim leftString As String = ""
            Dim rightString As String = ""
            If File.Exists(iniPath + "template.html") Then
                value = File.ReadAllText(iniPath + "template.html")

                leftString = Strings.Left(value, value.IndexOf("%%REPORT%%"))
                rightString = Strings.Right(value, Len(value) - (value.IndexOf("%%REPORT%%") + Len("%%REPORT%%")))

                value = Replace(value, "%%REPORT%%", html)
            End If


            'scrivo il file a pezzetti a causa delle grosse dimensioni
            If File.Exists(SaveFileDialog1.FileName) Then
                Kill(SaveFileDialog1.FileName)
            End If
            AppendHTML(SaveFileDialog1.FileName, leftString)
            Dim objReader As New System.IO.StreamReader(tmpHTML)
            Do While objReader.Peek() <> -1
                AppendHTML(SaveFileDialog1.FileName, objReader.ReadLine())
                Application.DoEvents()
            Loop
            objReader.Close()
            AppendHTML(SaveFileDialog1.FileName, rightString)



            'Catch ex As Exception
            'MsgBox(ex.Message)
            'End Try
            MsgBox("Esportazione avvenuta con successo!")
            LabelDB.Text = "Esportazione avvenuta con successo"
            Process.Start(SaveFileDialog1.FileName)
        End If


    End Sub


    Private Function ricalcolatotaleGiacPerEsportazione(idxArt As Integer, idxVariante As Integer) As String
        Dim i As Integer
        Dim totale As Integer = 0

        If varTag(idxArt).Giacenze(idxVariante).Giacenze Is Nothing Then
            Return varTag(idxArt).TotaleGiac(idxVariante)
        End If

        For i = 0 To UBound(varTag(idxArt).Giacenze(idxVariante).Giacenze) - 1

            If FormExport.CheckBoxMaxQta.Checked = False Then
                totale = totale + varTag(idxArt).Giacenze(idxVariante).Giacenze(i)
            Else
                If varTag(idxArt).Giacenze(idxVariante).Giacenze(i) >= FormExport.NumericUpDownMaxQta.Value Then
                    totale = totale + FormExport.NumericUpDownMaxQta.Value
                Else
                    totale = totale + varTag(idxArt).Giacenze(idxVariante).Giacenze(i)
                End If

            End If


        Next

        Return totale

    End Function


    Private Function ricalcolatotaleCalcPerEsportazione(idxArt As Integer, idxVariante As Integer) As String
        Dim i As Integer
        Dim totale As Integer = 0

        If varTag(idxArt).Giacenze(idxVariante).Giacenze Is Nothing Then
            Return varTag(idxArt).TotaleCalc(idxVariante)
        End If

        For i = 0 To UBound(varTag(idxArt).Giacenze(idxVariante).Giacenze) - 1

            If FormExport.CheckBoxMaxQta.Checked = False Then
                totale = totale + varTag(idxArt).Giacenze(idxVariante).Calcolato(i)
            Else
                If varTag(idxArt).Giacenze(idxVariante).Giacenze(i) >= FormExport.NumericUpDownMaxQta.Value Then
                    totale = totale + FormExport.NumericUpDownMaxQta.Value
                Else
                    totale = totale + varTag(idxArt).Giacenze(idxVariante).Calcolato(i)
                End If

            End If


        Next

        Return totale

    End Function


    Private Function convertBmpTobase64(img As System.Drawing.Bitmap) As String
        Dim resizeVal As Integer

        If img Is Nothing Then
            Return ""
        End If
        If img.Width > 100 Then
            resizeVal = 1
        End If
        If img.Width > 200 Then
            resizeVal = 2
        End If
        If img.Width > 300 Then
            resizeVal = 3
        End If
        If img.Width > 400 Then
            resizeVal = 4
        End If
        If img.Width > 800 Then
            resizeVal = 8
        End If
        If img.Width > 1000 Then
            resizeVal = 10
        End If

        ' riduce l'immagie di 10 volte per evitare che diventi troppo grande
        Dim tmpfilename As String = tmp_dir & "tmp.png"
        img = ResizeImage(img, (img.Width \ resizeVal), (img.Height \ resizeVal))
        img.Save(tmpfilename)
        Return Convert.ToBase64String(System.IO.File.ReadAllBytes(tmpfilename))

    End Function


    Private Function convertFileToTobase64(filename As String) As String
        Dim resizeVal As Integer
        Dim img As System.Drawing.Bitmap
        img = Image.FromFile(filename)

        If img Is Nothing Then
            Return ""
        End If
        If img.Width > 100 Then
            resizeVal = 1
        End If
        If img.Width > 200 Then
            resizeVal = 2
        End If
        If img.Width > 300 Then
            resizeVal = 3
        End If
        If img.Width > 400 Then
            resizeVal = 4
        End If
        If img.Width > 800 Then
            resizeVal = 8
        End If
        If img.Width > 1000 Then
            resizeVal = 10
        End If

        ' riduce l'immagie di 10 volte per evitare che diventi troppo grande
        Dim tmpfilename As String = tmp_dir & "tmp.png"
        img = ResizeImage(img, (img.Width \ resizeVal), (img.Height \ resizeVal))
        img.Save(tmpfilename)
        Return Convert.ToBase64String(System.IO.File.ReadAllBytes(tmpfilename))

    End Function


    Private Sub readIniFileImg(filename As String)
        'Dim VL_FileName As String = filename
        Dim sb As System.Text.StringBuilder
        'Dim Sezione As String = "IMPOSTAZIONI"

        ''-----------------------------------------------------------------
        '' FILTRI
        'sb = New System.Text.StringBuilder(500)
        'GetPrivateProfileString(Sezione, "cartella immagini", "", sb, sb.Capacity, VL_FileName)
        'If sb.ToString <> "" Then
        '    CODARTICOLO_DA = sb.ToString
        'End If

        'sb = New System.Text.StringBuilder(500)
        'GetPrivateProfileString(Sezione, "estensione", "", sb, sb.Capacity, VL_FileName)
        'If sb.ToString <> "" Then
        '    CODARTICOLO_DA = sb.ToString
        'End If
        Dim expIni As String = Path.GetDirectoryName(filename) & "\expGiacenze.ini"


        If System.IO.File.Exists(expIni) = True Then

            sb = New System.Text.StringBuilder(500)
            GetPrivateProfileString("IMPOSTAZIONI", "tmpDir", "", sb, sb.Capacity, expIni)
            If sb.ToString <> "" Then
                tmp_dir = sb.ToString
            End If

        End If


        Dim TextLine As String

        If System.IO.File.Exists(filename) = True Then

            Dim objReader As New System.IO.StreamReader(filename)

            Do While objReader.Peek() <> -1
                TextLine = objReader.ReadLine()
                If TextLine.Contains("cartella immagini") = True Then
                    FOLDER_IMG_VAR = Trim(Strings.Right(TextLine, TextLine.Length - Len("cartella immagini=")))
                    If Strings.Right(FOLDER_IMG_VAR, 1) <> "\" Then
                        FOLDER_IMG_VAR = FOLDER_IMG_VAR & "\"
                    End If
                    Continue Do
                End If

                If TextLine.Contains("estensione") = True Then
                    EXTENSION_IMG_VAR = Trim(Strings.Right(TextLine, TextLine.Length - Len("estensione=")))
                    If Strings.Left(EXTENSION_IMG_VAR, 1) = "." Then
                        EXTENSION_IMG_VAR = Strings.Right(EXTENSION_IMG_VAR, EXTENSION_IMG_VAR.length - 1)
                    End If
                    Continue Do
                End If

            Loop
        End If




    End Sub


    Private Sub readIniFile(filename As String)
        Dim VL_FileName As String = filename
        Dim sb As System.Text.StringBuilder
        Dim Sezione As String = "FILTRI"
        Dim Sezione_Server As String = "SERVER"
        Dim server As String = "MONICAGIO-WIN7\SISTEMI"
        Dim DB As String = "Esmoda38"
        Dim user As String = "sa"
        Dim psw As String = "Sistemi123"
        Dim i As Integer


        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "FLAG_ORDINATO", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            FLAG_ORDINATO = sb.ToString
        End If

        '-----------------------------------------------------------------
        ' FILTRI
        '-----------------------------------------------------------------
        '
        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "TIPOART_DA", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            TIPOART_DA = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "TIPOART_A", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            TIPOART_A = sb.ToString
        End If


        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODARTICOLO_DA", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODARTICOLO_DA = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODARTICOLO_A", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODARTICOLO_A = sb.ToString
        End If


        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODMARCA_DA", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODMARCA_DA = sb.ToString
        End If


        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "FAMIGLIA_DA", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            FAMIGLIA_DA = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "MACRFAM_DA", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            MACROFAM_DA = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "STAGIONE_DA", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            STAGIONE_DA = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "LINEA_DA", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            LINEA_DA = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODMAGAZZINO", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODMAGAZZINO1 = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "DETTMAG", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            DETTAGLIO_MAG = sb.ToString
        End If

        CODMAGAZZINO = "('"
        If CODMAGAZZINO1 <> "" Then CODMAGAZZINO = CODMAGAZZINO + Trim(CODMAGAZZINO1)
        CODMAGAZZINO = CODMAGAZZINO + "')"


        LISTAMAGAZZINI = ""
        If CODMAGAZZINO1 <> "" Then LISTAMAGAZZINI = LISTAMAGAZZINI + Trim(CODMAGAZZINO1)

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODGRUPPO", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODGRUPPO = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODUTEPERS", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODUTEPERS = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODLIS_VEN", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODLIS_VEN = sb.ToString
        End If


        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "DESLIS_VEN", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            DESLIS_VEN = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODLIS_ACQ", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODLIS_ACQ = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "DESLIS_ACQ", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            DESLIS_ACQ = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODFOR", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODFOR = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODAGE", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODAGE = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODZONA", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODZON = sb.ToString
        End If


        For i = 1 To 20
            sb = New System.Text.StringBuilder(500)
            GetPrivateProfileString(Sezione, "CS" & i & "_DA", "", sb, sb.Capacity, VL_FileName)
            codStatDA(i) = ""
            If sb.ToString <> "" Then
                codStatDA(i) = sb.ToString
            End If
        Next

        '-----------------------------------------------------------------
        ' lettura sezione server 
        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione_Server, "SERVER", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            server = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione_Server, "USER", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            user = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione_Server, "PWD", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            psw = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione_Server, "DATABASE", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            DB = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione_Server, "VERSIONE", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            versione = sb.ToString
        End If


        Connessione = "Provider=sqloledb;Data Source=" & server & ";Initial Catalog=" & DB & ";User ID=" & user & ";Password=" & psw & ";"



    End Sub



#Region " ResizeImage "
    Public Overloads Shared Function ResizeImage(SourceImage As Drawing.Image, TargetWidth As Int32, TargetHeight As Int32) As Drawing.Bitmap
        Dim bmSource = New Drawing.Bitmap(SourceImage)

        Return ResizeImage(bmSource, TargetWidth, TargetHeight)
    End Function

    Public Overloads Shared Function ResizeImage(bmSource As Drawing.Bitmap, TargetWidth As Int32, TargetHeight As Int32) As Drawing.Bitmap
        Dim bmDest As New Drawing.Bitmap(TargetWidth, TargetHeight, Drawing.Imaging.PixelFormat.Format32bppArgb)

        Dim nSourceAspectRatio = bmSource.Width / bmSource.Height
        Dim nDestAspectRatio = bmDest.Width / bmDest.Height

        Dim NewX = 0
        Dim NewY = 0
        Dim NewWidth = bmDest.Width
        Dim NewHeight = bmDest.Height

        If nDestAspectRatio = nSourceAspectRatio Then
            'same ratio
        ElseIf nDestAspectRatio > nSourceAspectRatio Then
            'Source is taller
            NewWidth = Convert.ToInt32(Math.Floor(nSourceAspectRatio * NewHeight))
            NewX = Convert.ToInt32(Math.Floor((bmDest.Width - NewWidth) / 2))
        Else
            'Source is wider
            NewHeight = Convert.ToInt32(Math.Floor((1 / nSourceAspectRatio) * NewWidth))
            NewY = Convert.ToInt32(Math.Floor((bmDest.Height - NewHeight) / 2))
        End If

        Using grDest = Drawing.Graphics.FromImage(bmDest)
            With grDest
                .CompositingQuality = Drawing.Drawing2D.CompositingQuality.HighQuality
                .InterpolationMode = Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
                .PixelOffsetMode = Drawing.Drawing2D.PixelOffsetMode.HighQuality
                .SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias
                .CompositingMode = Drawing.Drawing2D.CompositingMode.SourceOver

                .DrawImage(bmSource, NewX, NewY, NewWidth, NewHeight)
            End With
        End Using

        Return bmDest
    End Function

#End Region

    Private Sub BtnFiltri_Click(sender As Object, e As EventArgs) Handles BtnFiltri.Click

        If FormFiltri.ShowDialog() = vbOK Then
            Call PopolaGriglia(False)
        End If

    End Sub



    ' ----------------------------------------------
    ' gestione filtri ARTICOLI
    Private Sub AddArticoloInList(CodArt As String, descrizione As String)
        Dim mysize As Integer
        Dim i As Integer

        If Not ListFiltriArticoli Is Nothing Then
            ' cerca l'articolo se lo trova
            For i = 0 To ListFiltriArticoli.Count - 1
                If ListFiltriArticoli(i).Codice = CodArt Then
                    Exit Sub
                End If
            Next

        End If

        mysize = 0
        If Not ListFiltriArticoli Is Nothing Then
            mysize = UBound(ListFiltriArticoli)
        End If
        ReDim Preserve ListFiltriArticoli(mysize + 1)

        ListFiltriArticoli(UBound(ListFiltriArticoli) - 1).Codice = CodArt
        ListFiltriArticoli(UBound(ListFiltriArticoli) - 1).descrizione = descrizione
        ListFiltriArticoli(UBound(ListFiltriArticoli) - 1).visibile = True




    End Sub

    Private Function checkVisibilitaArticolo(codArt As String) As Boolean
        Dim res As Boolean = True

        ' il primo giro ignora i filtri
        If PrimoGiro = False Then
            ' cerca l'articolo se lo trova
            For i = 0 To ListFiltriArticoli.Count - 1
                If ListFiltriArticoli(i).Codice = codArt Then
                    res = ListFiltriArticoli(i).visibile
                    Exit For
                End If
            Next
        End If
        Return res
    End Function




    ' ----------------------------------------------
    ' gestione filtri VARIANTI
    Private Sub AddVarianteInList(CodVar As String)
        Dim mysize As Integer
        Dim i As Integer

        If Not ListFiltriVarianti Is Nothing Then
            ' cerca l'articolo se lo trova
            For i = 0 To ListFiltriVarianti.Count - 1
                If ListFiltriVarianti(i).Codice = CodVar Then
                    Exit Sub
                End If
            Next

        End If

        mysize = 0
        If Not ListFiltriVarianti Is Nothing Then
            mysize = UBound(ListFiltriVarianti)
        End If
        ReDim Preserve ListFiltriVarianti(mysize + 1)

        ListFiltriVarianti(UBound(ListFiltriVarianti) - 1).Codice = CodVar
        ListFiltriVarianti(UBound(ListFiltriVarianti) - 1).visibile = True

    End Sub


    Private Function checkVisibilitaVariante(CodVar As String) As Boolean
        Dim res As Boolean = True

        ' il primo giro ignora i filtri
        If PrimoGiro = False Then
            ' cerca l'articolo se lo trova
            For i = 0 To ListFiltriVarianti.Count - 1
                If ListFiltriVarianti(i).Codice = CodVar Then
                    res = ListFiltriVarianti(i).visibile
                    Exit For
                End If
            Next
        End If
        Return res
    End Function

    ' ----------------------------------------------
    ' gestione filtri TAGLIE
    Private Sub AddTagliaInList(codTaglia As String)
        Dim mysize As Integer
        Dim i As Integer

        If Not ListFiltriTaglie Is Nothing Then
            ' cerca l'articolo se lo trova
            For i = 0 To ListFiltriTaglie.Count - 1
                If ListFiltriTaglie(i).Codice = codTaglia Then
                    Exit Sub
                End If
            Next

        End If

        mysize = 0
        If Not ListFiltriTaglie Is Nothing Then
            mysize = UBound(ListFiltriTaglie)
        End If
        ReDim Preserve ListFiltriTaglie(mysize + 1)

        ListFiltriTaglie(UBound(ListFiltriTaglie) - 1).Codice = codTaglia
        ListFiltriTaglie(UBound(ListFiltriTaglie) - 1).visibile = True

    End Sub

    Private Function checkVisibilitaTaglia(CodTaglia As String, giacCalcolato As Integer, giacNormale As Integer) As Boolean
        Dim res As Boolean = True

        ' il primo giro ignora i filtri
        If PrimoGiro = False Then
            ' cerca l'articolo se lo trova
            For i = 0 To ListFiltriTaglie.Count - 1
                If ListFiltriTaglie(i).Codice = CodTaglia Then
                    If giacCalcolato > 0 And giacNormale > 0 Then
                        res = ListFiltriTaglie(i).visibile
                    End If
                    Exit For
                End If
            Next
        End If
        Return res

    End Function

    Private Sub getAssegnatoTaglie(codArt As String, codVar As String, ByRef assegnato() As Integer)
        Dim sql As String
        Dim rs As New ADODB.Recordset
        Dim res As Integer = 0
        Dim idx As Integer

        For idx = 1 To 30
            assegnato(idx) = 0
        Next

        sql = getSQLAssegnatoTaglie(codArt, codVar)
        Try
            rs.Open(sql, connSqlSrv, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(sql & "    ---SQLASSEGNATOTAGLIE IN GETASSEGNATOTAGLIE")
        End Try

        If rs.RecordCount = 1 Then

            For idx = 1 To 30
                assegnato(idx) = rs.Fields("tg" & idx).Value
            Next
        End If

        rs.Close()

    End Sub

    Private Function getSQLAssegnatoTaglie(codArt As String, codVar As String) As String
        Dim SQL As String

        SQL = SQL & "Select codArt, VarianteArt,  "

        SQL = SQL & " sum(QtaTagPrelevata_1)  As tg1, sum(QtaTagPrelevata_2) As tg2, sum(QtaTagPrelevata_3) As tg3,  sum(QtaTagPrelevata_4) As tg4, sum(QtaTagPrelevata_5)As tg5, sum(QtaTagPrelevata_6)As tg6, sum(QtaTagPrelevata_7)As tg7, sum(QtaTagPrelevata_8)As tg8, sum(QtaTagPrelevata_9) As tg9,sum(QtaTagPrelevata_10) As tg10,  "
        SQL = SQL & " sum(QtaTagPrelevata_11) As tg11, sum(QtaTagPrelevata_12) As tg12, sum(QtaTagPrelevata_13) As tg13,  sum(QtaTagPrelevata_14) As tg14, sum(QtaTagPrelevata_15) As tg15, sum(QtaTagPrelevata_16) As tg16, sum(QtaTagDaSpedire_17- QtaTagPrelevata_17) As tg17, sum(QtaTagPrelevata_18) As tg18, sum(QtaTagPrelevata_19) As tg19, sum(QtaTagPrelevata_20) As tg20,  "
        SQL = SQL & " sum(QtaTagPrelevata_21) As tg21, sum(QtaTagPrelevata_22) As tg22, sum(QtaTagPrelevata_23) As tg23,  sum(QtaTagPrelevata_24) As tg24, sum(QtaTagPrelevata_25) As tg25, sum(QtaTagPrelevata_26) As tg26, sum(QtaTagPrelevata_27) As tg27, sum(QtaTagPrelevata_28) As tg28, sum(QtaTagPrelevata_29) As tg29, sum(QtaTagPrelevata_30) As tg30  "

        SQL = SQL & " From ordSpedRighe   "
        SQL = SQL & "  Left Join ModaSpedTestate on (ModaSpedTestate.IddDoc = OrdSpedRighe.IdDocumento And ModaSpedTestate.DBGruppo = OrdSpedRighe.DBGruppo)  "
        SQL = SQL & "  Left Join ModaSpedRighe on (ModaSpedRighe.IddOrdSped = OrdSpedRighe.IdDocumento And ModaSpedRighe.IdrOrdSped = ordSpedRighe.IdRiga And ModaSpedTestate.DBGruppo = OrdSpedRighe.DBGruppo)"

        SQL = SQL & " where 1 = 1 "
        SQL = SQL & " And ordSpedRighe.RigaSaldata = 0 "
        SQL = SQL & " And ModaSpedTestate.OrdStampBolla = 0 "
        SQL = SQL & " And ordSpedRighe.CodMag In " & CODMAGAZZINO & " "
        SQL = SQL & " And OrdSpedRighe.CodArt = '" & Replace(Trim(codArt), "'", "''") & "' "
        SQL = SQL & " And ordSpedRighe.VarianteArt = '" & codVar & "' "
        SQL = SQL & " group by CodArt, VarianteArt"

        Return SQL
    End Function

    Private Function getSQLAssegnato(codArt As String, codVar As String) As String
        Dim SQL As String

        SQL = SQL & "Select codArt, VarianteArt, (sum(QtaDaSpedire) - sum(QtaPrelevata)) As Assegnato  from ordSpedRighe  "
        SQL = SQL & " Left Join ModaSpedTestate on (ModaSpedTestate.IddDoc = OrdSpedRighe.IdDocumento And ModaSpedTestate.DBGruppo = OrdSpedRighe.DBGruppo) "
        SQL = SQL & " where 1 = 1 "
        SQL = SQL & " And ordSpedRighe.RigaSaldata = 0 "
        SQL = SQL & " And ordSpedRighe.PrelievoConfermato = 0 "
        SQL = SQL & " And ModaSpedTestate.OrdStampBolla = 0 "
        SQL = SQL & " And ordSpedRighe.CodMag In " & CODMAGAZZINO & " "
        SQL = SQL & " And OrdSpedRighe.CodArt = '" & Replace(Trim(codArt), "'", "''") & "' "
        SQL = SQL & " And ordSpedRighe.VarianteArt = '" & codVar & "' "
        SQL = SQL & " group by CodArt, VarianteArt"

        Return SQL
    End Function

    ' esporta l'html delle giacenze
    Private Sub exporEXCEL(showGiacenze As Boolean, showCalcolato As Boolean, showImage As Boolean, LabelGiac As String, labelCalc As String)
        Dim indexcont As Integer = 0
        Dim stroutput As String
        Dim idxArt As Integer = 0
        Dim idxVariante As Integer
        Dim idxTaglia As Integer
        Dim strPathTmp As String = "C:\DatasoftTmp\" 'Environment.GetFolderPath(Environment.SpecialFolder.Windows) & 
        stroutput = ""
        Dim exportGiacZero As Boolean
        Dim exportGiacMinZero As Boolean
        Dim exportMaggioriDi As Boolean = 0
        Dim bEsportaVariante As Boolean = False
        Dim bEsportaArt As Boolean = False
        Dim bTaglieSempre As Boolean
        Dim bShowImgVariante As Boolean
        Dim indexPrimaRigaTaglia As Integer = 0
        Dim strOutVal As String = ""
        Dim exportQtaMinima As Boolean

        Dim escludiVar() As String
        Dim escludiTg() As String
        Dim idxArticolo As Integer = 1
        Dim MaxWidth As Integer = 0



        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim idxRow As Integer = 0
        Dim srcFilename As String

        Dim stampaTestata As Boolean
        Dim testgiastampata As Boolean

        Dim r As Excel.Range
        Dim shape As Excel.Shape
        Dim larghezza As Double
        Dim altezza As Double
        Dim mypoints As Double

        Dim strvAl As String = ""

        Dim stampaArt As Boolean

        Try
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets(1)

        Catch ex As Exception

            MsgBox(ex.Message)
            MsgBox("Non è stata riscontrata l'installazione di Excel sul PC o la versione non è compatibile - verificare")
            Return
        End Try


        bShowImgVariante = FormExport.CheckBoxImgVariante.Checked

        escludiVar = Split(FormExport.TextBoxEscludiVarianti.Text, ",")
        escludiTg = Split(FormExport.TextBoxEscludiTaglie.Text, ",")

        exportGiacZero = FormExport.CheckBoxTotaliZero.Checked
        bTaglieSempre = FormExport.CheckBoxTaglie.Checked
        exportGiacMinZero = FormExport.CheckBoxTotaliMinZero.Checked
        exportMaggioriDi = 1


        SaveFileDialog1.Filter = "File Excel|*.xlsx"
        SaveFileDialog1.Title = "Salva un file Excel"
        SaveFileDialog1.ShowDialog()



        If SaveFileDialog1.FileName <> "" Then

            LabelDB.Text = "Esportazione file excel in corso ...."


            ' data ora esportazione
            'idxRow = 1
            idxRow = idxRow + 1
            xlWorkSheet.Cells(idxRow, 1) = "Data ora esportazione"
            xlWorkSheet.Cells(idxRow, 1).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
            xlWorkSheet.Cells(idxRow, 1).font.bold = True
            xlWorkSheet.Cells(idxRow, 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

            xlWorkSheet.Cells(idxRow, 2) = Now
            xlWorkSheet.Cells(idxRow, 2).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
            xlWorkSheet.Cells(idxRow, 2).font.bold = True
            xlWorkSheet.Cells(idxRow, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous


            idxRow = idxRow + 1



            xlWorkSheet.Cells(idxRow, 1) = "Codici Magazzini"
            xlWorkSheet.Cells(idxRow, 1).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
            xlWorkSheet.Cells(idxRow, 1).font.bold = True
            xlWorkSheet.Cells(idxRow, 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

            xlWorkSheet.Cells(idxRow, 2) = LISTAMAGAZZINI
            xlWorkSheet.Cells(idxRow, 2).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
            xlWorkSheet.Cells(idxRow, 2).font.bold = True
            xlWorkSheet.Cells(idxRow, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous


            idxRow = idxRow + 2


            ProgressBarArticoli.Minimum = 0
            ProgressBarArticoli.Maximum = DataGridViewArticoli.Rows.Count
            ProgressBarArticoli.Value = 0
            For Each dr As DataGridViewRow In DataGridViewArticoli.Rows

                If dr.Visible = False Then
                    Continue For
                End If

                For i = 0 To UBound(varTag) - 1
                    If varTag(i).CodArt = dr.Cells(6).Value Then
                        idxArt = i
                        Exit For
                    End If
                Next

                ProgressBarArticoli.Value = ProgressBarArticoli.Value + 1
                indexPrimaRigaTaglia = 0

                bEsportaArt = False

                xlWorkSheet.Range("A1:X1").EntireColumn.AutoFit()


                Dim almenoUnaVar = 0

                If FormExport.CheckBoxGiac.Checked = True Then
                    For idxVariante = 0 To UBound(varTag(idxArt).Varianti) - 1
                        If Array.IndexOf(escludiVar, varTag(idxArt).CodVariante(idxVariante)) >= 0 Then
                            Continue For
                        End If

                        If (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) = 0) And exportGiacZero = False Then
                            Continue For
                        End If

                        If (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) < 0) And exportGiacMinZero = False Then
                            Continue For
                        End If
                        '
                        almenoUnaVar = 1
                    Next

                    If almenoUnaVar = 0 Then
                        stampaArt = False
                    Else
                        stampaArt = True
                    End If
                End If


                If FormExport.CheckBoxCalcolato.Checked = True Then
                    almenoUnaVar = 0
                    For idxVariante = 0 To UBound(varTag(idxArt).Varianti) - 1
                        If Array.IndexOf(escludiVar, varTag(idxArt).CodVariante(idxVariante)) >= 0 Then
                            Continue For
                        End If

                        If (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) = 0) And exportGiacZero = False Then
                            Continue For
                        End If

                        If (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) < 0) And exportGiacMinZero = False Then
                            Continue For
                        End If

                        almenoUnaVar = 1
                    Next

                    If almenoUnaVar = 0 Then
                        stampaArt = False
                    Else
                        stampaArt = True
                    End If

                End If

                If stampaArt = False Then
                    Continue For
                End If


                '--------------------
                ' articolo
                idxRow = idxRow + 1
                For Each cel As DataGridViewCell In dr.Cells
                    Select Case cel.ColumnIndex
                        Case 0

                            If showImage = True Then
                                ' immagine
                                mypoints = 0
                                'Dim larghezza As Double
                                'Dim altezza As Double

                                If bLoadImage = True Then
                                    srcFilename = getTmpImagefilename(cel.Value, larghezza, altezza)
                                Else
                                    srcFilename = getTmpImagefilename(My.Resources.url, larghezza, altezza)
                                End If

                                If FormExport.CheckBoxFormatoImgFisso.Checked = True Then
                                    larghezza = 100
                                    altezza = 100
                                Else
                                    If altezza > MAX_ALTEZZA Then
                                        larghezza = (larghezza * MAX_ALTEZZA) / altezza
                                        altezza = MAX_ALTEZZA
                                    End If


                                End If
                                Dim g As Graphics = CreateGraphics()
                                ' Dim r As Excel.Range
                                'Dim shape As Excel.Shape

                                mypoints = larghezza * 18 / g.DpiX


                                ' si salva la dimensione massima in larghezza 
                                If mypoints > MaxWidth Then
                                    MaxWidth = mypoints
                                Else
                                    mypoints = MaxWidth
                                End If


                                g.Dispose()
                                xlWorkSheet.Range("A" & idxRow).ColumnWidth = mypoints
                                xlWorkSheet.Range("A" & idxRow).RowHeight = altezza + 2
                                r = xlWorkSheet.Cells(idxRow, 1)
                                shape = xlWorkSheet.Shapes.AddPicture(srcFilename, False, True, r.Left, r.Top, larghezza, altezza)

                                xlWorkSheet.Cells(idxRow, 1).HorizontalAlignment = Excel.Constants.xlCenter
                                xlWorkSheet.Cells(idxRow, 1).VerticalAlignment = Excel.Constants.xlCenter


                            Else
                                xlWorkSheet.Cells(idxRow, 1) = ""

                            End If
                            ' xlWorkSheet.Cells(idxRow, 1).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray)

                        Case 1

                            stampaTestata = True
                            testgiastampata = False

                        Case 4
                            strvAl = ""
                            If (CODLIS_ACQ = "" Or FormExport.CheckBoxExpLisAcquisto.Checked = False) And (CODLIS_VEN = "" Or FormExport.CheckBoxExpLisVendita.Checked = False) Then
                                Continue For
                            End If

                            If CODLIS_ACQ <> "" Then
                                If FormExport.CheckBoxExpLisAcquisto.Checked = True Then
                                    strvAl += "Costo: " + vbCrLf + dr.Cells(4).Value.ToString + " Euro" & vbCrLf & vbCrLf
                                End If
                            End If
                            If CODLIS_VEN <> "" Then
                                If FormExport.CheckBoxExpLisVendita.Checked = True Then
                                    strvAl += "Prezzo: " + vbCrLf + dr.Cells(5).Value.ToString & " Euro" & vbCrLf & vbCrLf
                                End If
                            End If
                            xlWorkSheet.Cells(idxRow, 3) = strvAl
                            xlWorkSheet.Cells(idxRow, 3).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan)
                            xlWorkSheet.Cells(idxRow, 3).VerticalAlignment = Excel.Constants.xlCenter
                            xlWorkSheet.Cells(idxRow, 3).HorizontalAlignment = Excel.Constants.xlLeft
                            xlWorkSheet.Cells(idxRow, 3).font.bold = True
                            xlWorkSheet.Cells(idxRow, 3).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

                    End Select
                Next


                'idxRow = idxRow + 1
                '--------------------
                ' TODO - VERIFICARE  Varianti
                'html += vbTab + TAG_TABLE_VUOTO + vbNewLine
                'html += vbTab + TAG_TABLE_FIXED + vbNewLine
                'If bShowImgVariante = True Then
                '    html += vbTab + "<col width=""100px"" />" + vbNewLine
                'End If
                'html += vbTab + "<col width=""300px"" />" + vbNewLine
                'For i = 1 To 100
                '    html += vbTab + "<col width=""30px"" />" + vbNewLine

                'Next


                If Not varTag(idxArt).Varianti Is Nothing Then

                    For idxVariante = 0 To UBound(varTag(idxArt).Varianti) - 1
                        ' controllo per escludere la variante che deve essere ignorata
                        If Array.IndexOf(escludiVar, varTag(idxArt).CodVariante(idxVariante)) >= 0 Then
                            Continue For
                        End If

                        If showGiacenze = True And exportGiacZero = False Then
                            If (exportMaggioriDi = True) And (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) <= FormExport.NumericUpDownMaggioriDi.Value) Then
                                Continue For
                            Else

                                If (exportGiacZero = False) And (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) = 0) Then
                                    Continue For
                                Else
                                    'If (exportGiacMinZero = False) And (varTag(idxArt).TotaleGiac(idxVariante) < 0) Then
                                    If (exportGiacMinZero = False) And (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) < 0) Then
                                        Continue For
                                    End If
                                End If
                            End If
                        End If


                        idxRow = idxRow + 1
                        bEsportaVariante = False

                        ' TODO - VERIFICARE controlla se mostrare l'immagine variante
                        'If bShowImgVariante = True Then
                        '    srcFilename = getFileVarianti(varTag(idxArt).CodArt, varTag(idxArt).CodVariante(idxVariante))
                        '    If File.Exists(srcFilename) = True Then
                        '        ' carica immagine da file
                        '        'html += vbTab + TAG_CELL_VARIANTE + "<img style='display:block; width:100px;height:100px;' id='base64image' src='data:image/jpeg;base64, " & convertFileToTobase64(srcFilename) & "' />" + TAG_CELL_END

                        '        ' ho tolto la forzatura alla dimensione 100x100 dell'immagine
                        '        strvAl += "immagine Variante"
                        '        xlWorkSheet.Cells(idxRow, 3) = strvAl
                        '    Else
                        '        strvAl += "no immagineVariante"
                        '        xlWorkSheet.Cells(idxRow, 3) = strvAl
                        '        html += vbTab + TAG_CELL_VARIANTE + "" + TAG_CELL_END
                        '    End If

                        'End If

                        ' la riga commentata esportava solo il codice variante quella scommentata esporta codice + descrizione
                        ' html += vbTab + TAG_CELL_VARIANTE + "<font size=""3"">" + varTag(idxArt).CodVariante(idxVariante) + "</font>" + TAG_CELL_END
                        mypoints = 0
                        stampaImgVariante(varTag(idxArt).CodArt, varTag(idxArt).CodVariante(idxVariante), varTag(idxArt).DescStag, idxRow, xlWorkSheet, mypoints)

                        xlWorkSheet.Rows(idxRow).RowHeight = 120
                        xlWorkSheet.Cells(idxRow, 2) = varTag(idxArt).Varianti(idxVariante)
                        xlWorkSheet.Cells(idxRow, 2).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan)
                        xlWorkSheet.Cells(idxRow, 2).font.bold = True
                        xlWorkSheet.Cells(idxRow, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                        xlWorkSheet.Cells(idxRow, 2).HorizontalAlignment = Excel.Constants.xlCenter
                        xlWorkSheet.Cells(idxRow, 2).VerticalAlignment = Excel.Constants.xlCenter

                        'idxRow = idxRow + 1

                        If bTaglieSempre = True Or indexPrimaRigaTaglia = 0 Then
                            ' controlla se mostrare l'immagine variante
                            xlWorkSheet.Cells(idxRow, 3) = "UM"
                            xlWorkSheet.Cells(idxRow, 4) = "TOT"
                        Else
                            xlWorkSheet.Cells(idxRow, 3) = ""
                            xlWorkSheet.Cells(idxRow, 4) = ""

                        End If
                        xlWorkSheet.Cells(idxRow, 3).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan)
                        xlWorkSheet.Cells(idxRow, 4).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan)

                        xlWorkSheet.Cells(idxRow, 3).font.bold = True
                        xlWorkSheet.Cells(idxRow, 4).font.bold = True

                        xlWorkSheet.Cells(idxRow, 3).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                        xlWorkSheet.Cells(idxRow, 4).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

                        xlWorkSheet.Cells(idxRow, 3).HorizontalAlignment = Excel.Constants.xlCenter
                        xlWorkSheet.Cells(idxRow, 4).HorizontalAlignment = Excel.Constants.xlCenter

                        xlWorkSheet.Cells(idxRow, 3).VerticalAlignment = Excel.Constants.xlCenter
                        xlWorkSheet.Cells(idxRow, 4).VerticalAlignment = Excel.Constants.xlCenter


                        If (stampaTestata = True And testgiastampata = False) Then
                            idxRow = idxRow - 1
                            strvAl = ""


                            If FormExport.CheckBoxArtCodice.Checked = True Then
                                If FormExport.CheckBoxLingua.Checked = True Then
                                    strvAl += "ART: " + varTag(idxArt).CodArt + " - " + varTag(idxArt).DescArt + " - " + varTag(idxArt).DescrInglese & vbCrLf & vbCrLf
                                Else
                                    strvAl += "ART: " + varTag(idxArt).CodArt + " - " + varTag(idxArt).DescArt + " - " + varTag(idxArt).DescEstesa & vbCrLf & vbCrLf
                                End If
                            End If

                            'Stagione
                            If FormExport.CheckBoxArtStagione.Checked = True Then
                                If FormExport.CheckBoxLingua.Checked = True Then
                                    strvAl += "STAG: " + varTag(idxArt).CodStag & " - " & varTag(idxArt).DescStag & vbCrLf & vbCrLf
                                Else
                                    strvAl += "SEASON: " + varTag(idxArt).CodStag & " - " & varTag(idxArt).DescStag & vbCrLf & vbCrLf
                                End If

                            End If

                            'Famiglia
                            If FormExport.CheckBoxArtFamiglia.Checked = True Then
                                strvAl += "FAM: " + varTag(idxArt).Famiglia & vbCrLf & vbCrLf
                            End If

                            'Famiglia
                            If FormExport.CheckBoxArtComposizione.Checked = True Then
                                strvAl += "COMP: " + varTag(idxArt).Composizione & vbCrLf & vbCrLf
                            End If

                            'Marca
                            If FormExport.CheckBoxMarca.Checked = True Then
                                If FormExport.CheckBoxLingua.Checked = True Then
                                    strvAl += "LINEA: " + varTag(idxArt).CodMarca + " - " + varTag(idxArt).DesMarca & vbCrLf & vbCrLf
                                Else
                                    strvAl += "LINE: " + varTag(idxArt).CodMarca + " - " + varTag(idxArt).DesMarca & vbCrLf & vbCrLf
                                End If
                            End If

                            'Nomenclatura
                            If FormExport.CheckBoxCodNomenclatura.Checked = True Then
                                If FormExport.CheckBoxLingua.Checked = True Then
                                    strvAl += "COD NOMENCLATURA: " + varTag(idxArt).CodNomenclatura & vbCrLf & vbCrLf
                                Else
                                    strvAl += "HTS CODE: " + varTag(idxArt).CodNomenclatura & vbCrLf & vbCrLf
                                End If
                            End If

                            'MadeIn
                            If versione <> VER_35 Then
                                If FormExport.CheckBoxMadeIn.Checked = True Then
                                    strvAl += "MADE IN: " + varTag(idxArt).MadeIn & vbCrLf & vbCrLf
                                End If
                            End If

                            xlWorkSheet.Cells(idxRow, 2) = strvAl
                            xlWorkSheet.Range("B" & idxRow).ColumnWidth = 40
                            xlWorkSheet.Cells(idxRow, 2).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan)
                            xlWorkSheet.Cells(idxRow, 2).VerticalAlignment = Excel.Constants.xlCenter
                            xlWorkSheet.Cells(idxRow, 2).HorizontalAlignment = Excel.Constants.xlLeft
                            xlWorkSheet.Cells(idxRow, 2).font.bold = True
                            xlWorkSheet.Cells(idxRow, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

                            stampaTestata = True
                            testgiastampata = True

                            idxRow = idxRow + 1
                        End If

                        If varTag(idxArt).haTaglie = True Then
                            If Not varTag(idxArt).Giacenze(0).Giacenze Is Nothing Then

                                'intestazione taglie
                                Dim idxColonna As Integer = 5
                                For idxTaglia = 0 To UBound(varTag(idxArt).Giacenze(0).Taglie) - 1
                                    If Array.IndexOf(escludiTg, varTag(idxArt).Giacenze(idxVariante).Taglie(idxTaglia).ToString) >= 0 Then
                                        Continue For
                                    End If

                                    If bTaglieSempre = True Or indexPrimaRigaTaglia = 0 Then
                                        xlWorkSheet.Cells(idxRow, idxColonna) = varTag(idxArt).Giacenze(idxVariante).Taglie(idxTaglia).ToString

                                    Else
                                        xlWorkSheet.Cells(idxRow, idxColonna) = ""
                                    End If
                                    xlWorkSheet.Rows(idxRow).RowHeight = 120
                                    xlWorkSheet.Cells(idxRow, idxColonna).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan)
                                    xlWorkSheet.Cells(idxRow, idxColonna).font.bold = True
                                    xlWorkSheet.Cells(idxRow, idxColonna).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

                                    xlWorkSheet.Cells(idxRow, idxColonna).HorizontalAlignment = Excel.Constants.xlCenter
                                    xlWorkSheet.Cells(idxRow, idxColonna).VerticalAlignment = Excel.Constants.xlCenter

                                    idxColonna = idxColonna + 1
                                    Application.DoEvents()
                                Next
                                indexPrimaRigaTaglia = indexPrimaRigaTaglia + 1

                                If showGiacenze = True Then
                                    'If (exportGiacZero = False) And (varTag(idxArt).TotaleGiac(idxVariante) = 0) Then
                                    If (exportGiacZero = False) And (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) = 0) Then

                                    Else
                                        'If (exportGiacMinZero = False) And (varTag(idxArt).TotaleGiac(idxVariante) < 0) Then
                                        If (exportGiacMinZero = False) And (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) < 0) Then

                                        Else
                                            bEsportaVariante = True
                                            ' giacenze
                                            ' TODO VERIFICARE IMMAGINE VARIANTE
                                            'If bShowImgVariante = True Then
                                            '    html += vbTab + TAG_CELL_VARIANTE + "" + TAG_CELL_END
                                            'End If
                                            If (stampaTestata = True And testgiastampata = False) Then
                                                If FormExport.CheckBoxArtCodice.Checked = True Then
                                                    If FormExport.CheckBoxLingua.Checked = True Then
                                                        strvAl += "ART: " + varTag(idxArt).CodArt + " - " + varTag(idxArt).DescrInglese & vbCrLf
                                                    Else
                                                        strvAl += "ART: " + varTag(idxArt).CodArt + " - " + varTag(idxArt).DescArt + " - " + varTag(idxArt).DescEstesa & vbCrLf
                                                    End If
                                                End If

                                                'Stagione
                                                If FormExport.CheckBoxArtStagione.Checked = True Then
                                                    If FormExport.CheckBoxLingua.Checked = True Then
                                                        strvAl += "SEASON: " + varTag(idxArt).CodStag & " - " & varTag(idxArt).DescStag & vbCrLf
                                                    Else
                                                        strvAl += "STAG: " + varTag(idxArt).CodStag & " - " & varTag(idxArt).DescStag & vbCrLf
                                                    End If

                                                End If

                                                'Famiglia
                                                If FormExport.CheckBoxArtFamiglia.Checked = True Then
                                                    strvAl += "FAM: " + varTag(idxArt).Famiglia & vbCrLf
                                                End If

                                                'Famiglia
                                                If FormExport.CheckBoxArtComposizione.Checked = True Then
                                                    strvAl += "COMP: " + varTag(idxArt).Composizione & vbCrLf
                                                End If

                                                'Marca
                                                If FormExport.CheckBoxMarca.Checked = True Then
                                                    If FormExport.CheckBoxLingua.Checked = True Then
                                                        strvAl += "LINE: " + varTag(idxArt).CodMarca + " - " + varTag(idxArt).DesMarca & vbCrLf
                                                    Else
                                                        strvAl += "LINEA: " + varTag(idxArt).CodMarca + " - " + varTag(idxArt).DesMarca & vbCrLf
                                                    End If
                                                End If

                                                'Nomenclatura
                                                If FormExport.CheckBoxCodNomenclatura.Checked = True Then
                                                    If FormExport.CheckBoxLingua.Checked = True Then
                                                        strvAl += "HS CODE: " + varTag(idxArt).CodNomenclatura & vbCrLf
                                                    Else
                                                        strvAl += "COD NOMENCLATURA: " + varTag(idxArt).CodNomenclatura & vbCrLf
                                                    End If
                                                End If

                                                'MadeIn
                                                If versione <> VER_35 Then
                                                    If FormExport.CheckBoxMadeIn.Checked = True Then
                                                        strvAl += "MADE IN: " + varTag(idxArt).MadeIn & vbCrLf
                                                    End If
                                                End If

                                                xlWorkSheet.Cells(idxRow, 2) = strvAl
                                                xlWorkSheet.Range("B" & idxRow).ColumnWidth = 40
                                                xlWorkSheet.Cells(idxRow, 2).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan)
                                                xlWorkSheet.Cells(idxRow, 2).VerticalAlignment = Excel.Constants.xlCenter
                                                xlWorkSheet.Cells(idxRow, 2).HorizontalAlignment = Excel.Constants.xlLeft
                                                xlWorkSheet.Cells(idxRow, 2).font.bold = True
                                                xlWorkSheet.Cells(idxRow, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

                                                stampaTestata = True
                                                testgiastampata = True
                                            End If


                                            idxRow = idxRow + 1
                                            xlWorkSheet.Rows(idxRow).RowHeight = 30
                                            xlWorkSheet.Cells(idxRow, 2) = LabelGiac
                                            xlWorkSheet.Cells(idxRow, 3) = varTag(idxArt).UM
                                            xlWorkSheet.Cells(idxRow, 4) = ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante)

                                            xlWorkSheet.Cells(idxRow, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                            xlWorkSheet.Cells(idxRow, 3).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                            xlWorkSheet.Cells(idxRow, 4).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

                                            idxColonna = 5
                                            For idxTaglia = 0 To UBound(varTag(idxArt).Giacenze(0).Giacenze) - 1
                                                If Array.IndexOf(escludiTg, varTag(idxArt).Giacenze(idxVariante).Taglie(idxTaglia).ToString) >= 0 Then
                                                    Continue For
                                                End If

                                                strOutVal = ""
                                                If varTag(idxArt).Giacenze(idxVariante).Giacenze(idxTaglia) <> 0 Then
                                                    strOutVal = varTag(idxArt).Giacenze(idxVariante).Giacenze(idxTaglia).ToString
                                                End If

                                                'controlla se deve esportare un massimo
                                                If FormExport.CheckBoxMaxQta.Checked = True Then
                                                    If strOutVal <> "" Then
                                                        If Int(strOutVal) >= FormExport.NumericUpDownMaxQta.Value Then
                                                            strOutVal = FormExport.NumericUpDownMaxQta.Value.ToString
                                                        End If
                                                    End If
                                                End If


                                                xlWorkSheet.Rows(idxRow).RowHeight = 30
                                                xlWorkSheet.Cells(idxRow, idxColonna) = strOutVal
                                                xlWorkSheet.Cells(idxRow, idxColonna).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

                                                idxColonna = idxColonna + 1
                                                Application.DoEvents()
                                            Next
                                        End If

                                    End If
                                End If


                                If showCalcolato = True Then
                                    If (exportGiacZero = False) And (varTag(idxArt).TotaleCalc(idxVariante) = 0) Then

                                    Else
                                        'If (exportGiacMinZero = False) And (varTag(idxArt).TotaleCalc(idxVariante) < 0) Then
                                        If (exportGiacMinZero = False) And (ricalcolatotaleCalcPerEsportazione(idxArt, idxVariante) < 0) Then

                                        Else
                                            bEsportaVariante = True

                                            If (stampaTestata = True And testgiastampata = False) Then



                                                idxRow = idxRow - 1
                                                'codice
                                                strvAl = ""
                                                If FormExport.CheckBoxArtCodice.Checked = True Then
                                                    If FormExport.CheckBoxLingua.Checked = True Then
                                                        strvAl += "ART: " + varTag(idxArt).CodArt + " - " + varTag(idxArt).DescrInglese & vbCrLf
                                                    Else
                                                        strvAl += "ART: " + varTag(idxArt).CodArt + " - " + varTag(idxArt).DescArt + " - " + varTag(idxArt).DescEstesa & vbCrLf
                                                    End If
                                                End If

                                                'Stagione
                                                If FormExport.CheckBoxArtStagione.Checked = True Then
                                                    If FormExport.CheckBoxLingua.Checked = True Then
                                                        strvAl += "SEASON: " + varTag(idxArt).CodStag & " - " & varTag(idxArt).DescStag & vbCrLf
                                                    Else
                                                        strvAl += "STAG: " + varTag(idxArt).CodStag & " - " & varTag(idxArt).DescStag & vbCrLf
                                                    End If

                                                End If

                                                'Famiglia
                                                If FormExport.CheckBoxArtFamiglia.Checked = True Then
                                                    strvAl += "FAM: " + varTag(idxArt).Famiglia & vbCrLf
                                                End If

                                                'Famiglia
                                                If FormExport.CheckBoxArtComposizione.Checked = True Then
                                                    strvAl += "COMP: " + varTag(idxArt).Composizione & vbCrLf
                                                End If

                                                'Marca
                                                If FormExport.CheckBoxMarca.Checked = True Then
                                                    If FormExport.CheckBoxLingua.Checked = True Then
                                                        strvAl += "LINE: " + varTag(idxArt).CodMarca + " - " + varTag(idxArt).DesMarca & vbCrLf
                                                    Else
                                                        strvAl += "LINEA: " + varTag(idxArt).CodMarca + " - " + varTag(idxArt).DesMarca & vbCrLf
                                                    End If
                                                End If

                                                'Nomenclatura
                                                If FormExport.CheckBoxCodNomenclatura.Checked = True Then
                                                    If FormExport.CheckBoxLingua.Checked = True Then
                                                        strvAl += "HS CODE: " + varTag(idxArt).CodNomenclatura & vbCrLf
                                                    Else
                                                        strvAl += "COD NOMENCLATURA: " + varTag(idxArt).CodNomenclatura & vbCrLf
                                                    End If
                                                End If

                                                'MadeIn
                                                If versione <> VER_35 Then
                                                    If FormExport.CheckBoxMadeIn.Checked = True Then
                                                        strvAl += "MADE IN: " + varTag(idxArt).MadeIn & vbCrLf
                                                    End If
                                                End If

                                                xlWorkSheet.Cells(idxRow, 2) = strvAl
                                                xlWorkSheet.Range("B" & idxRow).ColumnWidth = 40
                                                xlWorkSheet.Cells(idxRow, 2).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan)
                                                xlWorkSheet.Cells(idxRow, 2).VerticalAlignment = Excel.Constants.xlCenter
                                                xlWorkSheet.Cells(idxRow, 2).HorizontalAlignment = Excel.Constants.xlLeft
                                                xlWorkSheet.Cells(idxRow, 2).font.bold = True
                                                xlWorkSheet.Cells(idxRow, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

                                                stampaTestata = False
                                                testgiastampata = True
                                            End If

                                            ' calcolato - TODO verificare immagini
                                            'html += vbTab + TAG_ROW
                                            'If bShowImgVariante = True Then
                                            '    html += vbTab + TAG_CELL_VARIANTE + "" + TAG_CELL_END
                                            'End If
                                            idxRow = idxRow + 1
                                            xlWorkSheet.Rows(idxRow).RowHeight = 30
                                            xlWorkSheet.Cells(idxRow, 2) = labelCalc
                                            xlWorkSheet.Cells(idxRow, 3) = varTag(idxArt).UM
                                            xlWorkSheet.Cells(idxRow, 4) = ricalcolatotaleCalcPerEsportazione(idxArt, idxVariante)

                                            xlWorkSheet.Cells(idxRow, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                            xlWorkSheet.Cells(idxRow, 3).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                            xlWorkSheet.Cells(idxRow, 4).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

                                            idxColonna = 5
                                            For idxTaglia = 0 To UBound(varTag(idxArt).Giacenze(0).Calcolato) - 1
                                                If Array.IndexOf(escludiTg, varTag(idxArt).Giacenze(idxVariante).Taglie(idxTaglia).ToString) >= 0 Then
                                                    Continue For
                                                End If
                                                strOutVal = ""
                                                If varTag(idxArt).Giacenze(idxVariante).Calcolato(idxTaglia) <> 0 Then
                                                    strOutVal = varTag(idxArt).Giacenze(idxVariante).Calcolato(idxTaglia).ToString
                                                End If

                                                'controlla se deve esportare un massimo
                                                If FormExport.CheckBoxMaxQta.Checked = True Then
                                                    If strOutVal <> "" Then
                                                        If Int(strOutVal) >= FormExport.NumericUpDownMaxQta.Value Then
                                                            strOutVal = FormExport.NumericUpDownMaxQta.Value.ToString
                                                        End If
                                                    End If
                                                End If
                                                xlWorkSheet.Rows(idxRow).RowHeight = 30
                                                xlWorkSheet.Cells(idxRow, idxColonna) = strOutVal
                                                xlWorkSheet.Cells(idxRow, idxColonna).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

                                                idxColonna = idxColonna + 1
                                                Application.DoEvents()
                                            Next
                                        End If
                                    End If
                                End If

                            End If
                        Else
                            If showGiacenze = True Then
                                'If (exportGiacZero = False) And (varTag(idxArt).TotaleGiac(idxVariante) = 0) Then
                                If (exportGiacZero = False) And (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) = 0) Then

                                Else
                                    'If (exportGiacMinZero = False) And (varTag(idxArt).TotaleGiac(idxVariante) < 0) Then
                                    If (exportGiacMinZero = False) And (ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante) < 0) Then

                                    Else
                                        bEsportaVariante = True
                                        ' giacenze - TODO VERIFICARE IMMAGINE PER VARIANTE
                                        'html += vbTab + TAG_ROW
                                        'If bShowImgVariante = True Then
                                        '    html += vbTab + TAG_CELL_VARIANTE + "" + TAG_CELL_END
                                        'End If

                                        If (stampaTestata = True And testgiastampata = False) Then


                                            idxRow = idxRow - 1
                                            'codice
                                            strvAl = ""
                                            If FormExport.CheckBoxArtCodice.Checked = True Then
                                                If FormExport.CheckBoxLingua.Checked = True Then
                                                    strvAl += "ART: " + varTag(idxArt).CodArt + " - " + varTag(idxArt).DescrInglese & vbCrLf
                                                Else
                                                    strvAl += "ART: " + varTag(idxArt).CodArt + " - " + varTag(idxArt).DescArt + " - " + varTag(idxArt).DescEstesa & vbCrLf
                                                End If
                                            End If

                                            'Stagione
                                            If FormExport.CheckBoxArtStagione.Checked = True Then
                                                If FormExport.CheckBoxLingua.Checked = True Then
                                                    strvAl += "SEASON: " + varTag(idxArt).CodStag & " - " & varTag(idxArt).DescStag & vbCrLf
                                                Else
                                                    strvAl += "STAG: " + varTag(idxArt).CodStag & " - " & varTag(idxArt).DescStag & vbCrLf
                                                End If

                                            End If

                                            'Famiglia
                                            If FormExport.CheckBoxArtFamiglia.Checked = True Then
                                                strvAl += "FAM: " + varTag(idxArt).Famiglia & vbCrLf
                                            End If

                                            'Famiglia
                                            If FormExport.CheckBoxArtComposizione.Checked = True Then
                                                strvAl += "COMP: " + varTag(idxArt).Composizione & vbCrLf
                                            End If

                                            'Marca
                                            If FormExport.CheckBoxMarca.Checked = True Then
                                                If FormExport.CheckBoxLingua.Checked = True Then
                                                    strvAl += "LINE: " + varTag(idxArt).CodMarca + " - " + varTag(idxArt).DesMarca & vbCrLf
                                                Else
                                                    strvAl += "LINEA: " + varTag(idxArt).CodMarca + " - " + varTag(idxArt).DesMarca & vbCrLf
                                                End If
                                            End If

                                            'Nomenclatura
                                            If FormExport.CheckBoxCodNomenclatura.Checked = True Then
                                                If FormExport.CheckBoxLingua.Checked = True Then
                                                    strvAl += "HS CODE: " + varTag(idxArt).CodNomenclatura & vbCrLf
                                                Else
                                                    strvAl += "COD NOMENCLATURA: " + varTag(idxArt).CodNomenclatura & vbCrLf
                                                End If
                                            End If

                                            'MadeIn
                                            If versione <> VER_35 Then
                                                If FormExport.CheckBoxMadeIn.Checked = True Then
                                                    strvAl += "MADE IN: " + varTag(idxArt).MadeIn & vbCrLf
                                                End If
                                            End If


                                            xlWorkSheet.Cells(idxRow, 2) = strvAl
                                            xlWorkSheet.Range("B" & idxRow).ColumnWidth = 40
                                            xlWorkSheet.Cells(idxRow, 2).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan)
                                            xlWorkSheet.Cells(idxRow, 2).VerticalAlignment = Excel.Constants.xlCenter
                                            xlWorkSheet.Cells(idxRow, 2).HorizontalAlignment = Excel.Constants.xlLeft
                                            xlWorkSheet.Cells(idxRow, 2).font.bold = True
                                            xlWorkSheet.Cells(idxRow, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

                                            stampaTestata = False
                                            testgiastampata = True
                                        End If


                                        idxRow = idxRow + 2
                                        xlWorkSheet.Rows(idxRow).RowHeight = 30
                                        xlWorkSheet.Cells(idxRow, 1) = LabelGiac
                                        xlWorkSheet.Cells(idxRow, 2) = varTag(idxArt).UM
                                        xlWorkSheet.Cells(idxRow, 3) = ricalcolatotaleGiacPerEsportazione(idxArt, idxVariante)

                                        xlWorkSheet.Cells(idxRow, 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        xlWorkSheet.Cells(idxRow, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        xlWorkSheet.Cells(idxRow, 3).Borders.LineStyle = Excel.XlLineStyle.xlContinuous




                                    End If
                                End If
                            End If

                            If showCalcolato = True Then
                                If (exportGiacZero = False) And (varTag(idxArt).TotaleCalc(idxVariante) = 0) Then

                                Else
                                    'If (exportGiacMinZero = False) And (varTag(idxArt).TotaleCalc(idxVariante) < 0) Then
                                    If (exportGiacMinZero = False) And (ricalcolatotaleCalcPerEsportazione(idxArt, idxVariante) < 0) Then

                                    Else
                                        bEsportaVariante = True
                                        ' calcolato - TODO verificare immagine variante
                                        'html += vbTab + TAG_ROW
                                        'If bShowImgVariante = True Then
                                        '    html += vbTab + TAG_CELL_VARIANTE + "" + TAG_CELL_END
                                        'End If
                                        stampaTestata = True

                                        If (stampaTestata = True And testgiastampata = False) Then

                                            idxRow = idxRow - 1
                                            'codice
                                            strvAl = ""
                                            If FormExport.CheckBoxArtCodice.Checked = True Then
                                                If FormExport.CheckBoxLingua.Checked = True Then
                                                    strvAl += "ART: " + varTag(idxArt).CodArt + " - " + varTag(idxArt).DescrInglese & vbCrLf
                                                Else
                                                    strvAl += "ART: " + varTag(idxArt).CodArt + " - " + varTag(idxArt).DescArt + " - " + varTag(idxArt).DescEstesa & vbCrLf
                                                End If
                                            End If

                                            'Stagione
                                            If FormExport.CheckBoxArtStagione.Checked = True Then
                                                If FormExport.CheckBoxLingua.Checked = True Then
                                                    strvAl += "SEASON: " + varTag(idxArt).CodStag & " - " & varTag(idxArt).DescStag & vbCrLf
                                                Else
                                                    strvAl += "STAG: " + varTag(idxArt).CodStag & " - " & varTag(idxArt).DescStag & vbCrLf
                                                End If

                                            End If

                                            'Famiglia
                                            If FormExport.CheckBoxArtFamiglia.Checked = True Then
                                                strvAl += "FAM: " + varTag(idxArt).Famiglia & vbCrLf
                                            End If

                                            'Famiglia
                                            If FormExport.CheckBoxArtComposizione.Checked = True Then
                                                strvAl += "COMP: " + varTag(idxArt).Composizione & vbCrLf
                                            End If

                                            'Marca
                                            If FormExport.CheckBoxMarca.Checked = True Then
                                                If FormExport.CheckBoxLingua.Checked = True Then
                                                    strvAl += "LINE: " + varTag(idxArt).CodMarca + " - " + varTag(idxArt).DesMarca & vbCrLf
                                                Else
                                                    strvAl += "LINEA: " + varTag(idxArt).CodMarca + " - " + varTag(idxArt).DesMarca & vbCrLf
                                                End If
                                            End If

                                            'Nomenclatura
                                            If FormExport.CheckBoxCodNomenclatura.Checked = True Then
                                                If FormExport.CheckBoxLingua.Checked = True Then
                                                    strvAl += "HS CODE: " + varTag(idxArt).CodNomenclatura & vbCrLf
                                                Else
                                                    strvAl += "COD NOMENCLATURA: " + varTag(idxArt).CodNomenclatura & vbCrLf
                                                End If
                                            End If

                                            'MadeIn
                                            If versione <> VER_35 Then
                                                If FormExport.CheckBoxMadeIn.Checked = True Then
                                                    strvAl += "MADE IN: " + varTag(idxArt).MadeIn & vbCrLf
                                                End If
                                            End If

                                            xlWorkSheet.Cells(idxRow, 2) = strvAl
                                            xlWorkSheet.Range("B" & idxRow).ColumnWidth = 40
                                            xlWorkSheet.Cells(idxRow, 2).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan)
                                            xlWorkSheet.Cells(idxRow, 2).VerticalAlignment = Excel.Constants.xlCenter
                                            xlWorkSheet.Cells(idxRow, 2).HorizontalAlignment = Excel.Constants.xlLeft
                                            xlWorkSheet.Cells(idxRow, 2).font.bold = True
                                            xlWorkSheet.Cells(idxRow, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

                                            stampaTestata = False
                                            testgiastampata = True
                                        End If




                                        idxRow = idxRow + 2
                                        xlWorkSheet.Rows(idxRow).RowHeight = 30
                                        xlWorkSheet.Cells(idxRow, 1) = labelCalc
                                        xlWorkSheet.Cells(idxRow, 2) = varTag(idxArt).UM
                                        xlWorkSheet.Cells(idxRow, 3) = ricalcolatotaleCalcPerEsportazione(idxArt, idxVariante)

                                        xlWorkSheet.Cells(idxRow, 1).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        xlWorkSheet.Cells(idxRow, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        xlWorkSheet.Cells(idxRow, 3).Borders.LineStyle = Excel.XlLineStyle.xlContinuous




                                    End If
                                End If
                            End If
                            indexPrimaRigaTaglia = indexPrimaRigaTaglia + 1

                        End If

                        If bEsportaVariante = False Then
                            ' ripristino il vecchio valore in quanto non è stata esportata nessuna riga
                            indexPrimaRigaTaglia = indexPrimaRigaTaglia - 1
                        Else
                            ' se ha esportato almeno una volta lascia l'articolo
                            bEsportaArt = True
                            stampaTestata = True

                            If (stampaTestata = True And testgiastampata = False) Then

                                idxRow = idxRow - 2
                                'codice
                                strvAl = ""
                                If FormExport.CheckBoxArtCodice.Checked = True Then
                                    If FormExport.CheckBoxLingua.Checked = True Then
                                        strvAl += "ART: " + varTag(idxArt).CodArt + " - " + varTag(idxArt).DescrInglese & vbCrLf
                                    Else
                                        strvAl += "ART: " + varTag(idxArt).CodArt + " - " + varTag(idxArt).DescArt + " - " + varTag(idxArt).DescEstesa & vbCrLf
                                    End If
                                End If

                                'Stagione
                                If FormExport.CheckBoxArtStagione.Checked = True Then
                                    If FormExport.CheckBoxLingua.Checked = True Then
                                        strvAl += "SEASON: " + varTag(idxArt).CodStag & " - " & varTag(idxArt).DescStag & vbCrLf
                                    Else
                                        strvAl += "STAG: " + varTag(idxArt).CodStag & " - " & varTag(idxArt).DescStag & vbCrLf
                                    End If

                                End If

                                'Famiglia
                                If FormExport.CheckBoxArtFamiglia.Checked = True Then
                                    strvAl += "FAM: " + varTag(idxArt).Famiglia & vbCrLf
                                End If

                                'Famiglia
                                If FormExport.CheckBoxArtComposizione.Checked = True Then
                                    strvAl += "COMP: " + varTag(idxArt).Composizione & vbCrLf
                                End If

                                'Marca
                                If FormExport.CheckBoxMarca.Checked = True Then
                                    If FormExport.CheckBoxLingua.Checked = True Then
                                        strvAl += "LINE: " + varTag(idxArt).CodMarca + " - " + varTag(idxArt).DesMarca & vbCrLf
                                    Else
                                        strvAl += "LINEA: " + varTag(idxArt).CodMarca + " - " + varTag(idxArt).DesMarca & vbCrLf
                                    End If
                                End If

                                'Nomenclatura
                                If FormExport.CheckBoxCodNomenclatura.Checked = True Then
                                    If FormExport.CheckBoxLingua.Checked = True Then
                                        strvAl += "HS CODE: " + varTag(idxArt).CodNomenclatura & vbCrLf
                                    Else
                                        strvAl += "COD NOMENCLATURA: " + varTag(idxArt).CodNomenclatura & vbCrLf
                                    End If
                                End If

                                'MadeIn
                                If versione <> VER_35 Then
                                    If FormExport.CheckBoxMadeIn.Checked = True Then
                                        strvAl += "MADE IN: " + varTag(idxArt).MadeIn & vbCrLf
                                    End If
                                End If

                                xlWorkSheet.Cells(idxRow, 2) = strvAl
                                xlWorkSheet.Range("B" & idxRow).ColumnWidth = 40
                                xlWorkSheet.Cells(idxRow, 2).interior.color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan)
                                xlWorkSheet.Cells(idxRow, 2).VerticalAlignment = Excel.Constants.xlCenter
                                xlWorkSheet.Cells(idxRow, 2).HorizontalAlignment = Excel.Constants.xlLeft
                                xlWorkSheet.Cells(idxRow, 2).font.bold = True
                                xlWorkSheet.Cells(idxRow, 2).Borders.LineStyle = Excel.XlLineStyle.xlContinuous

                                stampaTestata = False
                                testgiastampata = True
                            End If
                        End If
                        Application.DoEvents()



                    Next
                    idxRow = idxRow + 3
                End If

                'idxRow = idxRow + 3


                ' controlla se deve gestire l'interruzione pagina
                If FormExport.CheckBoxInterPag.Checked = True Then
                    If idxArticolo = FormExport.NumericUpDownInterrPag.Value Then
                        idxArticolo = 1
                    Else
                        idxArticolo = idxArticolo + 1
                    End If
                End If



                'idxArt = idxArt + 1
                Application.DoEvents()
            Next


            LabelDB.Text = "Creazione file in corso ..."
            Application.DoEvents()

            xlWorkBook.SaveAs(SaveFileDialog1.FileName)
            xlWorkBook.Close()
            xlApp.Quit()

            MsgBox("Esportazione avvenuta con successo!")
            LabelDB.Text = "Esportazione avvenuta con successo"


            Process.Start(SaveFileDialog1.FileName)

        End If

    End Sub

    Private Sub stampaImgVariante(codArt As String, Var As String, stag As String, row As Integer, xlsSheet As Excel.Worksheet, mypoints As Double)
        Dim largh As Double
        Dim alt As Double
        Dim file As String
        Dim shp As Excel.Shape = Nothing

        file = "Z:\varianti colore\" + codArt + "_" + Var + ".jpg"

        If System.IO.File.Exists(file) = False Then
            file = "Z:\varianti colore\cartella colori\" + Var + ".jpg"
            If System.IO.File.Exists(file) = False Then
                file = ""
            End If
        End If

        If file = "" Then
            file = getTmpImagefilename(My.Resources.url, largh, alt)
            'Else
            'file = getTmpImagefilenameFromFile(file, largh, alt)
        End If

        largh = 100
        alt = 100

        'If FormExport.CheckBoxFormatoImgFisso.Checked = True Then
        '    largh = 100
        '    alt = 100
        'Else
        '    If alt > MAX_ALTEZZA Then
        '        largh = (largh * MAX_ALTEZZA) / alt
        '        alt = MAX_ALTEZZA
        '    End If
        'End If

        Dim g As Graphics = CreateGraphics()
        Dim r As Excel.Range

        mypoints = largh * 18 / g.DpiX

        g.Dispose()

        'xlsSheet.Range("A" & row & ":A" & row + 2).MergeCells = True

        'xlsSheet.Range("A" & row).ColumnWidth = mypoints
        'xlsSheet.Range("A" & row).RowHeight = alt + 4
        r = xlsSheet.Cells(row, 1)

        'Dim left = 0
        shp = xlsSheet.Shapes.AddPicture(file, False, True, r.Left, r.Top, largh, alt)

        shp.Left = r.Left + (r.Width - shp.Width) / 2
        'shp.Top = r.Top + (r.Height - shp.Height) / 2

        'shp.Left = r.Left + 10
        shp.Top = r.Top + 10

    End Sub

    Private Function getTmpImagefilenameFromFile(imgName As String, ByRef imgWith As Integer, ByRef imgHeigth As Integer) As String

        Dim resizeVal As Integer
        Dim img As System.Drawing.Image = Image.FromFile(imgName)

        If img Is Nothing Then
            Return ""
        End If

        imgWith = img.Width
        imgHeigth = img.Height

        resizeVal = 1

        If img.Width > 100 Then
            resizeVal = 2
        End If
        If img.Width > 200 Then
            resizeVal = 4
        End If
        If img.Width > 300 Then
            resizeVal = 6
        End If
        If img.Width > 400 Then
            resizeVal = 8
        End If
        If img.Width > 800 Then
            resizeVal = 16
        End If
        If img.Width > 1000 Then
            resizeVal = 20
        End If

        ' riduce l'immagie di 10 volte per evitare che diventi troppo grande

        Dim tmpfilename As String = tmp_dir & Now.ToString("yyyyMMdd_HHmmss") & ".png" '"tmp.png"
        img = ResizeImage(img, (img.Width \ resizeVal), (img.Height \ resizeVal))
        'img.Save(tmpfilename)

        Return tmpfilename
    End Function

    Private Function getTmpImagefilename(img As System.Drawing.Bitmap, ByRef imgWith As Integer, ByRef imgHeigth As Integer) As String
        Dim resizeVal As Integer

        If img Is Nothing Then
            Return ""
        End If

        imgWith = img.Width
        imgHeigth = img.Height

        If img.Width > 100 Then
            resizeVal = 1
        End If
        If img.Width > 200 Then
            resizeVal = 2
        End If
        If img.Width > 300 Then
            resizeVal = 3
        End If
        If img.Width > 400 Then
            resizeVal = 4
        End If
        If img.Width > 800 Then
            resizeVal = 8
        End If
        If img.Width > 1000 Then
            resizeVal = 10
        End If

        ' riduce l'immagie di 10 volte per evitare che diventi troppo grande
        Dim tmpfilename As String = tmp_dir & "tmp.png"
        img = ResizeImage(img, (img.Width \ resizeVal), (img.Height \ resizeVal))
        img.Save(tmpfilename)

        Return tmpfilename

    End Function

    Private Sub ButtonExportEXCEL_Click(sender As Object, e As EventArgs) Handles ButtonExportEXCEL.Click
        If FormExport.ShowDialog = Windows.Forms.DialogResult.OK Then
            Call exporEXCEL(FormExport.CheckBoxGiac.Checked, FormExport.CheckBoxCalcolato.Checked, FormExport.CheckBoxImgArticolo.Checked, FormExport.TextBoxGiacenza.Text, FormExport.TextBoxDispTeorica.Text)
        End If
    End Sub

    Private Sub FormMain_MenuComplete(sender As Object, e As EventArgs) Handles Me.MenuComplete

    End Sub

End Class
