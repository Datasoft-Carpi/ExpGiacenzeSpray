Imports Microsoft.Win32


Public Class FormExport

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles ButtonCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub


    Private Sub ButtonOK_Click(sender As System.Object, e As System.EventArgs) Handles ButtonOK.Click
        Try
            'generali
            writeVal("EXP_GIAC", CheckBoxGiac.Checked)
            writeVal("EXP_CALC", CheckBoxCalcolato.Checked)
            writeVal("EXP_IMG", CheckBoxImgArticolo.Checked)
            writeVal("EXP_IMG_VAR", CheckBoxImgVariante.Checked)
            writeVal("EXP_TOT_ZERO", CheckBoxTotaliZero.Checked)
            writeVal("EXP_TOT_MIN_ZERO", CheckBoxTotaliMinZero.Checked)
            writeVal("LABEL_GIAC", TextBoxGiacenza.Text)
            writeVal("LABEL_CALC", TextBoxDispTeorica.Text)
            writeVal("ESCLUDI_VAR", TextBoxEscludiVarianti.Text)
            writeVal("ESCLUDI_TG", TextBoxEscludiTaglie.Text)
            writeVal("SCRIVI_TG", CheckBoxTaglie.Checked)

            'interruzione pagina
            writeVal("INTER_PAG", CheckBoxInterPag.Checked)
            writeVal("INTER_PAG_VAL", NumericUpDownInterrPag.Value)

            'massima quantità esportabile
            writeVal("MAX_QTA_CK", CheckBoxMaxQta.Checked)
            writeVal("MAX_QTA_VAL", NumericUpDownMaxQta.Value)

            'listini
            writeVal("EXP_LISACQ", CheckBoxExpLisAcquisto.Checked)
            writeVal("EXP_LISVEN", CheckBoxExpLisVendita.Checked)


            ' articolo
            writeVal("ART_COD", CheckBoxArtCodice.Checked)
            writeVal("ART_STAG", CheckBoxArtStagione.Checked)
            writeVal("ART_COMP", CheckBoxArtComposizione.Checked)
            writeVal("ART_FAM", CheckBoxArtFamiglia.Checked)
            writeVal("ART_NOMENCL", CheckBoxCodNomenclatura.Checked)
            writeVal("ART_MARCA", CheckBoxMarca.Checked)
            writeVal("ART_MADE", CheckBoxMadeIn.Checked)
            writeVal("DIM_IMG", CheckBoxFormatoImgFisso.Checked)

        Catch ex As Exception

        End Try

        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub FormExport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        If FormMain.CODLIS_VEN = "" Then
            CheckBoxExpLisVendita.Checked = False
            CheckBoxExpLisVendita.Visible = False
        End If

        If FormMain.CODLIS_ACQ = "" Then
            CheckBoxExpLisAcquisto.Checked = False
            CheckBoxExpLisAcquisto.Visible = False
        End If

        If FormMain.FOLDER_IMG_VAR = "" Then
            CheckBoxImgVariante.Checked = False
            CheckBoxImgVariante.Visible = False
        End If

    End Sub

    Private Sub FormExport_Shown(sender As System.Object, e As System.EventArgs) Handles MyBase.Shown
        Dim regKey As RegistryKey
        regKey = Registry.LocalMachine.OpenSubKey("Software\expGiacenze", True)
        If regKey Is Nothing Then
            Exit Sub
        End If

        Try
            ' generali
            CheckBoxGiac.Checked = Convert.ToBoolean(readVal("EXP_GIAC"))
            CheckBoxCalcolato.Checked = Convert.ToBoolean(readVal("EXP_CALC"))
            CheckBoxImgArticolo.Checked = Convert.ToBoolean(readVal("EXP_IMG"))
            CheckBoxImgVariante.Checked = Convert.ToBoolean(readVal("EXP_IMG_VAR"))
            CheckBoxTotaliZero.Checked = Convert.ToBoolean(readVal("EXP_TOT_ZERO"))
            CheckBoxTotaliMinZero.Checked = Convert.ToBoolean(readVal("EXP_TOT_MIN_ZERO"))

            NumericUpDownMaggioriDi.Value = 0

            TextBoxGiacenza.Text = readVal("LABEL_GIAC")
            TextBoxDispTeorica.Text = readVal("LABEL_CALC")
            TextBoxEscludiTaglie.Text = readVal("ESCLUDI_TG")
            TextBoxEscludiVarianti.Text = readVal("ESCLUDI_VAR")
            CheckBoxTaglie.Checked = Convert.ToBoolean(readVal("SCRIVI_TG"))

            ' articolo
            CheckBoxArtCodice.Checked = Convert.ToBoolean(readVal("ART_COD"))
            CheckBoxArtStagione.Checked = Convert.ToBoolean(readVal("ART_STAG"))
            CheckBoxArtComposizione.Checked = Convert.ToBoolean(readVal("ART_COMP"))
            CheckBoxArtFamiglia.Checked = Convert.ToBoolean(readVal("ART_FAM"))
            CheckBoxCodNomenclatura.Checked = Convert.ToBoolean(readVal("ART_NOMENCL"))
            CheckBoxMarca.Checked = Convert.ToBoolean(readVal("ART_MARCA"))
            CheckBoxMadeIn.Checked = Convert.ToBoolean(readVal("ART_MADE"))

            'listini
            CheckBoxExpLisAcquisto.Checked = Convert.ToBoolean(readVal("EXP_LISACQ"))
            CheckBoxExpLisVendita.Checked = Convert.ToBoolean(readVal("EXP_LISVEN"))

            CheckBoxFormatoImgFisso.Checked = Convert.ToBoolean(readVal("DIM_IMG"))

            ' interruzione pagina
            CheckBoxInterPag.Checked = Convert.ToBoolean(readVal("INTER_PAG"))
            NumericUpDownInterrPag.Value = readVal("INTER_PAG_VAL")


            'massima quantità esportabile
            CheckBoxMaxQta.Checked = Convert.ToBoolean(readVal("MAX_QTA_CK"))
            NumericUpDownMaxQta.Value = readVal("MAX_QTA_VAL")

        Catch ex As Exception

            TextBoxDispTeorica.Text = "Disp. Teorica"


        End Try


        If FormMain.versione = FormMain.VER_35 Then
            CheckBoxArtComposizione.Checked = False
            CheckBoxArtComposizione.Visible = False



            CheckBoxMadeIn.Checked = False
            CheckBoxMadeIn.Visible = False

        End If
        TextBoxDispTeorica.Text = "Disp. Teorica"

    End Sub

    Private Function readVal(Key As String) As String
        Dim regKey As RegistryKey
        Dim ver As String = ""
        regKey = Registry.LocalMachine.OpenSubKey("Software\expGiacenze", True)
        If Not regKey Is Nothing Then
            ver = regKey.GetValue(Key, "")
        End If
        regKey.Close()

        Return ver
    End Function

    Private Sub writeVal(Key As String, val As String)
        Dim regKey As RegistryKey

        regKey = Registry.LocalMachine.OpenSubKey("Software\expGiacenze", True)
        If regKey Is Nothing Then

            regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE", True)
            regKey.CreateSubKey("expGiacenze")

        End If
        regKey.SetValue(Key, val.ToString)
        regKey.Close()

    End Sub

End Class