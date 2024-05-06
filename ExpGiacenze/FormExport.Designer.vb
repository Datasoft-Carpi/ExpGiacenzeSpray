<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormExport
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
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

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla nell'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.CheckBoxGiac = New System.Windows.Forms.CheckBox()
        Me.CheckBoxCalcolato = New System.Windows.Forms.CheckBox()
        Me.ButtonOK = New System.Windows.Forms.Button()
        Me.ButtonCancel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBoxGiacenza = New System.Windows.Forms.TextBox()
        Me.TextBoxDispTeorica = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CheckBoxImgArticolo = New System.Windows.Forms.CheckBox()
        Me.CheckBoxTotaliZero = New System.Windows.Forms.CheckBox()
        Me.TextBoxEscludiVarianti = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBoxEscludiTaglie = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CheckBoxTaglie = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxMadeIn = New System.Windows.Forms.CheckBox()
        Me.CheckBoxCodNomenclatura = New System.Windows.Forms.CheckBox()
        Me.CheckBoxMarca = New System.Windows.Forms.CheckBox()
        Me.CheckBoxArtFamiglia = New System.Windows.Forms.CheckBox()
        Me.CheckBoxArtComposizione = New System.Windows.Forms.CheckBox()
        Me.CheckBoxArtStagione = New System.Windows.Forms.CheckBox()
        Me.CheckBoxArtCodice = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxLingua = New System.Windows.Forms.CheckBox()
        Me.NumericUpDownMaxQta = New System.Windows.Forms.NumericUpDown()
        Me.CheckBoxMaxQta = New System.Windows.Forms.CheckBox()
        Me.NumericUpDownInterrPag = New System.Windows.Forms.NumericUpDown()
        Me.CheckBoxInterPag = New System.Windows.Forms.CheckBox()
        Me.CheckBoxTotaliMinZero = New System.Windows.Forms.CheckBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxExpLisAcquisto = New System.Windows.Forms.CheckBox()
        Me.CheckBoxExpLisVendita = New System.Windows.Forms.CheckBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxFormatoImgFisso = New System.Windows.Forms.CheckBox()
        Me.CheckBoxImgVariante = New System.Windows.Forms.CheckBox()
        Me.NumericUpDownMaggioriDi = New System.Windows.Forms.NumericUpDown()
        Me.chkBuchiTaglia = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.NumericUpDownMaxQta, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDownInterrPag, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.NumericUpDownMaggioriDi, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CheckBoxGiac
        '
        Me.CheckBoxGiac.AutoSize = True
        Me.CheckBoxGiac.Checked = True
        Me.CheckBoxGiac.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxGiac.Location = New System.Drawing.Point(13, 30)
        Me.CheckBoxGiac.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxGiac.Name = "CheckBoxGiac"
        Me.CheckBoxGiac.Size = New System.Drawing.Size(204, 20)
        Me.CheckBoxGiac.TabIndex = 0
        Me.CheckBoxGiac.Text = "Esporta Giacenza magazzino"
        Me.CheckBoxGiac.UseVisualStyleBackColor = True
        '
        'CheckBoxCalcolato
        '
        Me.CheckBoxCalcolato.AutoSize = True
        Me.CheckBoxCalcolato.Checked = True
        Me.CheckBoxCalcolato.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxCalcolato.Location = New System.Drawing.Point(13, 58)
        Me.CheckBoxCalcolato.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxCalcolato.Name = "CheckBoxCalcolato"
        Me.CheckBoxCalcolato.Size = New System.Drawing.Size(160, 20)
        Me.CheckBoxCalcolato.TabIndex = 1
        Me.CheckBoxCalcolato.Text = "Esporta Disp. Teorica"
        Me.CheckBoxCalcolato.UseVisualStyleBackColor = True
        '
        'ButtonOK
        '
        Me.ButtonOK.Location = New System.Drawing.Point(829, 411)
        Me.ButtonOK.Margin = New System.Windows.Forms.Padding(4)
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.Size = New System.Drawing.Size(100, 28)
        Me.ButtonOK.TabIndex = 2
        Me.ButtonOK.Text = "OK"
        Me.ButtonOK.UseVisualStyleBackColor = True
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Location = New System.Drawing.Point(960, 411)
        Me.ButtonCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(100, 28)
        Me.ButtonCancel.TabIndex = 3
        Me.ButtonCancel.Text = "ANNULLA"
        Me.ButtonCancel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(595, 249)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(116, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Etichetta giacenza"
        '
        'TextBoxGiacenza
        '
        Me.TextBoxGiacenza.Location = New System.Drawing.Point(729, 240)
        Me.TextBoxGiacenza.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBoxGiacenza.Name = "TextBoxGiacenza"
        Me.TextBoxGiacenza.Size = New System.Drawing.Size(329, 22)
        Me.TextBoxGiacenza.TabIndex = 5
        Me.TextBoxGiacenza.Text = "Giacenza magazzino"
        '
        'TextBoxDispTeorica
        '
        Me.TextBoxDispTeorica.Location = New System.Drawing.Point(729, 281)
        Me.TextBoxDispTeorica.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBoxDispTeorica.Name = "TextBoxDispTeorica"
        Me.TextBoxDispTeorica.Size = New System.Drawing.Size(329, 22)
        Me.TextBoxDispTeorica.TabIndex = 7
        Me.TextBoxDispTeorica.Text = "Disp. Teorica"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(595, 284)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(116, 16)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Etichetta calcolato"
        '
        'CheckBoxImgArticolo
        '
        Me.CheckBoxImgArticolo.AutoSize = True
        Me.CheckBoxImgArticolo.Checked = True
        Me.CheckBoxImgArticolo.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxImgArticolo.Location = New System.Drawing.Point(21, 23)
        Me.CheckBoxImgArticolo.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxImgArticolo.Name = "CheckBoxImgArticolo"
        Me.CheckBoxImgArticolo.Size = New System.Drawing.Size(185, 20)
        Me.CheckBoxImgArticolo.TabIndex = 8
        Me.CheckBoxImgArticolo.Text = "Esporta Immagine articolo"
        Me.CheckBoxImgArticolo.UseVisualStyleBackColor = True
        '
        'CheckBoxTotaliZero
        '
        Me.CheckBoxTotaliZero.AutoSize = True
        Me.CheckBoxTotaliZero.Checked = True
        Me.CheckBoxTotaliZero.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxTotaliZero.Location = New System.Drawing.Point(12, 86)
        Me.CheckBoxTotaliZero.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxTotaliZero.Name = "CheckBoxTotaliZero"
        Me.CheckBoxTotaliZero.Size = New System.Drawing.Size(205, 20)
        Me.CheckBoxTotaliZero.TabIndex = 9
        Me.CheckBoxTotaliZero.Text = "Esporta righe con totali a zero"
        Me.CheckBoxTotaliZero.UseVisualStyleBackColor = True
        '
        'TextBoxEscludiVarianti
        '
        Me.TextBoxEscludiVarianti.Location = New System.Drawing.Point(599, 36)
        Me.TextBoxEscludiVarianti.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBoxEscludiVarianti.Multiline = True
        Me.TextBoxEscludiVarianti.Name = "TextBoxEscludiVarianti"
        Me.TextBoxEscludiVarianti.Size = New System.Drawing.Size(460, 84)
        Me.TextBoxEscludiVarianti.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(595, 16)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(400, 16)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Escludi Varianti (dividi i valori da escludere con la , es: val1,val2,...)"
        '
        'TextBoxEscludiTaglie
        '
        Me.TextBoxEscludiTaglie.Location = New System.Drawing.Point(596, 148)
        Me.TextBoxEscludiTaglie.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBoxEscludiTaglie.Multiline = True
        Me.TextBoxEscludiTaglie.Name = "TextBoxEscludiTaglie"
        Me.TextBoxEscludiTaglie.Size = New System.Drawing.Size(463, 84)
        Me.TextBoxEscludiTaglie.TabIndex = 13
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(592, 128)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(394, 16)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Escludi Taglie (dividi i valori da escludere con la , es: val1,val2,...)"
        '
        'CheckBoxTaglie
        '
        Me.CheckBoxTaglie.AutoSize = True
        Me.CheckBoxTaglie.Checked = True
        Me.CheckBoxTaglie.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxTaglie.Location = New System.Drawing.Point(12, 171)
        Me.CheckBoxTaglie.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxTaglie.Name = "CheckBoxTaglie"
        Me.CheckBoxTaglie.Size = New System.Drawing.Size(174, 20)
        Me.CheckBoxTaglie.TabIndex = 14
        Me.CheckBoxTaglie.Text = "Ripeti intestazione taglie"
        Me.CheckBoxTaglie.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CheckBoxMadeIn)
        Me.GroupBox1.Controls.Add(Me.CheckBoxCodNomenclatura)
        Me.GroupBox1.Controls.Add(Me.CheckBoxMarca)
        Me.GroupBox1.Controls.Add(Me.CheckBoxArtFamiglia)
        Me.GroupBox1.Controls.Add(Me.CheckBoxArtComposizione)
        Me.GroupBox1.Controls.Add(Me.CheckBoxArtStagione)
        Me.GroupBox1.Controls.Add(Me.CheckBoxArtCodice)
        Me.GroupBox1.Location = New System.Drawing.Point(333, 15)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(247, 294)
        Me.GroupBox1.TabIndex = 15
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Articolo"
        '
        'CheckBoxMadeIn
        '
        Me.CheckBoxMadeIn.AutoSize = True
        Me.CheckBoxMadeIn.Checked = True
        Me.CheckBoxMadeIn.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxMadeIn.Location = New System.Drawing.Point(21, 199)
        Me.CheckBoxMadeIn.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxMadeIn.Name = "CheckBoxMadeIn"
        Me.CheckBoxMadeIn.Size = New System.Drawing.Size(77, 20)
        Me.CheckBoxMadeIn.TabIndex = 6
        Me.CheckBoxMadeIn.Text = "Made in"
        Me.CheckBoxMadeIn.UseVisualStyleBackColor = True
        '
        'CheckBoxCodNomenclatura
        '
        Me.CheckBoxCodNomenclatura.AutoSize = True
        Me.CheckBoxCodNomenclatura.Checked = True
        Me.CheckBoxCodNomenclatura.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxCodNomenclatura.Location = New System.Drawing.Point(21, 171)
        Me.CheckBoxCodNomenclatura.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxCodNomenclatura.Name = "CheckBoxCodNomenclatura"
        Me.CheckBoxCodNomenclatura.Size = New System.Drawing.Size(156, 20)
        Me.CheckBoxCodNomenclatura.TabIndex = 5
        Me.CheckBoxCodNomenclatura.Text = "Codice nomenclatura"
        Me.CheckBoxCodNomenclatura.UseVisualStyleBackColor = True
        '
        'CheckBoxMarca
        '
        Me.CheckBoxMarca.AutoSize = True
        Me.CheckBoxMarca.Checked = True
        Me.CheckBoxMarca.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxMarca.Location = New System.Drawing.Point(21, 86)
        Me.CheckBoxMarca.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxMarca.Name = "CheckBoxMarca"
        Me.CheckBoxMarca.Size = New System.Drawing.Size(149, 20)
        Me.CheckBoxMarca.TabIndex = 4
        Me.CheckBoxMarca.Text = "Marca + descrizione"
        Me.CheckBoxMarca.UseVisualStyleBackColor = True
        '
        'CheckBoxArtFamiglia
        '
        Me.CheckBoxArtFamiglia.AutoSize = True
        Me.CheckBoxArtFamiglia.Checked = True
        Me.CheckBoxArtFamiglia.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxArtFamiglia.Location = New System.Drawing.Point(21, 143)
        Me.CheckBoxArtFamiglia.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxArtFamiglia.Name = "CheckBoxArtFamiglia"
        Me.CheckBoxArtFamiglia.Size = New System.Drawing.Size(81, 20)
        Me.CheckBoxArtFamiglia.TabIndex = 3
        Me.CheckBoxArtFamiglia.Text = "Famiglia"
        Me.CheckBoxArtFamiglia.UseVisualStyleBackColor = True
        '
        'CheckBoxArtComposizione
        '
        Me.CheckBoxArtComposizione.AutoSize = True
        Me.CheckBoxArtComposizione.Checked = True
        Me.CheckBoxArtComposizione.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxArtComposizione.Location = New System.Drawing.Point(21, 114)
        Me.CheckBoxArtComposizione.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxArtComposizione.Name = "CheckBoxArtComposizione"
        Me.CheckBoxArtComposizione.Size = New System.Drawing.Size(159, 20)
        Me.CheckBoxArtComposizione.TabIndex = 2
        Me.CheckBoxArtComposizione.Text = "Composizione estesa"
        Me.CheckBoxArtComposizione.UseVisualStyleBackColor = True
        '
        'CheckBoxArtStagione
        '
        Me.CheckBoxArtStagione.AutoSize = True
        Me.CheckBoxArtStagione.Checked = True
        Me.CheckBoxArtStagione.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxArtStagione.Location = New System.Drawing.Point(21, 58)
        Me.CheckBoxArtStagione.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxArtStagione.Name = "CheckBoxArtStagione"
        Me.CheckBoxArtStagione.Size = New System.Drawing.Size(167, 20)
        Me.CheckBoxArtStagione.TabIndex = 1
        Me.CheckBoxArtStagione.Text = "Stagione + Descrizione"
        Me.CheckBoxArtStagione.UseVisualStyleBackColor = True
        '
        'CheckBoxArtCodice
        '
        Me.CheckBoxArtCodice.AutoSize = True
        Me.CheckBoxArtCodice.Checked = True
        Me.CheckBoxArtCodice.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxArtCodice.Location = New System.Drawing.Point(21, 30)
        Me.CheckBoxArtCodice.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxArtCodice.Name = "CheckBoxArtCodice"
        Me.CheckBoxArtCodice.Size = New System.Drawing.Size(204, 20)
        Me.CheckBoxArtCodice.TabIndex = 0
        Me.CheckBoxArtCodice.Text = "Codice Articolo + Descrizione"
        Me.CheckBoxArtCodice.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkBuchiTaglia)
        Me.GroupBox2.Controls.Add(Me.CheckBoxLingua)
        Me.GroupBox2.Controls.Add(Me.NumericUpDownMaxQta)
        Me.GroupBox2.Controls.Add(Me.CheckBoxMaxQta)
        Me.GroupBox2.Controls.Add(Me.NumericUpDownInterrPag)
        Me.GroupBox2.Controls.Add(Me.CheckBoxInterPag)
        Me.GroupBox2.Controls.Add(Me.CheckBoxTotaliMinZero)
        Me.GroupBox2.Controls.Add(Me.CheckBoxGiac)
        Me.GroupBox2.Controls.Add(Me.CheckBoxCalcolato)
        Me.GroupBox2.Controls.Add(Me.CheckBoxTaglie)
        Me.GroupBox2.Controls.Add(Me.CheckBoxTotaliZero)
        Me.GroupBox2.Location = New System.Drawing.Point(16, 15)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Size = New System.Drawing.Size(300, 294)
        Me.GroupBox2.TabIndex = 16
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Generale"
        '
        'CheckBoxLingua
        '
        Me.CheckBoxLingua.AutoSize = True
        Me.CheckBoxLingua.Location = New System.Drawing.Point(13, 257)
        Me.CheckBoxLingua.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxLingua.Name = "CheckBoxLingua"
        Me.CheckBoxLingua.Size = New System.Drawing.Size(193, 20)
        Me.CheckBoxLingua.TabIndex = 20
        Me.CheckBoxLingua.Text = "Descrizioni in lingua italiano"
        Me.CheckBoxLingua.UseVisualStyleBackColor = True
        '
        'NumericUpDownMaxQta
        '
        Me.NumericUpDownMaxQta.Location = New System.Drawing.Point(224, 225)
        Me.NumericUpDownMaxQta.Margin = New System.Windows.Forms.Padding(4)
        Me.NumericUpDownMaxQta.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDownMaxQta.Name = "NumericUpDownMaxQta"
        Me.NumericUpDownMaxQta.Size = New System.Drawing.Size(67, 22)
        Me.NumericUpDownMaxQta.TabIndex = 19
        Me.NumericUpDownMaxQta.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'CheckBoxMaxQta
        '
        Me.CheckBoxMaxQta.AutoSize = True
        Me.CheckBoxMaxQta.Checked = True
        Me.CheckBoxMaxQta.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxMaxQta.Location = New System.Drawing.Point(13, 229)
        Me.CheckBoxMaxQta.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxMaxQta.Name = "CheckBoxMaxQta"
        Me.CheckBoxMaxQta.Size = New System.Drawing.Size(191, 20)
        Me.CheckBoxMaxQta.TabIndex = 18
        Me.CheckBoxMaxQta.Text = "Valore max in esportazione"
        Me.CheckBoxMaxQta.UseVisualStyleBackColor = True
        '
        'NumericUpDownInterrPag
        '
        Me.NumericUpDownInterrPag.Location = New System.Drawing.Point(241, 194)
        Me.NumericUpDownInterrPag.Margin = New System.Windows.Forms.Padding(4)
        Me.NumericUpDownInterrPag.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDownInterrPag.Name = "NumericUpDownInterrPag"
        Me.NumericUpDownInterrPag.Size = New System.Drawing.Size(49, 22)
        Me.NumericUpDownInterrPag.TabIndex = 17
        Me.NumericUpDownInterrPag.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'CheckBoxInterPag
        '
        Me.CheckBoxInterPag.AutoSize = True
        Me.CheckBoxInterPag.Checked = True
        Me.CheckBoxInterPag.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxInterPag.Location = New System.Drawing.Point(13, 198)
        Me.CheckBoxInterPag.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxInterPag.Name = "CheckBoxInterPag"
        Me.CheckBoxInterPag.Size = New System.Drawing.Size(206, 20)
        Me.CheckBoxInterPag.TabIndex = 16
        Me.CheckBoxInterPag.Text = "Inter. stampa pag. ogni articoli"
        Me.CheckBoxInterPag.UseVisualStyleBackColor = True
        '
        'CheckBoxTotaliMinZero
        '
        Me.CheckBoxTotaliMinZero.AutoSize = True
        Me.CheckBoxTotaliMinZero.Checked = True
        Me.CheckBoxTotaliMinZero.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxTotaliMinZero.Location = New System.Drawing.Point(12, 114)
        Me.CheckBoxTotaliMinZero.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxTotaliMinZero.Name = "CheckBoxTotaliMinZero"
        Me.CheckBoxTotaliMinZero.Size = New System.Drawing.Size(247, 20)
        Me.CheckBoxTotaliMinZero.TabIndex = 15
        Me.CheckBoxTotaliMinZero.Text = "Esporta righe con totali minori di zero"
        Me.CheckBoxTotaliMinZero.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.CheckBoxExpLisAcquisto)
        Me.GroupBox3.Controls.Add(Me.CheckBoxExpLisVendita)
        Me.GroupBox3.Location = New System.Drawing.Point(17, 317)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox3.Size = New System.Drawing.Size(299, 122)
        Me.GroupBox3.TabIndex = 17
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Listini"
        '
        'CheckBoxExpLisAcquisto
        '
        Me.CheckBoxExpLisAcquisto.AutoSize = True
        Me.CheckBoxExpLisAcquisto.Checked = True
        Me.CheckBoxExpLisAcquisto.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxExpLisAcquisto.Location = New System.Drawing.Point(12, 23)
        Me.CheckBoxExpLisAcquisto.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxExpLisAcquisto.Name = "CheckBoxExpLisAcquisto"
        Me.CheckBoxExpLisAcquisto.Size = New System.Drawing.Size(210, 20)
        Me.CheckBoxExpLisAcquisto.TabIndex = 2
        Me.CheckBoxExpLisAcquisto.Text = "Esporta prezzo listino acquisto"
        Me.CheckBoxExpLisAcquisto.UseVisualStyleBackColor = True
        '
        'CheckBoxExpLisVendita
        '
        Me.CheckBoxExpLisVendita.AutoSize = True
        Me.CheckBoxExpLisVendita.Checked = True
        Me.CheckBoxExpLisVendita.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBoxExpLisVendita.Location = New System.Drawing.Point(14, 51)
        Me.CheckBoxExpLisVendita.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxExpLisVendita.Name = "CheckBoxExpLisVendita"
        Me.CheckBoxExpLisVendita.Size = New System.Drawing.Size(203, 20)
        Me.CheckBoxExpLisVendita.TabIndex = 1
        Me.CheckBoxExpLisVendita.Text = "Esporta prezzo listino vendita"
        Me.CheckBoxExpLisVendita.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.CheckBoxFormatoImgFisso)
        Me.GroupBox4.Controls.Add(Me.CheckBoxImgVariante)
        Me.GroupBox4.Controls.Add(Me.CheckBoxImgArticolo)
        Me.GroupBox4.Location = New System.Drawing.Point(333, 317)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox4.Size = New System.Drawing.Size(247, 122)
        Me.GroupBox4.TabIndex = 18
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Immagini"
        '
        'CheckBoxFormatoImgFisso
        '
        Me.CheckBoxFormatoImgFisso.AutoSize = True
        Me.CheckBoxFormatoImgFisso.Location = New System.Drawing.Point(21, 79)
        Me.CheckBoxFormatoImgFisso.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxFormatoImgFisso.Name = "CheckBoxFormatoImgFisso"
        Me.CheckBoxFormatoImgFisso.Size = New System.Drawing.Size(208, 20)
        Me.CheckBoxFormatoImgFisso.TabIndex = 10
        Me.CheckBoxFormatoImgFisso.Text = "Dim. immagine fissa (100x100)"
        Me.CheckBoxFormatoImgFisso.UseVisualStyleBackColor = True
        '
        'CheckBoxImgVariante
        '
        Me.CheckBoxImgVariante.AutoSize = True
        Me.CheckBoxImgVariante.Location = New System.Drawing.Point(21, 51)
        Me.CheckBoxImgVariante.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBoxImgVariante.Name = "CheckBoxImgVariante"
        Me.CheckBoxImgVariante.Size = New System.Drawing.Size(189, 20)
        Me.CheckBoxImgVariante.TabIndex = 9
        Me.CheckBoxImgVariante.Text = "Esporta Immagine variante"
        Me.CheckBoxImgVariante.UseVisualStyleBackColor = True
        '
        'NumericUpDownMaggioriDi
        '
        Me.NumericUpDownMaggioriDi.Location = New System.Drawing.Point(0, 0)
        Me.NumericUpDownMaggioriDi.Name = "NumericUpDownMaggioriDi"
        Me.NumericUpDownMaggioriDi.Size = New System.Drawing.Size(120, 22)
        Me.NumericUpDownMaggioriDi.TabIndex = 0
        '
        'chkBuchiTaglia
        '
        Me.chkBuchiTaglia.AutoSize = True
        Me.chkBuchiTaglia.Checked = True
        Me.chkBuchiTaglia.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkBuchiTaglia.Location = New System.Drawing.Point(12, 141)
        Me.chkBuchiTaglia.Margin = New System.Windows.Forms.Padding(4)
        Me.chkBuchiTaglia.Name = "chkBuchiTaglia"
        Me.chkBuchiTaglia.Size = New System.Drawing.Size(205, 20)
        Me.chkBuchiTaglia.TabIndex = 21
        Me.chkBuchiTaglia.Text = "Esporta righe con buchi taglia"
        Me.chkBuchiTaglia.UseVisualStyleBackColor = True
        '
        'FormExport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1076, 458)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TextBoxEscludiTaglie)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBoxEscludiVarianti)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBoxDispTeorica)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBoxGiacenza)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.ButtonOK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "FormExport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Opzioni di esportazione  HTML"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.NumericUpDownMaxQta, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDownInterrPag, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.NumericUpDownMaggioriDi, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CheckBoxGiac As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxCalcolato As System.Windows.Forms.CheckBox
    Friend WithEvents ButtonOK As System.Windows.Forms.Button
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBoxGiacenza As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxDispTeorica As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CheckBoxImgArticolo As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxTotaliZero As System.Windows.Forms.CheckBox
    Friend WithEvents TextBoxEscludiVarianti As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBoxEscludiTaglie As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CheckBoxTaglie As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBoxArtFamiglia As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxArtComposizione As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxArtStagione As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxArtCodice As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBoxExpLisAcquisto As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxExpLisVendita As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxMarca As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxMadeIn As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxCodNomenclatura As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxTotaliMinZero As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBoxImgVariante As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxFormatoImgFisso As CheckBox
    Friend WithEvents CheckBoxInterPag As CheckBox
    Friend WithEvents NumericUpDownInterrPag As NumericUpDown
    Friend WithEvents NumericUpDownMaxQta As NumericUpDown
    Friend WithEvents NumericUpDownMaggioriDi As NumericUpDown
    Friend WithEvents CheckBoxMaxQta As CheckBox
    Friend WithEvents CheckBoxLingua As CheckBox
    Friend WithEvents chkBuchiTaglia As CheckBox
End Class
