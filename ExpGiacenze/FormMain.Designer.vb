<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormMain
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormMain))
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.DataGridViewArticoli = New System.Windows.Forms.DataGridView()
        Me.Immagine = New System.Windows.Forms.DataGridViewImageColumn()
        Me.Immagine.DefaultCellStyle.NullValue = Nothing
        Me.codice = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Articolo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IDArray = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LisAcquisto = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LisVendita = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CodArticolo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ButtonExportHTML = New System.Windows.Forms.Button()
        Me.LabelDB = New System.Windows.Forms.Label()
        Me.DataGridViewGiacenze = New System.Windows.Forms.DataGridView()
        Me.LabelArticoli = New System.Windows.Forms.Label()
        Me.ProgressBarArticoli = New System.Windows.Forms.ProgressBar()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.LabelLisVen = New System.Windows.Forms.Label()
        Me.LabelLisAcq = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.BtnFiltri = New System.Windows.Forms.Button()
        Me.ButtonExportEXCEL = New System.Windows.Forms.Button()
        CType(Me.DataGridViewArticoli, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridViewGiacenze, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridViewArticoli
        '
        Me.DataGridViewArticoli.AllowUserToAddRows = False
        Me.DataGridViewArticoli.AllowUserToDeleteRows = False
        Me.DataGridViewArticoli.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewArticoli.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridViewArticoli.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridViewArticoli.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Immagine, Me.codice, Me.Articolo, Me.IDArray, Me.LisAcquisto, Me.LisVendita, Me.CodArticolo})
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridViewArticoli.DefaultCellStyle = DataGridViewCellStyle5
        Me.DataGridViewArticoli.Location = New System.Drawing.Point(12, 52)
        Me.DataGridViewArticoli.Name = "DataGridViewArticoli"
        Me.DataGridViewArticoli.ReadOnly = True
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewArticoli.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.DataGridViewArticoli.RowTemplate.Height = 140
        Me.DataGridViewArticoli.Size = New System.Drawing.Size(442, 442)
        Me.DataGridViewArticoli.TabIndex = 0
        '
        'Immagine
        '
        Me.Immagine.HeaderText = "Immagine"
        Me.Immagine.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Stretch
        Me.Immagine.Name = "Immagine"
        Me.Immagine.ReadOnly = True
        Me.Immagine.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Immagine.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'codice
        '
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.codice.DefaultCellStyle = DataGridViewCellStyle2
        Me.codice.HeaderText = "Codice"
        Me.codice.Name = "codice"
        Me.codice.ReadOnly = True
        Me.codice.Width = 175
        '
        'Articolo
        '
        Me.Articolo.HeaderText = "Articolo"
        Me.Articolo.Name = "Articolo"
        Me.Articolo.ReadOnly = True
        Me.Articolo.Visible = False
        '
        'IDArray
        '
        Me.IDArray.HeaderText = "IDArray"
        Me.IDArray.Name = "IDArray"
        Me.IDArray.ReadOnly = True
        Me.IDArray.Visible = False
        '
        'LisAcquisto
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.Format = "N2"
        DataGridViewCellStyle3.NullValue = Nothing
        Me.LisAcquisto.DefaultCellStyle = DataGridViewCellStyle3
        Me.LisAcquisto.HeaderText = "Acquisto"
        Me.LisAcquisto.Name = "LisAcquisto"
        Me.LisAcquisto.ReadOnly = True
        Me.LisAcquisto.Visible = False
        Me.LisAcquisto.Width = 50
        '
        'LisVendita
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.Format = "N2"
        DataGridViewCellStyle4.NullValue = Nothing
        Me.LisVendita.DefaultCellStyle = DataGridViewCellStyle4
        Me.LisVendita.HeaderText = "Vendita"
        Me.LisVendita.Name = "LisVendita"
        Me.LisVendita.ReadOnly = True
        Me.LisVendita.Visible = False
        Me.LisVendita.Width = 50
        '
        'CodArticolo
        '
        Me.CodArticolo.HeaderText = "CodArticolo"
        Me.CodArticolo.Name = "CodArticolo"
        Me.CodArticolo.ReadOnly = True
        Me.CodArticolo.Visible = False
        '
        'ButtonExportHTML
        '
        Me.ButtonExportHTML.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonExportHTML.Image = CType(resources.GetObject("ButtonExportHTML.Image"), System.Drawing.Image)
        Me.ButtonExportHTML.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonExportHTML.Location = New System.Drawing.Point(947, 11)
        Me.ButtonExportHTML.Name = "ButtonExportHTML"
        Me.ButtonExportHTML.Size = New System.Drawing.Size(112, 34)
        Me.ButtonExportHTML.TabIndex = 1
        Me.ButtonExportHTML.Text = "Esporta HTML"
        Me.ButtonExportHTML.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ButtonExportHTML.UseVisualStyleBackColor = True
        '
        'LabelDB
        '
        Me.LabelDB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelDB.AutoSize = True
        Me.LabelDB.Location = New System.Drawing.Point(13, 517)
        Me.LabelDB.Name = "LabelDB"
        Me.LabelDB.Size = New System.Drawing.Size(39, 13)
        Me.LabelDB.TabIndex = 2
        Me.LabelDB.Text = "Label1"
        '
        'DataGridViewGiacenze
        '
        Me.DataGridViewGiacenze.AllowUserToAddRows = False
        Me.DataGridViewGiacenze.AllowUserToDeleteRows = False
        Me.DataGridViewGiacenze.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewGiacenze.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.DataGridViewGiacenze.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridViewGiacenze.DefaultCellStyle = DataGridViewCellStyle8
        Me.DataGridViewGiacenze.Location = New System.Drawing.Point(460, 52)
        Me.DataGridViewGiacenze.Name = "DataGridViewGiacenze"
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewGiacenze.RowHeadersDefaultCellStyle = DataGridViewCellStyle9
        Me.DataGridViewGiacenze.Size = New System.Drawing.Size(717, 382)
        Me.DataGridViewGiacenze.TabIndex = 3
        Me.DataGridViewGiacenze.Visible = False
        '
        'LabelArticoli
        '
        Me.LabelArticoli.AutoSize = True
        Me.LabelArticoli.Location = New System.Drawing.Point(12, 32)
        Me.LabelArticoli.Name = "LabelArticoli"
        Me.LabelArticoli.Size = New System.Drawing.Size(38, 13)
        Me.LabelArticoli.TabIndex = 4
        Me.LabelArticoli.Text = "Articoli"
        '
        'ProgressBarArticoli
        '
        Me.ProgressBarArticoli.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ProgressBarArticoli.Location = New System.Drawing.Point(12, 499)
        Me.ProgressBarArticoli.Name = "ProgressBarArticoli"
        Me.ProgressBarArticoli.Size = New System.Drawing.Size(442, 13)
        Me.ProgressBarArticoli.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(487, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(122, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Giacenze / Disp. teorica"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.LabelLisVen)
        Me.GroupBox1.Controls.Add(Me.LabelLisAcq)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(460, 440)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(717, 72)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "[ Listini ]"
        '
        'LabelLisVen
        '
        Me.LabelLisVen.AutoSize = True
        Me.LabelLisVen.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelLisVen.Location = New System.Drawing.Point(73, 49)
        Me.LabelLisVen.Name = "LabelLisVen"
        Me.LabelLisVen.Size = New System.Drawing.Size(45, 13)
        Me.LabelLisVen.TabIndex = 3
        Me.LabelLisVen.Text = "LisVen"
        '
        'LabelLisAcq
        '
        Me.LabelLisAcq.AutoSize = True
        Me.LabelLisAcq.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelLisAcq.Location = New System.Drawing.Point(73, 23)
        Me.LabelLisAcq.Name = "LabelLisAcq"
        Me.LabelLisAcq.Size = New System.Drawing.Size(45, 13)
        Me.LabelLisAcq.TabIndex = 2
        Me.LabelLisAcq.Text = "LisAcq"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(20, 49)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Vendita:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Acquisto:"
        '
        'BtnFiltri
        '
        Me.BtnFiltri.Image = CType(resources.GetObject("BtnFiltri.Image"), System.Drawing.Image)
        Me.BtnFiltri.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnFiltri.Location = New System.Drawing.Point(837, 12)
        Me.BtnFiltri.Name = "BtnFiltri"
        Me.BtnFiltri.Size = New System.Drawing.Size(104, 32)
        Me.BtnFiltri.TabIndex = 10
        Me.BtnFiltri.Text = "Filtra"
        Me.BtnFiltri.UseVisualStyleBackColor = True
        '
        'ButtonExportEXCEL
        '
        Me.ButtonExportEXCEL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonExportEXCEL.Image = CType(resources.GetObject("ButtonExportEXCEL.Image"), System.Drawing.Image)
        Me.ButtonExportEXCEL.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonExportEXCEL.Location = New System.Drawing.Point(1065, 10)
        Me.ButtonExportEXCEL.Name = "ButtonExportEXCEL"
        Me.ButtonExportEXCEL.Size = New System.Drawing.Size(112, 34)
        Me.ButtonExportEXCEL.TabIndex = 12
        Me.ButtonExportEXCEL.Text = "Esporta Excel"
        Me.ButtonExportEXCEL.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ButtonExportEXCEL.UseVisualStyleBackColor = True
        '
        'FormMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1189, 542)
        Me.Controls.Add(Me.ButtonExportEXCEL)
        Me.Controls.Add(Me.BtnFiltri)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ProgressBarArticoli)
        Me.Controls.Add(Me.LabelArticoli)
        Me.Controls.Add(Me.DataGridViewGiacenze)
        Me.Controls.Add(Me.LabelDB)
        Me.Controls.Add(Me.ButtonExportHTML)
        Me.Controls.Add(Me.DataGridViewArticoli)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FormMain"
        Me.Text = "Esporta giacenze"
        CType(Me.DataGridViewArticoli, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridViewGiacenze, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridViewArticoli As System.Windows.Forms.DataGridView
    Friend WithEvents ButtonExportHTML As System.Windows.Forms.Button
    Friend WithEvents LabelDB As System.Windows.Forms.Label
    Friend WithEvents DataGridViewGiacenze As System.Windows.Forms.DataGridView
    Friend WithEvents LabelArticoli As System.Windows.Forms.Label
    Friend WithEvents ProgressBarArticoli As System.Windows.Forms.ProgressBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LabelLisVen As System.Windows.Forms.Label
    Friend WithEvents LabelLisAcq As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BtnFiltri As Button
    Friend WithEvents Immagine As DataGridViewImageColumn
    Friend WithEvents codice As DataGridViewTextBoxColumn
    Friend WithEvents Articolo As DataGridViewTextBoxColumn
    Friend WithEvents IDArray As DataGridViewTextBoxColumn
    Friend WithEvents LisAcquisto As DataGridViewTextBoxColumn
    Friend WithEvents LisVendita As DataGridViewTextBoxColumn
    Friend WithEvents CodArticolo As DataGridViewTextBoxColumn
    Friend WithEvents ButtonExportEXCEL As Button
End Class
