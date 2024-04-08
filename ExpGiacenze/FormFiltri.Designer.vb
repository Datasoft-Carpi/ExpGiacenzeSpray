<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormFiltri
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    'Non modificarla mediante l'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TabControlFiltri = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.TextBoxSearchArt = New System.Windows.Forms.TextBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.CheckedListBoxArticoli = New System.Windows.Forms.CheckedListBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.CheckedListBoxVarianti = New System.Windows.Forms.CheckedListBox()
        Me.ButtonOK = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TabControlFiltri.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControlFiltri
        '
        Me.TabControlFiltri.Controls.Add(Me.TabPage1)
        Me.TabControlFiltri.Controls.Add(Me.TabPage2)
        Me.TabControlFiltri.Location = New System.Drawing.Point(7, 6)
        Me.TabControlFiltri.Name = "TabControlFiltri"
        Me.TabControlFiltri.SelectedIndex = 0
        Me.TabControlFiltri.Size = New System.Drawing.Size(513, 420)
        Me.TabControlFiltri.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Button6)
        Me.TabPage1.Controls.Add(Me.TextBoxSearchArt)
        Me.TabPage1.Controls.Add(Me.Button3)
        Me.TabPage1.Controls.Add(Me.Button2)
        Me.TabPage1.Controls.Add(Me.CheckedListBoxArticoli)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(505, 394)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Articoli"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(144, 10)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(75, 23)
        Me.Button6.TabIndex = 6
        Me.Button6.Text = "Spunta"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'TextBoxSearchArt
        '
        Me.TextBoxSearchArt.Location = New System.Drawing.Point(6, 12)
        Me.TextBoxSearchArt.Name = "TextBoxSearchArt"
        Me.TextBoxSearchArt.Size = New System.Drawing.Size(131, 20)
        Me.TextBoxSearchArt.TabIndex = 5
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(343, 10)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "nessuno"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(424, 10)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "tutti"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'CheckedListBoxArticoli
        '
        Me.CheckedListBoxArticoli.CheckOnClick = True
        Me.CheckedListBoxArticoli.FormattingEnabled = True
        Me.CheckedListBoxArticoli.Location = New System.Drawing.Point(3, 39)
        Me.CheckedListBoxArticoli.Name = "CheckedListBoxArticoli"
        Me.CheckedListBoxArticoli.Size = New System.Drawing.Size(499, 349)
        Me.CheckedListBoxArticoli.TabIndex = 1
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Button4)
        Me.TabPage2.Controls.Add(Me.Button5)
        Me.TabPage2.Controls.Add(Me.CheckedListBoxVarianti)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(505, 394)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Varianti"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(346, 9)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 23)
        Me.Button4.TabIndex = 6
        Me.Button4.Text = "nessuno"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(427, 9)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(75, 23)
        Me.Button5.TabIndex = 5
        Me.Button5.Text = "tutti"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'CheckedListBoxVarianti
        '
        Me.CheckedListBoxVarianti.CheckOnClick = True
        Me.CheckedListBoxVarianti.FormattingEnabled = True
        Me.CheckedListBoxVarianti.Location = New System.Drawing.Point(3, 38)
        Me.CheckedListBoxVarianti.Name = "CheckedListBoxVarianti"
        Me.CheckedListBoxVarianti.Size = New System.Drawing.Size(499, 349)
        Me.CheckedListBoxVarianti.TabIndex = 2
        '
        'ButtonOK
        '
        Me.ButtonOK.Location = New System.Drawing.Point(441, 439)
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.Size = New System.Drawing.Size(75, 23)
        Me.ButtonOK.TabIndex = 1
        Me.ButtonOK.Text = "APPLICA"
        Me.ButtonOK.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(354, 439)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "ANNULLA"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FormFiltri
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(527, 481)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ButtonOK)
        Me.Controls.Add(Me.TabControlFiltri)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "FormFiltri"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Filtri"
        Me.TabControlFiltri.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabControlFiltri As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents ButtonOK As Button
    Friend WithEvents CheckedListBoxArticoli As CheckedListBox
    Friend WithEvents CheckedListBoxVarianti As CheckedListBox
    Friend WithEvents Button3 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents TextBoxSearchArt As TextBox
End Class
