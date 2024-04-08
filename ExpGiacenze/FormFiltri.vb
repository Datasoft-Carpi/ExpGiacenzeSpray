Public Class FormFiltri
    Dim countArt As Integer = FormMain.ListFiltriArticoli.Count - 2
    Dim countVar As Integer = FormMain.ListFiltriVarianti.Count - 2




    ' popola 
    Private Sub FormFiltri_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub


    Private Sub popolaListView()
        Dim i As Integer
        ' popola la lista degli articoli
        CheckedListBoxArticoli.Items.Clear()
        For i = 0 To countArt
            CheckedListBoxArticoli.Items.Add(FormMain.ListFiltriArticoli(i).Codice & " - " & FormMain.ListFiltriArticoli(i).descrizione)
            CheckedListBoxArticoli.SetItemChecked(i, FormMain.ListFiltriArticoli(i).visibile)
        Next

        ' popola la lista degli articoli
        CheckedListBoxVarianti.Items.Clear()
        For i = 0 To countVar
            CheckedListBoxVarianti.Items.Add(FormMain.ListFiltriVarianti(i).Codice)
            CheckedListBoxVarianti.SetItemChecked(i, FormMain.ListFiltriVarianti(i).visibile)
        Next

    End Sub


    Private Sub ButtonOK_Click(sender As Object, e As EventArgs) Handles ButtonOK.Click

        ' aggiorna le strutture
        Call AggiornastrutturaArticoli()
        Call AggiornastrutturaVarianti()

        Me.DialogResult = vbOK
    End Sub

    Private Sub FormFiltri_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Call popolaListView()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call selezionaAll(CheckedListBoxArticoli, False)

    End Sub

    Private Sub selezionaAll(ckList As CheckedListBox, value As Boolean)
        For i = 0 To ckList.Items.Count - 1
            ckList.SetItemChecked(i, value)
        Next

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call selezionaAll(CheckedListBoxArticoli, True)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call selezionaAll(CheckedListBoxVarianti, False)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call selezionaAll(CheckedListBoxVarianti, True)

    End Sub



    Private Sub AggiornastrutturaArticoli()
        Dim strCod As String
        Dim i As Integer
        Dim j As Integer

        For i = 0 To CheckedListBoxArticoli.Items.Count - 1
            strCod = CheckedListBoxArticoli.Items(i)
            For j = 0 To countArt
                If strCod = FormMain.ListFiltriArticoli(j).Codice & " - " & FormMain.ListFiltriArticoli(j).descrizione Then
                    FormMain.ListFiltriArticoli(j).visibile = CheckedListBoxArticoli.GetItemCheckState(i)
                End If

            Next
        Next
    End Sub


    Private Sub AggiornastrutturaVarianti()
        Dim strCod As String
        Dim i As Integer
        Dim j As Integer

        For i = 0 To CheckedListBoxVarianti.Items.Count - 1
            strCod = CheckedListBoxVarianti.Items(i)
            For j = 0 To countVar
                If strCod = FormMain.ListFiltriVarianti(j).Codice Then
                    FormMain.ListFiltriVarianti(j).visibile = CheckedListBoxVarianti.GetItemCheckState(i)
                End If

            Next
        Next
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Me.DialogResult = vbCancel
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim strCod As String
        Dim i As Integer
        Dim j As Integer

        If TextBoxSearchArt.Text = "" Then
            MsgBox("Nessun Articolo digitato ")
            Exit Sub
        End If


        For i = 0 To CheckedListBoxArticoli.Items.Count - 1
            strCod = CheckedListBoxArticoli.Items(i)

            If strCod.Contains(TextBoxSearchArt.Text) = True Then
                CheckedListBoxArticoli.SetItemChecked(i, True)
                'Else
             '   CheckedListBoxArticoli.SetItemChecked(i, False)
            End If


        Next

    End Sub

End Class