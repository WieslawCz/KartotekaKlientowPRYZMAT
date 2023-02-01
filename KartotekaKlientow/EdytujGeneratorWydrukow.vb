Public Class EdytujGeneratorWydrukow

    Private Sub EdytujGeneratorWydrukow_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        InicjujListeWyrownaniaTekstu(cbxWyrownanie)
        PobierzDane()
    End Sub

    Private Sub EdytujGeneratorWydrukow_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        ZapiszDane()
        oAktualizuj = True
        Me.Close()
    End Sub

    Private Sub PobierzDane()
        txtRozmiar.Text = GenRozmiar
        txtNaglowek.Text = GenNaglowek
        lblNazwa.Text = GenNazwa
        lblTyp.Text = GenTyp
        txtMaxIlWierszy.Text = GenMaxIlWierszy
        cbxWyrownanie.Text = GenWyrownanie
    End Sub

    Private Sub ZapiszDane()
        GenRozmiar = txtRozmiar.Text
        GenNaglowek = txtNaglowek.Text
        GenWyrownanie = cbxWyrownanie.Text
        GenMaxIlWierszy = txtMaxIlWierszy.Text
    End Sub


    Private Sub txtNaglowek_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNaglowek.TextChanged
        'TestLen(txtNaglowek, "NAG£OWEK", 20)
    End Sub
    Private Sub txtNaglowek_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNaglowek.GotFocus
        StartEdycjiTxt(txtNaglowek)
    End Sub
    Private Sub txtNaglowek_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNaglowek.LostFocus
        KoniecEdycjiTxt(txtNaglowek)
    End Sub

    Private Sub txtRozmiar_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRozmiar.TextChanged
        If TestNum(txtRozmiar, "ROZMIAR") Then
        End If
    End Sub
    Private Sub txtRozmiar_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRozmiar.GotFocus
        StartEdycjiTxt(txtRozmiar)
    End Sub
    Private Sub txtRozmiar_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRozmiar.LostFocus
        KoniecEdycjiTxt(txtRozmiar)
    End Sub

    Private Sub txtMaxIlWierszy_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMaxIlWierszy.TextChanged
        If TestNum(txtMaxIlWierszy, "MAX ILOSC WIERSZY") Then
        End If
    End Sub
    Private Sub txtMaxIlWierszy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMaxIlWierszy.GotFocus
        StartEdycjiTxt(txtMaxIlWierszy)
    End Sub
    Private Sub txtMaxIlWierszy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMaxIlWierszy.LostFocus
        KoniecEdycjiTxt(txtMaxIlWierszy)
    End Sub

    Private Sub cbxWyrownanie_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxWyrownanie.TextChanged
        'TestCbxLen(cbxWyrownanie, "WYROWNANIE", 20)
    End Sub
    Private Sub cbxWyrownanie_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxWyrownanie.GotFocus
        StartEdycjiCbx(cbxWyrownanie)
    End Sub
    Private Sub cbxWyrownanie_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxWyrownanie.LostFocus
        KoniecEdycjiCbx(cbxWyrownanie)
    End Sub
End Class