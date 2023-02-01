Public Class RaportUsunAkcjaMarketingowa

    Private Sub RaportWykonania_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        lblFunkcja.Text = ""
        lblOperacja.Text = ""
        Progres.Maximum = 0
        Progres.Value = 0

        lblFunkcja.Text = "Zapisuje informacje o akcji marketingowej " & _Ident

        Me.Opacity = 0.8

        Timer1.Interval = 200
        Timer1.Enabled = True
    End Sub

    Private Sub RaportWykonania_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False

        Dim foundRow As DataRow
        Dim dvRow As DataRowView
        ' Create an array for the key values to find.
        Dim findTheseVals(1) As Object

        'DOPISZ informacje o klientach
        Dim i As Long
        lblOperacja.Text = "Zapamiêtuje informacje o klientach"
        Progres.Maximum = _dvKlienci.Count - 1
        For i = 0 To _dvKlienci.Count - 1
            Progres.Value = i
            System.Windows.Forms.Application.DoEvents()

            dvRow = _dvKlienci.Item(i)
            oIdentKli = dvRow.Item("NRKARTY")


            findTheseVals(0) = _Ident
            findTheseVals(1) = oIdentKli
            foundRow = _dsAkcjeSpec.Tables(0).Rows.Find(findTheseVals)
            ' sprawdz czy jest...
            If Not (foundRow Is Nothing) Then
                foundRow.Delete()
            End If
        Next
        _AktualizujAkcjeSpec()
        Me.Close()
    End Sub


End Class