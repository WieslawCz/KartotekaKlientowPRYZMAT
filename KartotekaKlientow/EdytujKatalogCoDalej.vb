Public Class EdytujKatalogCoDalej
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblNazwa2 As System.Windows.Forms.Label
    Friend WithEvents lblNazwa1 As System.Windows.Forms.Label
    Friend WithEvents lblNazwaHandlowa As System.Windows.Forms.Label
    Friend WithEvents lblIdent As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdWycofajSie As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents txtUwagi As System.Windows.Forms.TextBox
    Friend WithEvents txtRodzajDzialania As System.Windows.Forms.TextBox
    Friend WithEvents cmdWybierzZeSlownika As System.Windows.Forms.Button
    Friend WithEvents cmdKlient As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EdytujKatalogCoDalej))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdKlient = New System.Windows.Forms.Button()
        Me.cmdWybierzZeSlownika = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblNazwa2 = New System.Windows.Forms.Label()
        Me.lblNazwa1 = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa = New System.Windows.Forms.Label()
        Me.lblIdent = New System.Windows.Forms.Label()
        Me.txtUwagi = New System.Windows.Forms.TextBox()
        Me.txtRodzajDzialania = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdWycofajSie = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cmdKlient)
        Me.Panel1.Controls.Add(Me.cmdWybierzZeSlownika)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.lblNazwa2)
        Me.Panel1.Controls.Add(Me.lblNazwa1)
        Me.Panel1.Controls.Add(Me.lblNazwaHandlowa)
        Me.Panel1.Controls.Add(Me.lblIdent)
        Me.Panel1.Controls.Add(Me.txtUwagi)
        Me.Panel1.Controls.Add(Me.txtRodzajDzialania)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(536, 288)
        Me.Panel1.TabIndex = 27
        '
        'cmdKlient
        '
        Me.cmdKlient.Image = CType(resources.GetObject("cmdKlient.Image"), System.Drawing.Image)
        Me.cmdKlient.Location = New System.Drawing.Point(200, 9)
        Me.cmdKlient.Name = "cmdKlient"
        Me.cmdKlient.Size = New System.Drawing.Size(32, 23)
        Me.cmdKlient.TabIndex = 89
        '
        'cmdWybierzZeSlownika
        '
        Me.cmdWybierzZeSlownika.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzZeSlownika.Image = CType(resources.GetObject("cmdWybierzZeSlownika.Image"), System.Drawing.Image)
        Me.cmdWybierzZeSlownika.Location = New System.Drawing.Point(488, 95)
        Me.cmdWybierzZeSlownika.Name = "cmdWybierzZeSlownika"
        Me.cmdWybierzZeSlownika.Size = New System.Drawing.Size(31, 20)
        Me.cmdWybierzZeSlownika.TabIndex = 88
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(0, 88)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(535, 1)
        Me.Panel2.TabIndex = 23
        '
        'Label8
        '
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(8, 48)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 16)
        Me.Label8.TabIndex = 22
        Me.Label8.Text = "Adres klienta . . . . . . . . ."
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(8, 32)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 16)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "Nazwa klienta . . . . . . . . ."
        '
        'lblNazwa2
        '
        Me.lblNazwa2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa2.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa2.Location = New System.Drawing.Point(120, 64)
        Me.lblNazwa2.Name = "lblNazwa2"
        Me.lblNazwa2.Size = New System.Drawing.Size(399, 16)
        Me.lblNazwa2.TabIndex = 20
        '
        'lblNazwa1
        '
        Me.lblNazwa1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa1.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa1.Location = New System.Drawing.Point(120, 48)
        Me.lblNazwa1.Name = "lblNazwa1"
        Me.lblNazwa1.Size = New System.Drawing.Size(399, 16)
        Me.lblNazwa1.TabIndex = 19
        '
        'lblNazwaHandlowa
        '
        Me.lblNazwaHandlowa.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwaHandlowa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwaHandlowa.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwaHandlowa.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwaHandlowa.Location = New System.Drawing.Point(120, 32)
        Me.lblNazwaHandlowa.Name = "lblNazwaHandlowa"
        Me.lblNazwaHandlowa.Size = New System.Drawing.Size(399, 16)
        Me.lblNazwaHandlowa.TabIndex = 18
        '
        'lblIdent
        '
        Me.lblIdent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdent.ForeColor = System.Drawing.Color.Purple
        Me.lblIdent.Location = New System.Drawing.Point(120, 8)
        Me.lblIdent.Name = "lblIdent"
        Me.lblIdent.Size = New System.Drawing.Size(80, 24)
        Me.lblIdent.TabIndex = 17
        Me.lblIdent.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtUwagi
        '
        Me.txtUwagi.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUwagi.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtUwagi.ForeColor = System.Drawing.Color.Purple
        Me.txtUwagi.Location = New System.Drawing.Point(120, 121)
        Me.txtUwagi.Multiline = True
        Me.txtUwagi.Name = "txtUwagi"
        Me.txtUwagi.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtUwagi.Size = New System.Drawing.Size(399, 159)
        Me.txtUwagi.TabIndex = 15
        '
        'txtRodzajDzialania
        '
        Me.txtRodzajDzialania.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRodzajDzialania.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtRodzajDzialania.ForeColor = System.Drawing.Color.Purple
        Me.txtRodzajDzialania.Location = New System.Drawing.Point(120, 95)
        Me.txtRodzajDzialania.Name = "txtRodzajDzialania"
        Me.txtRodzajDzialania.ReadOnly = True
        Me.txtRodzajDzialania.Size = New System.Drawing.Size(399, 20)
        Me.txtRodzajDzialania.TabIndex = 14
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 124)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Opis . . . . . . . . . . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 98)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 16)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Rodzaj dzia³ania . . . . . . . . . . . . . "
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Identyfikator klienta . . . . . . . . ."
        '
        'cmdWycofajSie
        '
        Me.cmdWycofajSie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWycofajSie.Image = CType(resources.GetObject("cmdWycofajSie.Image"), System.Drawing.Image)
        Me.cmdWycofajSie.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWycofajSie.Location = New System.Drawing.Point(355, 301)
        Me.cmdWycofajSie.Name = "cmdWycofajSie"
        Me.cmdWycofajSie.Size = New System.Drawing.Size(104, 25)
        Me.cmdWycofajSie.TabIndex = 24
        Me.cmdWycofajSie.Text = "Wycofaj siê"
        Me.cmdWycofajSie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(465, 301)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 25)
        Me.cmdPowrot.TabIndex = 23
        Me.cmdPowrot.Text = "Zapisz"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 301)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 26
        Me.PictureBox2.TabStop = False
        '
        'EdytujKatalogCoDalej
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(554, 329)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdWycofajSie)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "EdytujKatalogCoDalej"
        Me.ShowInTaskbar = False
        Me.Text = "Co dalej"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub EdytujKatalogCoDalej_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        PobierzDane()
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
            cmdKlient.Enabled = False
            cmdWybierzZeSlownika.Enabled = True
        Else
            cmdKlient.Enabled = True
            cmdWybierzZeSlownika.Enabled = True
        End If
    End Sub

    Private Sub EdytujKatalogCoDalej_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
        Else
        End If
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        ZapiszDane()
        oAktualizuj = True
        Me.Close()
    End Sub


    Private Sub cmdWycofajSie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWycofajSie.Click
        oAktualizuj = False
        Me.Close()
    End Sub


    Private Sub EdytujKatalogCoDalej_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub


    Private Sub cmdKlient_Click(sender As Object, e As EventArgs) Handles cmdKlient.Click
        oCoMamRobic = "WYBIERAJ"
        oKlient = Trim(lblIdent.Text)
        Dim form As New KatalogKlientow
        form.ShowDialog()
        If Len(Trim(oKlient)) > 0 Then
            lblIdent.Text = oKlient
            IdentKlienta(oKlient)
            lblNazwaHandlowa.Text = oNazwa1Kli
            lblNazwa1.Text = oUlicaKli
            lblNazwa2.Text = oKodPoczKli & "  " & oMiejscKli
        End If
    End Sub

    Private Sub cmdWybierzZeSlownika_Click(sender As Object, e As EventArgs) Handles cmdWybierzZeSlownika.Click
        oCoMamRobic = "WYBIERAJ"
        oCoDalej = Trim(txtRodzajDzialania.Text)
        Dim form As New KatalogSlownikCoDalej
        form.ShowDialog()
        If Len(Trim(oCoDalej)) > 0 Then
            txtRodzajDzialania.Text = oCoDalej
        End If
    End Sub



    '====================================================
    'UWAGA - trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    Private Sub EdytujKatalogCoDalej_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me.ClosedByControlBox Then
            'e.Cancel = True     'nie pozwalaj zamkn¹c forme
            If MessageBox.Show("Zdecydowa³eœ opuœciæ formê bez zapisania wprowadzonych zmian." & vbCrLf & _
                "Czy mam zapisaæ zmiany w Bazie Danych ?", _
                "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.YesNo, _
                MessageBoxIcon.Question, _
                MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then

                ZapiszDane()
                oAktualizuj = True
            Else
                oAktualizuj = False
            End If
            'Me._closeClick = False
        Else
            'MsgBox("ZAPISZ lub WYCOFAJ SIÊ")
        End If
    End Sub
    '====================================================




    '***************************************************
    '* procedury pobierania/zapisywania danych
    '***************************************************

    Private Sub PobierzDane()
        lblIdent.Text = cdIDENTKLIENTA

        IdentKlienta(cdIDENTKLIENTA)
        lblNazwaHandlowa.Text = oNazwa1Kli
        lblNazwa1.Text = oUlicaKli
        lblNazwa2.Text = oKodPoczKli & "  " & oMiejscKli

        txtRodzajDzialania.Text = cdIDENT
        txtUwagi.Text = cdOPIS
    End Sub

    Private Sub ZapiszDane()
        cdIDENTKLIENTA = lblIdent.Text
        cdIDENT = txtRodzajDzialania.Text
        cdOPIS = txtUwagi.Text
    End Sub

    '*******************************************
    '** procedury obslugi edycji
    '*******************************************

    Private Sub txtTemat_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRodzajDzialania.TextChanged
        TestLen(txtRodzajDzialania, "RODZAJ DZIA£ANIA", 20)
    End Sub

    Private Sub txtTemat_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRodzajDzialania.GotFocus
        StartEdycjiTxt(txtRodzajDzialania)
    End Sub
    Private Sub txtTemat_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRodzajDzialania.LostFocus
        KoniecEdycjiTxt(txtRodzajDzialania)
    End Sub

    Private Sub txtUwagi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.TextChanged
        'TestLen(txtUwagi, "NOTATKA", 50)
    End Sub
    Private Sub txtUwagi_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.GotFocus
        StartEdycjiTxt(txtUwagi)
    End Sub
    Private Sub txtUwagi_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.LostFocus
        KoniecEdycjiTxt(txtUwagi)
    End Sub

End Class
