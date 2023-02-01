Public Class EdytujKatalogKontaktow
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
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdWycofajSie As System.Windows.Forms.Button
    Friend WithEvents cmdPrzywroc As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents txtPracownik As System.Windows.Forms.TextBox
    Friend WithEvents txtUwagi As System.Windows.Forms.TextBox
    Friend WithEvents txtRodzajKontaktu As System.Windows.Forms.TextBox
    Friend WithEvents cmdRodzajKontaktu As System.Windows.Forms.Button
    Friend WithEvents cmdKlient As System.Windows.Forms.Button
    Friend WithEvents dtpDataKontaktu As DateTimePicker
    Friend WithEvents cmdPracownik As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EdytujKatalogKontaktow))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.dtpDataKontaktu = New System.Windows.Forms.DateTimePicker()
        Me.cmdKlient = New System.Windows.Forms.Button()
        Me.cmdRodzajKontaktu = New System.Windows.Forms.Button()
        Me.cmdPracownik = New System.Windows.Forms.Button()
        Me.txtPracownik = New System.Windows.Forms.TextBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblNazwa2 = New System.Windows.Forms.Label()
        Me.lblNazwa1 = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa = New System.Windows.Forms.Label()
        Me.lblIdent = New System.Windows.Forms.Label()
        Me.txtUwagi = New System.Windows.Forms.TextBox()
        Me.txtRodzajKontaktu = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdWycofajSie = New System.Windows.Forms.Button()
        Me.cmdPrzywroc = New System.Windows.Forms.Button()
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
        Me.Panel1.Controls.Add(Me.dtpDataKontaktu)
        Me.Panel1.Controls.Add(Me.cmdKlient)
        Me.Panel1.Controls.Add(Me.cmdRodzajKontaktu)
        Me.Panel1.Controls.Add(Me.cmdPracownik)
        Me.Panel1.Controls.Add(Me.txtPracownik)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.lblNazwa2)
        Me.Panel1.Controls.Add(Me.lblNazwa1)
        Me.Panel1.Controls.Add(Me.lblNazwaHandlowa)
        Me.Panel1.Controls.Add(Me.lblIdent)
        Me.Panel1.Controls.Add(Me.txtUwagi)
        Me.Panel1.Controls.Add(Me.txtRodzajKontaktu)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(648, 369)
        Me.Panel1.TabIndex = 27
        '
        'dtpDataKontaktu
        '
        Me.dtpDataKontaktu.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtpDataKontaktu.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpDataKontaktu.CustomFormat = "yyyy-MM-dd"
        Me.dtpDataKontaktu.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDataKontaktu.Location = New System.Drawing.Point(120, 96)
        Me.dtpDataKontaktu.Name = "dtpDataKontaktu"
        Me.dtpDataKontaktu.Size = New System.Drawing.Size(112, 20)
        Me.dtpDataKontaktu.TabIndex = 186
        Me.dtpDataKontaktu.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'cmdKlient
        '
        Me.cmdKlient.Image = CType(resources.GetObject("cmdKlient.Image"), System.Drawing.Image)
        Me.cmdKlient.Location = New System.Drawing.Point(200, 9)
        Me.cmdKlient.Name = "cmdKlient"
        Me.cmdKlient.Size = New System.Drawing.Size(32, 23)
        Me.cmdKlient.TabIndex = 89
        '
        'cmdRodzajKontaktu
        '
        Me.cmdRodzajKontaktu.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdRodzajKontaktu.Image = CType(resources.GetObject("cmdRodzajKontaktu.Image"), System.Drawing.Image)
        Me.cmdRodzajKontaktu.Location = New System.Drawing.Point(600, 119)
        Me.cmdRodzajKontaktu.Name = "cmdRodzajKontaktu"
        Me.cmdRodzajKontaktu.Size = New System.Drawing.Size(32, 20)
        Me.cmdRodzajKontaktu.TabIndex = 88
        '
        'cmdPracownik
        '
        Me.cmdPracownik.Image = CType(resources.GetObject("cmdPracownik.Image"), System.Drawing.Image)
        Me.cmdPracownik.Location = New System.Drawing.Point(488, 96)
        Me.cmdPracownik.Name = "cmdPracownik"
        Me.cmdPracownik.Size = New System.Drawing.Size(32, 20)
        Me.cmdPracownik.TabIndex = 87
        '
        'txtPracownik
        '
        Me.txtPracownik.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtPracownik.ForeColor = System.Drawing.Color.Purple
        Me.txtPracownik.Location = New System.Drawing.Point(376, 96)
        Me.txtPracownik.Name = "txtPracownik"
        Me.txtPracownik.Size = New System.Drawing.Size(112, 20)
        Me.txtPracownik.TabIndex = 9
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(0, 88)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(648, 1)
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
        Me.lblNazwa2.Size = New System.Drawing.Size(512, 16)
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
        Me.lblNazwa1.Size = New System.Drawing.Size(512, 16)
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
        Me.lblNazwaHandlowa.Size = New System.Drawing.Size(512, 16)
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
        Me.txtUwagi.Location = New System.Drawing.Point(120, 144)
        Me.txtUwagi.Multiline = True
        Me.txtUwagi.Name = "txtUwagi"
        Me.txtUwagi.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtUwagi.Size = New System.Drawing.Size(512, 216)
        Me.txtUwagi.TabIndex = 15
        '
        'txtRodzajKontaktu
        '
        Me.txtRodzajKontaktu.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRodzajKontaktu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtRodzajKontaktu.ForeColor = System.Drawing.Color.Purple
        Me.txtRodzajKontaktu.Location = New System.Drawing.Point(120, 120)
        Me.txtRodzajKontaktu.Name = "txtRodzajKontaktu"
        Me.txtRodzajKontaktu.ReadOnly = True
        Me.txtRodzajKontaktu.Size = New System.Drawing.Size(480, 20)
        Me.txtRodzajKontaktu.TabIndex = 14
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(296, 96)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 16)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Pracownik . . . . . . . . . . . . ."
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 144)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Opis . . . . . . . . . . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 120)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 16)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Rodzaj kontaktu . . . . . . . . . . . . . "
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 96)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 16)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Data kontaktu . . . . . . . . ."
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
        Me.cmdWycofajSie.Location = New System.Drawing.Point(468, 382)
        Me.cmdWycofajSie.Name = "cmdWycofajSie"
        Me.cmdWycofajSie.Size = New System.Drawing.Size(104, 24)
        Me.cmdWycofajSie.TabIndex = 24
        Me.cmdWycofajSie.Text = "Wycofaj siê"
        Me.cmdWycofajSie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPrzywroc
        '
        Me.cmdPrzywroc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPrzywroc.Image = CType(resources.GetObject("cmdPrzywroc.Image"), System.Drawing.Image)
        Me.cmdPrzywroc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPrzywroc.Location = New System.Drawing.Point(370, 382)
        Me.cmdPrzywroc.Name = "cmdPrzywroc"
        Me.cmdPrzywroc.Size = New System.Drawing.Size(92, 24)
        Me.cmdPrzywroc.TabIndex = 25
        Me.cmdPrzywroc.Text = "Przywróæ"
        Me.cmdPrzywroc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(578, 382)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 24)
        Me.cmdPowrot.TabIndex = 23
        Me.cmdPowrot.Text = "Zapisz"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 382)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(88, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 26
        Me.PictureBox2.TabStop = False
        '
        'EdytujKatalogKontaktow
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(666, 409)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdWycofajSie)
        Me.Controls.Add(Me.cmdPrzywroc)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "EdytujKatalogKontaktow"
        Me.ShowInTaskbar = False
        Me.Text = "Edytuj Katalog Kontaktów Handlowych"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub EdytujKatalogKontaktow_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        PobierzDane()
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
            IdentUzytkownika(Program_IdUzytkownika)
            If OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaPracownik Then
                dtpDataKontaktu.Enabled = False
                txtPracownik.Enabled = False
                cmdPracownik.Enabled = False
                cmdRodzajKontaktu.Enabled = False
                txtUwagi.Enabled = False
            End If
        Else
            'txtData.Enabled = True
            'cmdData.Enabled = True
            'cmdKlient.Enabled = True
        End If
    End Sub

    Private Sub EdytujKatalogKontaktow_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
            txtPracownik.Focus()
        Else
            dtpDataKontaktu.Focus()
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

    Private Sub cmdPrzywroc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrzywroc.Click
        PobierzDane()
    End Sub

    Private Sub EdytujKatalogKontaktow_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdRodzajKontaktu_Click(sender As Object, e As EventArgs) Handles cmdRodzajKontaktu.Click
        oCoMamRobic = "WYBIERAJ"
        oRodzajKontaktu = Trim(txtRodzajKontaktu.Text)
        Dim form As New KatalogRodzajowKontaktow
        form.ShowDialog()
        If Len(Trim(oRodzajKontaktu)) > 0 Then
            txtRodzajKontaktu.Text = oRodzajKontaktu
        End If
    End Sub


    Private Sub cmdPracownik_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPracownik.Click
        oCoMamRobic = "WYBIERAJ"
        oPracownik = Trim(txtPracownik.Text)
        Dim form As New KatalogUzytkownikow
        form.ShowDialog()
        If Len(Trim(oPracownik)) > 0 Then
            txtPracownik.Text = oPracownik
        End If
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

    '====================================================
    'UWAGA - trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    Private Sub EdytujKatalogKontaktow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me.ClosedByControlBox Then
            'e.Cancel = True     'nie pozwalaj zamkn¹c forme
            If MessageBox.Show("Zdecydowa³eœ opuœciæ formê bez zapisania wprowadzonych zmian." & vbCrLf &
                "Czy mam zapisaæ zmiany w Bazie Danych ?",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
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
        lblIdent.Text = oIdentKon

        IdentKlienta(oIdentKon)
        lblNazwaHandlowa.Text = oNazwa1Kli
        lblNazwa1.Text = oUlicaKli
        lblNazwa2.Text = oKodPoczKli & "  " & oMiejscKli

        dtpDataKontaktu.Value = oDataKon
        txtPracownik.Text = oPracownikKon
        txtRodzajKontaktu.Text = oTematKon
        txtUwagi.Text = oUwagiKon
    End Sub

    Private Sub ZapiszDane()
        oIdentKon = lblIdent.Text
        oIdentKli = lblIdent.Text

        oDataKon = dtpDataKontaktu.Value.ToString("yyyy-MM-dd")
        oPracownikKon = txtPracownik.Text
        oTematKon = txtRodzajKontaktu.Text
        oUwagiKon = txtUwagi.Text
    End Sub

    '*******************************************
    '** procedury obslugi edycji
    '*******************************************

    'Private Sub txtData_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtData.TextChanged
    '    TestLen(txtData, "DATA KONTAKTU", 10)
    'End Sub
    'Private Sub txtData_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtData.GotFocus
    '    StartEdycjiTxt(txtData)
    'End Sub
    'Private Sub txtData_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtData.LostFocus
    '    KoniecEdycjiTxt(txtData)
    'End Sub


    Private Sub txtPracownik_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPracownik.TextChanged
        TestLen(txtPracownik, "PRACOWNIK", 10)
    End Sub
    Private Sub txtPracownik_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPracownik.GotFocus
        StartEdycjiTxt(txtPracownik)
    End Sub
    Private Sub txtPracownik_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPracownik.LostFocus
        KoniecEdycjiTxt(txtPracownik)
    End Sub


    Private Sub txtTemat_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRodzajKontaktu.TextChanged
        TestLen(txtRodzajKontaktu, "RODZAJ KONTAKTU", 50)
    End Sub

    Private Sub txtUwagi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.TextChanged
        'TestLen(txtUwagi, "NOTATKA", 50)
    End Sub
    Private Sub txtTemat_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRodzajKontaktu.GotFocus
        StartEdycjiTxt(txtRodzajKontaktu)
    End Sub
    Private Sub txtTemat_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRodzajKontaktu.LostFocus
        KoniecEdycjiTxt(txtRodzajKontaktu)
    End Sub


    Private Sub txtUwagi_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.GotFocus
        StartEdycjiTxt(txtUwagi)
    End Sub
    Private Sub txtUwagi_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.LostFocus
        KoniecEdycjiTxt(txtUwagi)
    End Sub
End Class
