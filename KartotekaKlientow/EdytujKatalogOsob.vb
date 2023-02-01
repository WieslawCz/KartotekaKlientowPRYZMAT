Public Class EdytujKatalogOsob
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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdWycofajSie As System.Windows.Forms.Button
    Friend WithEvents cmdPrzywroc As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblIdent As System.Windows.Forms.Label
    Friend WithEvents lblNazwa2 As System.Windows.Forms.Label
    Friend WithEvents lblNazwa1 As System.Windows.Forms.Label
    Friend WithEvents lblNazwaHandlowa As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txteMail As System.Windows.Forms.TextBox
    Friend WithEvents txtTelefon As System.Windows.Forms.TextBox
    Friend WithEvents txtKompetencje As System.Windows.Forms.TextBox
    Friend WithEvents txtStanowisko As System.Windows.Forms.TextBox
    Friend WithEvents txtOsoba As System.Windows.Forms.TextBox
    Friend WithEvents chbVIP As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtTelefonKom As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EdytujKatalogOsob))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtTelefonKom = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.chbVIP = New System.Windows.Forms.CheckBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblNazwa2 = New System.Windows.Forms.Label()
        Me.lblNazwa1 = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa = New System.Windows.Forms.Label()
        Me.lblIdent = New System.Windows.Forms.Label()
        Me.txteMail = New System.Windows.Forms.TextBox()
        Me.txtTelefon = New System.Windows.Forms.TextBox()
        Me.txtKompetencje = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtStanowisko = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtOsoba = New System.Windows.Forms.TextBox()
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
        Me.Panel1.Controls.Add(Me.txtTelefonKom)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.chbVIP)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.lblNazwa2)
        Me.Panel1.Controls.Add(Me.lblNazwa1)
        Me.Panel1.Controls.Add(Me.lblNazwaHandlowa)
        Me.Panel1.Controls.Add(Me.lblIdent)
        Me.Panel1.Controls.Add(Me.txteMail)
        Me.Panel1.Controls.Add(Me.txtTelefon)
        Me.Panel1.Controls.Add(Me.txtKompetencje)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.txtStanowisko)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.txtOsoba)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(546, 256)
        Me.Panel1.TabIndex = 2
        '
        'txtTelefonKom
        '
        Me.txtTelefonKom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTelefonKom.ForeColor = System.Drawing.Color.Purple
        Me.txtTelefonKom.Location = New System.Drawing.Point(120, 192)
        Me.txtTelefonKom.Name = "txtTelefonKom"
        Me.txtTelefonKom.Size = New System.Drawing.Size(401, 20)
        Me.txtTelefonKom.TabIndex = 26
        '
        'Label10
        '
        Me.Label10.ForeColor = System.Drawing.Color.Navy
        Me.Label10.Location = New System.Drawing.Point(8, 192)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(112, 16)
        Me.Label10.TabIndex = 27
        Me.Label10.Text = "Telefon komórkowy . . . . . . . . . . . . . "
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(482, 214)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(39, 23)
        Me.Button1.TabIndex = 7
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label9
        '
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(464, 96)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(32, 10)
        Me.Label9.TabIndex = 25
        Me.Label9.Text = "V.I.P."
        '
        'chbVIP
        '
        Me.chbVIP.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chbVIP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.chbVIP.ForeColor = System.Drawing.Color.Navy
        Me.chbVIP.Location = New System.Drawing.Point(504, 96)
        Me.chbVIP.Name = "chbVIP"
        Me.chbVIP.Size = New System.Drawing.Size(16, 16)
        Me.chbVIP.TabIndex = 2
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(0, 88)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(541, 1)
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
        Me.Label7.Text = "Nazwa Klienta . . . . . . . . ."
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
        Me.lblNazwa2.Size = New System.Drawing.Size(401, 16)
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
        Me.lblNazwa1.Size = New System.Drawing.Size(401, 16)
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
        Me.lblNazwaHandlowa.Size = New System.Drawing.Size(401, 16)
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
        'txteMail
        '
        Me.txteMail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txteMail.ForeColor = System.Drawing.Color.Purple
        Me.txteMail.Location = New System.Drawing.Point(120, 216)
        Me.txteMail.Name = "txteMail"
        Me.txteMail.Size = New System.Drawing.Size(356, 20)
        Me.txteMail.TabIndex = 6
        '
        'txtTelefon
        '
        Me.txtTelefon.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTelefon.ForeColor = System.Drawing.Color.Purple
        Me.txtTelefon.Location = New System.Drawing.Point(120, 168)
        Me.txtTelefon.Name = "txtTelefon"
        Me.txtTelefon.Size = New System.Drawing.Size(401, 20)
        Me.txtTelefon.TabIndex = 5
        '
        'txtKompetencje
        '
        Me.txtKompetencje.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtKompetencje.ForeColor = System.Drawing.Color.Purple
        Me.txtKompetencje.Location = New System.Drawing.Point(120, 144)
        Me.txtKompetencje.Name = "txtKompetencje"
        Me.txtKompetencje.Size = New System.Drawing.Size(401, 20)
        Me.txtKompetencje.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(8, 216)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 16)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Adres email . . . . . . . . . . . . . "
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(8, 168)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 16)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Telefon s³u¿bowy . . . . . . . . . . . . . "
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 144)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Kompetencje . . . . . . . . . . . ."
        '
        'txtStanowisko
        '
        Me.txtStanowisko.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtStanowisko.ForeColor = System.Drawing.Color.Purple
        Me.txtStanowisko.Location = New System.Drawing.Point(120, 120)
        Me.txtStanowisko.Name = "txtStanowisko"
        Me.txtStanowisko.Size = New System.Drawing.Size(401, 20)
        Me.txtStanowisko.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 120)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 16)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Stanowisko/Funkcja . . . . . . . . ."
        '
        'txtOsoba
        '
        Me.txtOsoba.ForeColor = System.Drawing.Color.Purple
        Me.txtOsoba.Location = New System.Drawing.Point(120, 96)
        Me.txtOsoba.Name = "txtOsoba"
        Me.txtOsoba.Size = New System.Drawing.Size(320, 20)
        Me.txtOsoba.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 96)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 16)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Nazwisko i Imie . . . . . . . . ."
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Ident. klienta . . . . . . . . ."
        '
        'cmdWycofajSie
        '
        Me.cmdWycofajSie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWycofajSie.Image = CType(resources.GetObject("cmdWycofajSie.Image"), System.Drawing.Image)
        Me.cmdWycofajSie.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWycofajSie.Location = New System.Drawing.Point(366, 271)
        Me.cmdWycofajSie.Name = "cmdWycofajSie"
        Me.cmdWycofajSie.Size = New System.Drawing.Size(104, 25)
        Me.cmdWycofajSie.TabIndex = 9
        Me.cmdWycofajSie.Text = "Wycofaj siê"
        Me.cmdWycofajSie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPrzywroc
        '
        Me.cmdPrzywroc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPrzywroc.Image = CType(resources.GetObject("cmdPrzywroc.Image"), System.Drawing.Image)
        Me.cmdPrzywroc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPrzywroc.Location = New System.Drawing.Point(269, 271)
        Me.cmdPrzywroc.Name = "cmdPrzywroc"
        Me.cmdPrzywroc.Size = New System.Drawing.Size(91, 25)
        Me.cmdPrzywroc.TabIndex = 8
        Me.cmdPrzywroc.Text = "Przywróæ"
        Me.cmdPrzywroc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(476, 271)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 25)
        Me.cmdPowrot.TabIndex = 10
        Me.cmdPowrot.Text = "Zapisz"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 271)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(88, 25)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 21
        Me.PictureBox2.TabStop = False
        '
        'EdytujKatalogOsob
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(562, 302)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdWycofajSie)
        Me.Controls.Add(Me.cmdPrzywroc)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Name = "EdytujKatalogOsob"
        Me.ShowInTaskbar = False
        Me.Text = "Edytuj Katalog osób Kontaktowych "
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub EdytujKatalogOsob_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        PobierzDane()
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
            txtOsoba.Enabled = False
        Else
            txtOsoba.Enabled = True
        End If

        lblIdent.Text = oIdentKli
        'lblTOFI.Text = oNrTOFITxtKli
        lblNazwaHandlowa.Text = oNazwa1Kli
        lblNazwa1.Text = oUlicaKli
        lblNazwa2.Text = oKodPoczKli & "  " & oMiejscKli
    End Sub

    Private Sub EdytujKatalogOsob_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
            txtStanowisko.Focus()
        Else
            txtOsoba.Focus()
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

    '====================================================
    'UWAGA - trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    Private Sub EdytujKatalogOsob_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        txtOsoba.Text = oOsobaOso
        txtStanowisko.Text = oStanowiskoOso
        txtKompetencje.Text = oKompetencjeOso
        txtTelefon.Text = oTelefonOso
        txtTelefonKom.Text = oTelefonKomOso
        txteMail.Text = oeMailOso
        chbVIP.Checked = oVIPOso
    End Sub

    Private Sub ZapiszDane()
        oOsobaOso = txtOsoba.Text
        oStanowiskoOso = txtStanowisko.Text
        oKompetencjeOso = txtKompetencje.Text
        oTelefonOso = txtTelefon.Text
        oTelefonKomOso = txtTelefonKom.Text
        oeMailOso = txteMail.Text
        oVIPOso = chbVIP.Checked
    End Sub

    '*******************************************
    '** procedury obslugi edycji
    '*******************************************

    Private Sub txtOsoba_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOsoba.TextChanged
        TestLen(txtOsoba, "NAZWISKO I IMIÊ", 100)
    End Sub

    Private Sub txtStanowisko_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStanowisko.TextChanged
        TestLen(txtStanowisko, "STANOWISKO", 100)
    End Sub

    Private Sub txtKompetencje_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKompetencje.TextChanged
        TestLen(txtKompetencje, "KOMPETENCJE", 100)
    End Sub

    Private Sub txtTelefon_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTelefon.TextChanged
        TestLen(txtTelefon, "TELEFON S£U¯BOWY", 100)
    End Sub

    Private Sub txtTelefonKom_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTelefonKom.TextChanged
        TestLen(txtTelefonKom, "TELEFON KOMÓRKOWY", 100)
    End Sub

    Private Sub txteMail_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txteMail.TextChanged
        TestLen(txteMail, "EMAIL", 100)
    End Sub

    '**************************************************************
    '** obsluga obiektu w trakcie edycji
    '**************************************************************

    Private Sub txtOsoba_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOsoba.GotFocus
        StartEdycjiTxt(txtOsoba)
    End Sub
    Private Sub txtOsoba_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOsoba.LostFocus
        KoniecEdycjiTxt(txtOsoba)
    End Sub

    Private Sub txtStanowisko_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStanowisko.GotFocus
        StartEdycjiTxt(txtStanowisko)
    End Sub
    Private Sub txtStanowisko_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStanowisko.LostFocus
        KoniecEdycjiTxt(txtStanowisko)
    End Sub

    Private Sub txtKompetencje_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKompetencje.GotFocus
        StartEdycjiTxt(txtKompetencje)
    End Sub
    Private Sub txtKompetencje_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKompetencje.LostFocus
        KoniecEdycjiTxt(txtKompetencje)
    End Sub

    Private Sub txtTelefon_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTelefon.GotFocus
        StartEdycjiTxt(txtTelefon)
    End Sub
    Private Sub txtTelefon_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTelefon.LostFocus
        KoniecEdycjiTxt(txtTelefon)
    End Sub

    Private Sub txtTelefonKom_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTelefonKom.GotFocus
        StartEdycjiTxt(txtTelefonKom)
    End Sub
    Private Sub txtTelefonKom_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTelefonKom.LostFocus
        KoniecEdycjiTxt(txtTelefonKom)
    End Sub

    Private Sub txteMail_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txteMail.GotFocus
        StartEdycjiTxt(txteMail)
    End Sub
    Private Sub txteMail_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txteMail.LostFocus
        KoniecEdycjiTxt(txteMail)
    End Sub

    Private Sub chbVIP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles chbVIP.GotFocus
        StartEdycjiChb(chbVIP)
    End Sub
    Private Sub chbVIP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles chbVIP.LostFocus
        KoniecEdycjiChb(chbVIP)
    End Sub

    Private Sub EdytujKatalogOsob_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TestujParametryeMail() Then
            Select Case ParL_eMailService
                Case "Wbudowany"
                    Dim Form As New SendeMail(Trim(txteMail.Text), "", "", "")
                    Form.Show()
                Case "MS Outlook"
                    SendByOutlook(Trim(txteMail.Text), "", "", "", "Edytuj")
                Case "Mozilla Thunderbird"
                    SendByThunderbird(txteMail.Text, "", "", "", "", "")

                Case Else
            End Select
            'Dim Form As New SendeMail(Trim(txteMail.Text), "", "", "")
            'Form.Show()
        End If
    End Sub
End Class
