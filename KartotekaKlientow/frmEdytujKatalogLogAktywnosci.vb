Public Class frmEdytujKatalogLogAktywnosci
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Dim _IdKlienta As String = ""

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _IdKlienta = ""
    End Sub

    Public Sub New(ByVal pIdklienta As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _IdKlienta = pIdklienta
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
    Friend WithEvents Identyfikacja As System.Windows.Forms.TabPage
    Friend WithEvents cmdWycofajSie As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents cmdeMail As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents chkVIP As System.Windows.Forms.CheckBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txteMail As System.Windows.Forms.TextBox
    Friend WithEvents txtTelefon As System.Windows.Forms.TextBox
    Friend WithEvents txtKompetencje As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtStanowisko As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtNazwisko As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtImiona As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cmdKlient As System.Windows.Forms.Button
    Friend WithEvents lblAdres2 As System.Windows.Forms.Label
    Friend WithEvents lblAdres1 As System.Windows.Forms.Label
    Friend WithEvents lblNazwa2 As System.Windows.Forms.Label
    Friend WithEvents lblNazwa1 As System.Windows.Forms.Label
    Friend WithEvents lblIdent As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdWybierzNumer As System.Windows.Forms.Button
    Friend WithEvents txtUwagi As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents chkUprawnionyOF As System.Windows.Forms.CheckBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEdytujKatalogLogAktywnosci))
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.Identyfikacja = New System.Windows.Forms.TabPage()
        Me.chkUprawnionyOF = New System.Windows.Forms.CheckBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtUwagi = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.chkVIP = New System.Windows.Forms.CheckBox()
        Me.cmdWybierzNumer = New System.Windows.Forms.Button()
        Me.txtImiona = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cmdKlient = New System.Windows.Forms.Button()
        Me.lblAdres2 = New System.Windows.Forms.Label()
        Me.lblAdres1 = New System.Windows.Forms.Label()
        Me.lblNazwa2 = New System.Windows.Forms.Label()
        Me.lblNazwa1 = New System.Windows.Forms.Label()
        Me.lblIdent = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdeMail = New System.Windows.Forms.Button()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txteMail = New System.Windows.Forms.TextBox()
        Me.txtTelefon = New System.Windows.Forms.TextBox()
        Me.txtKompetencje = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtStanowisko = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtNazwisko = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdWycofajSie = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.TabControl1.SuspendLayout()
        Me.Identyfikacja.SuspendLayout()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.Identyfikacja)
        Me.TabControl1.Location = New System.Drawing.Point(8, 8)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(550, 435)
        Me.TabControl1.TabIndex = 220
        '
        'Identyfikacja
        '
        Me.Identyfikacja.Controls.Add(Me.chkUprawnionyOF)
        Me.Identyfikacja.Controls.Add(Me.Label12)
        Me.Identyfikacja.Controls.Add(Me.txtUwagi)
        Me.Identyfikacja.Controls.Add(Me.Label11)
        Me.Identyfikacja.Controls.Add(Me.chkVIP)
        Me.Identyfikacja.Controls.Add(Me.cmdWybierzNumer)
        Me.Identyfikacja.Controls.Add(Me.txtImiona)
        Me.Identyfikacja.Controls.Add(Me.Label10)
        Me.Identyfikacja.Controls.Add(Me.cmdKlient)
        Me.Identyfikacja.Controls.Add(Me.lblAdres2)
        Me.Identyfikacja.Controls.Add(Me.lblAdres1)
        Me.Identyfikacja.Controls.Add(Me.lblNazwa2)
        Me.Identyfikacja.Controls.Add(Me.lblNazwa1)
        Me.Identyfikacja.Controls.Add(Me.lblIdent)
        Me.Identyfikacja.Controls.Add(Me.Label8)
        Me.Identyfikacja.Controls.Add(Me.Label7)
        Me.Identyfikacja.Controls.Add(Me.Label4)
        Me.Identyfikacja.Controls.Add(Me.cmdeMail)
        Me.Identyfikacja.Controls.Add(Me.Label9)
        Me.Identyfikacja.Controls.Add(Me.Panel2)
        Me.Identyfikacja.Controls.Add(Me.txteMail)
        Me.Identyfikacja.Controls.Add(Me.txtTelefon)
        Me.Identyfikacja.Controls.Add(Me.txtKompetencje)
        Me.Identyfikacja.Controls.Add(Me.Label6)
        Me.Identyfikacja.Controls.Add(Me.Label5)
        Me.Identyfikacja.Controls.Add(Me.Label3)
        Me.Identyfikacja.Controls.Add(Me.txtStanowisko)
        Me.Identyfikacja.Controls.Add(Me.Label2)
        Me.Identyfikacja.Controls.Add(Me.txtNazwisko)
        Me.Identyfikacja.Controls.Add(Me.Label1)
        Me.Identyfikacja.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Identyfikacja.ForeColor = System.Drawing.Color.Navy
        Me.Identyfikacja.Location = New System.Drawing.Point(4, 22)
        Me.Identyfikacja.Name = "Identyfikacja"
        Me.Identyfikacja.Size = New System.Drawing.Size(542, 409)
        Me.Identyfikacja.TabIndex = 0
        Me.Identyfikacja.Text = "Identyfikacja"
        '
        'chkUprawnionyOF
        '
        Me.chkUprawnionyOF.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkUprawnionyOF.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkUprawnionyOF.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.chkUprawnionyOF.ForeColor = System.Drawing.Color.Navy
        Me.chkUprawnionyOF.Location = New System.Drawing.Point(508, 138)
        Me.chkUprawnionyOF.Name = "chkUprawnionyOF"
        Me.chkUprawnionyOF.Size = New System.Drawing.Size(16, 16)
        Me.chkUprawnionyOF.TabIndex = 353
        '
        'Label12
        '
        Me.Label12.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label12.ForeColor = System.Drawing.Color.Navy
        Me.Label12.Location = New System.Drawing.Point(406, 138)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(118, 16)
        Me.Label12.TabIndex = 354
        Me.Label12.Text = "Upr. do Odb.Faktur . . . . "
        '
        'txtUwagi
        '
        Me.txtUwagi.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUwagi.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtUwagi.ForeColor = System.Drawing.Color.Purple
        Me.txtUwagi.Location = New System.Drawing.Point(124, 255)
        Me.txtUwagi.Multiline = True
        Me.txtUwagi.Name = "txtUwagi"
        Me.txtUwagi.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtUwagi.Size = New System.Drawing.Size(400, 139)
        Me.txtUwagi.TabIndex = 352
        '
        'Label11
        '
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(12, 255)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(140, 16)
        Me.Label11.TabIndex = 351
        Me.Label11.Text = "Opis osoby / Uwagi . . . . . . . . . . . . . . . . ."
        '
        'chkVIP
        '
        Me.chkVIP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkVIP.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkVIP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.chkVIP.ForeColor = System.Drawing.Color.Navy
        Me.chkVIP.Location = New System.Drawing.Point(508, 111)
        Me.chkVIP.Name = "chkVIP"
        Me.chkVIP.Size = New System.Drawing.Size(16, 16)
        Me.chkVIP.TabIndex = 27
        '
        'cmdWybierzNumer
        '
        Me.cmdWybierzNumer.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzNumer.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierzNumer.Image = CType(resources.GetObject("cmdWybierzNumer.Image"), System.Drawing.Image)
        Me.cmdWybierzNumer.Location = New System.Drawing.Point(485, 206)
        Me.cmdWybierzNumer.Name = "cmdWybierzNumer"
        Me.cmdWybierzNumer.Size = New System.Drawing.Size(39, 23)
        Me.cmdWybierzNumer.TabIndex = 350
        Me.cmdWybierzNumer.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtImiona
        '
        Me.txtImiona.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtImiona.ForeColor = System.Drawing.Color.Purple
        Me.txtImiona.Location = New System.Drawing.Point(124, 135)
        Me.txtImiona.Name = "txtImiona"
        Me.txtImiona.Size = New System.Drawing.Size(276, 20)
        Me.txtImiona.TabIndex = 120
        '
        'Label10
        '
        Me.Label10.ForeColor = System.Drawing.Color.Navy
        Me.Label10.Location = New System.Drawing.Point(12, 135)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(140, 16)
        Me.Label10.TabIndex = 121
        Me.Label10.Text = "Imiona . . . . . . . . . . . . . ."
        '
        'cmdKlient
        '
        Me.cmdKlient.Location = New System.Drawing.Point(212, 7)
        Me.cmdKlient.Name = "cmdKlient"
        Me.cmdKlient.Size = New System.Drawing.Size(24, 22)
        Me.cmdKlient.TabIndex = 118
        '
        'lblAdres2
        '
        Me.lblAdres2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAdres2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAdres2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblAdres2.ForeColor = System.Drawing.Color.Purple
        Me.lblAdres2.Location = New System.Drawing.Point(124, 80)
        Me.lblAdres2.Name = "lblAdres2"
        Me.lblAdres2.Size = New System.Drawing.Size(400, 16)
        Me.lblAdres2.TabIndex = 115
        '
        'lblAdres1
        '
        Me.lblAdres1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAdres1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAdres1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblAdres1.ForeColor = System.Drawing.Color.Purple
        Me.lblAdres1.Location = New System.Drawing.Point(124, 64)
        Me.lblAdres1.Name = "lblAdres1"
        Me.lblAdres1.Size = New System.Drawing.Size(400, 16)
        Me.lblAdres1.TabIndex = 114
        '
        'lblNazwa2
        '
        Me.lblNazwa2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa2.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa2.Location = New System.Drawing.Point(124, 48)
        Me.lblNazwa2.Name = "lblNazwa2"
        Me.lblNazwa2.Size = New System.Drawing.Size(400, 16)
        Me.lblNazwa2.TabIndex = 119
        '
        'lblNazwa1
        '
        Me.lblNazwa1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa1.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa1.Location = New System.Drawing.Point(124, 32)
        Me.lblNazwa1.Name = "lblNazwa1"
        Me.lblNazwa1.Size = New System.Drawing.Size(400, 16)
        Me.lblNazwa1.TabIndex = 113
        '
        'lblIdent
        '
        Me.lblIdent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdent.ForeColor = System.Drawing.Color.Purple
        Me.lblIdent.Location = New System.Drawing.Point(124, 8)
        Me.lblIdent.Name = "lblIdent"
        Me.lblIdent.Size = New System.Drawing.Size(112, 20)
        Me.lblIdent.TabIndex = 112
        Me.lblIdent.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label8
        '
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(12, 64)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(140, 16)
        Me.Label8.TabIndex = 117
        Me.Label8.Text = "Adres klienta . . . . . . . . ."
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(12, 32)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(140, 16)
        Me.Label7.TabIndex = 116
        Me.Label7.Text = "Nazwa klienta . . . . . . . . ."
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(12, 11)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(140, 16)
        Me.Label4.TabIndex = 111
        Me.Label4.Text = "Identyfikator klienta . . . . . . . . ."
        '
        'cmdeMail
        '
        Me.cmdeMail.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdeMail.ForeColor = System.Drawing.Color.Black
        Me.cmdeMail.Image = CType(resources.GetObject("cmdeMail.Image"), System.Drawing.Image)
        Me.cmdeMail.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdeMail.Location = New System.Drawing.Point(485, 230)
        Me.cmdeMail.Name = "cmdeMail"
        Me.cmdeMail.Size = New System.Drawing.Size(39, 23)
        Me.cmdeMail.TabIndex = 33
        Me.cmdeMail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label9
        '
        Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(470, 111)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(54, 16)
        Me.Label9.TabIndex = 46
        Me.Label9.Text = "V.I.P. . . . . "
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(4, 104)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(536, 1)
        Me.Panel2.TabIndex = 45
        '
        'txteMail
        '
        Me.txteMail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txteMail.ForeColor = System.Drawing.Color.Purple
        Me.txteMail.Location = New System.Drawing.Point(124, 231)
        Me.txteMail.Name = "txteMail"
        Me.txteMail.Size = New System.Drawing.Size(355, 20)
        Me.txteMail.TabIndex = 31
        '
        'txtTelefon
        '
        Me.txtTelefon.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTelefon.ForeColor = System.Drawing.Color.Purple
        Me.txtTelefon.Location = New System.Drawing.Point(124, 207)
        Me.txtTelefon.Name = "txtTelefon"
        Me.txtTelefon.Size = New System.Drawing.Size(355, 20)
        Me.txtTelefon.TabIndex = 30
        '
        'txtKompetencje
        '
        Me.txtKompetencje.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtKompetencje.ForeColor = System.Drawing.Color.Purple
        Me.txtKompetencje.Location = New System.Drawing.Point(124, 183)
        Me.txtKompetencje.Name = "txtKompetencje"
        Me.txtKompetencje.Size = New System.Drawing.Size(400, 20)
        Me.txtKompetencje.TabIndex = 29
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(12, 231)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(140, 16)
        Me.Label6.TabIndex = 38
        Me.Label6.Text = "Adres email . . . . . . . . . . . . . "
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(12, 207)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(140, 16)
        Me.Label5.TabIndex = 37
        Me.Label5.Text = "Telefon . . . . . . . . . . . . . "
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(12, 183)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(140, 16)
        Me.Label3.TabIndex = 36
        Me.Label3.Text = "Kompetencje . . . . . . . . . . . ."
        '
        'txtStanowisko
        '
        Me.txtStanowisko.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtStanowisko.ForeColor = System.Drawing.Color.Purple
        Me.txtStanowisko.Location = New System.Drawing.Point(124, 159)
        Me.txtStanowisko.Name = "txtStanowisko"
        Me.txtStanowisko.Size = New System.Drawing.Size(400, 20)
        Me.txtStanowisko.TabIndex = 28
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(12, 159)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(140, 16)
        Me.Label2.TabIndex = 35
        Me.Label2.Text = "Stanowisko / Funkcja . . . . . . . . ."
        '
        'txtNazwisko
        '
        Me.txtNazwisko.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNazwisko.ForeColor = System.Drawing.Color.Purple
        Me.txtNazwisko.Location = New System.Drawing.Point(124, 111)
        Me.txtNazwisko.Name = "txtNazwisko"
        Me.txtNazwisko.Size = New System.Drawing.Size(276, 20)
        Me.txtNazwisko.TabIndex = 26
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(12, 111)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(140, 16)
        Me.Label1.TabIndex = 34
        Me.Label1.Text = "Nazwisko . . . . . . . . . . . ."
        '
        'cmdWycofajSie
        '
        Me.cmdWycofajSie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWycofajSie.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWycofajSie.Location = New System.Drawing.Point(368, 451)
        Me.cmdWycofajSie.Name = "cmdWycofajSie"
        Me.cmdWycofajSie.Size = New System.Drawing.Size(103, 22)
        Me.cmdWycofajSie.TabIndex = 205
        Me.cmdWycofajSie.Text = "Wycofaj siê"
        Me.cmdWycofajSie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(478, 451)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 22)
        Me.cmdPowrot.TabIndex = 200
        Me.cmdPowrot.Text = "Zapisz"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.Location = New System.Drawing.Point(16, 451)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(88, 24)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 22
        Me.picLogo.TabStop = False
        '
        'frmEdytujKatalogLogAktywnosci
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(566, 481)
        Me.Controls.Add(Me.cmdWycofajSie)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "frmEdytujKatalogLogAktywnosci"
        Me.ShowInTaskbar = False
        Me.Text = "Informacje o Osobie Kontaktowej"
        Me.TabControl1.ResumeLayout(False)
        Me.Identyfikacja.ResumeLayout(False)
        Me.Identyfikacja.PerformLayout()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region



    '---------------------------------------------------------------------
    'Osoby Kontaktowe
    '---------------------------------------------------------------------
    'Public OsK_UniqId As String              'UNIQID           Text, 40
    'Public OsK_IdentKlienta As String        'IDENTKLIENTA     Text, 10
    'Public OsK_Imie As String                'IMIE             Text, 100
    'Public OsK_Nazwisko As String            'NAZWISKO         Text, 100
    'Public OsK_Stanowisko As String          'STANOWISKO       Text, 100
    'Public OsK_Kompetencje As String         'KOMPETENCJE      Text, 100
    'Public OsK_Telefon As String             'TELEFON          Text, 100
    'Public OsK_eMail As String               'EMAIL            Text, 100
    'Public OsK_VIP As Boolean                'VIP              boolean
    'Public OsK_UPRAWNIONYOF As Boolean        'UPRAWNIONYOF     boolean
    'Public OsK_Uwagi As String               'UWAGI            Memo
    'Public OsK_Wersja As Integer             'WERSJA           Integer

    '*********************************************************
    '** Inicjowanie
    '*********************************************************

    Private StartFormy As Boolean = True
    Dim _CoRobimy As String = ""

    Private Sub frmEdytujKatalogBankow_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.MrJones
        Me.picLogo.Image = My.Resources.logomini
        Me.cmdPowrot.Image = My.Resources._EXIT.ToBitmap
        Me.cmdWycofajSie.Image = My.Resources.ESCAPE.ToBitmap

        Me.cmdKlient.Image = My.Resources.SZUKAJ2.ToBitmap
        '---------------------------------------------
        If Len(_IdKlienta) > 0 Then
            'nie mozna edytowaæ klienta
            cmdKlient.Visible = False
        End If

        PobierzDane()
        _CoRobimy = oOperacja
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
        ElseIf oOperacja = "POKA¯" Then
        Else
        End If

        StartFormy = True
    End Sub

    Private Sub frmEdytujKatalogBankow_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If StartFormy Then
            'Select Case Par_IdentUzytkownika
            '    Case __UserTRES
            '        If TabControl1.TabCount >= 1 Then TylkoKapitaliki(Me.TabControl1.TabPages(0))
            '        If TabControl1.TabCount >= 2 Then TylkoKapitaliki(Me.TabControl1.TabPages(1))
            '        If TabControl1.TabCount >= 3 Then TylkoKapitaliki(Me.TabControl1.TabPages(2))
            '        If TabControl1.TabCount >= 4 Then TylkoKapitaliki(Me.TabControl1.TabPages(3))
            '        If TabControl1.TabCount >= 5 Then TylkoKapitaliki(Me.TabControl1.TabPages(4))
            '        If TabControl1.TabCount >= 6 Then TylkoKapitaliki(Me.TabControl1.TabPages(5))
            '        If TabControl1.TabCount >= 7 Then TylkoKapitaliki(Me.TabControl1.TabPages(6))
            '        If TabControl1.TabCount >= 8 Then TylkoKapitaliki(Me.TabControl1.TabPages(7))
            '        If TabControl1.TabCount >= 9 Then TylkoKapitaliki(Me.TabControl1.TabPages(8))
            '        If TabControl1.TabCount >= 10 Then TylkoKapitaliki(Me.TabControl1.TabPages(9))
            '        If TabControl1.TabCount >= 11 Then TylkoKapitaliki(Me.TabControl1.TabPages(10))
            '        If TabControl1.TabCount >= 12 Then TylkoKapitaliki(Me.TabControl1.TabPages(11))
            '        If TabControl1.TabCount >= 13 Then TylkoKapitaliki(Me.TabControl1.TabPages(12))
            '        If TabControl1.TabCount >= 14 Then TylkoKapitaliki(Me.TabControl1.TabPages(13))
            '        If TabControl1.TabCount >= 15 Then TylkoKapitaliki(Me.TabControl1.TabPages(14))
            '        If TabControl1.TabCount >= 16 Then TylkoKapitaliki(Me.TabControl1.TabPages(15))
            '        If TabControl1.TabCount >= 17 Then TylkoKapitaliki(Me.TabControl1.TabPages(16))
            '        If TabControl1.TabCount >= 18 Then TylkoKapitaliki(Me.TabControl1.TabPages(17))
            '        If TabControl1.TabCount >= 19 Then TylkoKapitaliki(Me.TabControl1.TabPages(18))
            '        If TabControl1.TabCount >= 20 Then TylkoKapitaliki(Me.TabControl1.TabPages(19))
            '        If TabControl1.TabCount >= 21 Then TylkoKapitaliki(Me.TabControl1.TabPages(20))
            '        If TabControl1.TabCount >= 22 Then TylkoKapitaliki(Me.TabControl1.TabPages(21))
            '        If TabControl1.TabCount >= 23 Then TylkoKapitaliki(Me.TabControl1.TabPages(22))

            'End Select

            Select Case _CoRobimy
                Case "POKA¯"
                    '------------Cursor.Current = Cursors.Default
                    'uniemo¿liwiæ edycje pól we wszystkich kontenerach
                    'NieEdytowac(Me.Panel1)
                    'NieEdytowac(Me.Panel2)
                    'NieEdytowac(Me.Panel3)
                    'NieEdytowac(Me.Panel4)
                    'NieEdytowac(Me.Panel5)
                    'NieEdytowac(Me.Panel6)
                    If TabControl1.TabCount >= 1 Then NieEdytowac(Me.TabControl1.TabPages(0))
                    If TabControl1.TabCount >= 2 Then NieEdytowac(Me.TabControl1.TabPages(1))
                    If TabControl1.TabCount >= 3 Then NieEdytowac(Me.TabControl1.TabPages(2))
                    If TabControl1.TabCount >= 4 Then NieEdytowac(Me.TabControl1.TabPages(3))
                    If TabControl1.TabCount >= 5 Then NieEdytowac(Me.TabControl1.TabPages(4))
                    If TabControl1.TabCount >= 6 Then NieEdytowac(Me.TabControl1.TabPages(5))
                    If TabControl1.TabCount >= 7 Then NieEdytowac(Me.TabControl1.TabPages(6))
                    If TabControl1.TabCount >= 8 Then NieEdytowac(Me.TabControl1.TabPages(7))
                    If TabControl1.TabCount >= 9 Then NieEdytowac(Me.TabControl1.TabPages(8))
                    If TabControl1.TabCount >= 10 Then NieEdytowac(Me.TabControl1.TabPages(9))
                    If TabControl1.TabCount >= 11 Then NieEdytowac(Me.TabControl1.TabPages(10))
                    If TabControl1.TabCount >= 12 Then NieEdytowac(Me.TabControl1.TabPages(11))
                    If TabControl1.TabCount >= 13 Then NieEdytowac(Me.TabControl1.TabPages(12))
                    If TabControl1.TabCount >= 14 Then NieEdytowac(Me.TabControl1.TabPages(13))
                    If TabControl1.TabCount >= 15 Then NieEdytowac(Me.TabControl1.TabPages(14))
                    If TabControl1.TabCount >= 16 Then NieEdytowac(Me.TabControl1.TabPages(15))
                    If TabControl1.TabCount >= 17 Then NieEdytowac(Me.TabControl1.TabPages(16))
                    If TabControl1.TabCount >= 18 Then NieEdytowac(Me.TabControl1.TabPages(17))
                    If TabControl1.TabCount >= 19 Then NieEdytowac(Me.TabControl1.TabPages(18))
                    If TabControl1.TabCount >= 20 Then NieEdytowac(Me.TabControl1.TabPages(19))
                    If TabControl1.TabCount >= 21 Then NieEdytowac(Me.TabControl1.TabPages(20))
                    If TabControl1.TabCount >= 22 Then NieEdytowac(Me.TabControl1.TabPages(21))
                    If TabControl1.TabCount >= 23 Then NieEdytowac(Me.TabControl1.TabPages(22))



                Case Else
                    'edycja pozycji cennika
                    'ustaw na pocz¹tek na górze....
            End Select

            'ustaw na pocz¹tek na górze....
            cmdPowrot.Focus()
            StartFormy = False
        Else
        End If
    End Sub

    Private Sub frmEdytujKatalogKlientów_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    '*********************************************************
    '** Komendy ekranowe
    '*********************************************************

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        'If Len(txtIdent.Text) = 0 Then
        '    MessageBox.Show("Proszê wpisaæ Osobê Kontaktow¹.", _
        '        "Uwaga:", _
        '        System.Windows.Forms.MessageBoxButtons.OK, _
        '        MessageBoxIcon.Information, _
        '        MessageBoxDefaultButton.Button1)
        'Else
        ZapiszDane()
        oAktualizuj = True
        Me.Close()
        'End If
    End Sub

    Private Sub cmdWycofajSie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWycofajSie.Click
        oAktualizuj = False
        Me.Close()
    End Sub


    '***************************************************
    '* procedury obslugi ekranu
    '***************************************************

    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
    End Sub


    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Select Case TabControl1.SelectedIndex
            Case 0
                cmdPowrot.Focus()
        End Select
    End Sub

    '***************************************************
    '* procedury pobierania/zapisywania danych
    '***************************************************

    Private Sub PobierzDane()
        'lblIdent.Text = OsK_IdentKlienta
        'IdentKlienta(OsK_IdentKlienta)
        'lblNazwa1.Text = K_Nazwa1Kli
        'lblNazwa2.Text = K_Nazwa2Kli
        'lblAdres1.Text = K_KodPoczKli & " " & K_MiejscKli
        'lblAdres2.Text = K_UlicaKli

        'chkVIP.Checked = OsK_VIP
        'chkUprawnionyOF.Checked = OsK_UPRAWNIONYOF
        'txtNazwisko.Text = OsK_Nazwisko
        'txtImiona.Text = OsK_Imie
        'txtStanowisko.Text = OsK_Stanowisko
        'txtKompetencje.Text = OsK_Kompetencje
        'txtTelefon.Text = OsK_Telefon
        'txteMail.Text = OsK_eMail
        'txtUwagi.Text = OsK_Uwagi
    End Sub

    Private Sub ZapiszDane()
        'OsK_IdentKlienta = lblIdent.Text
        'OsK_VIP = chkVIP.Checked
        'OsK_UprawnionyOF = chkUprawnionyOF.Checked
        'OsK_Nazwisko = txtNazwisko.Text
        'OsK_Imie = txtImiona.Text
        'OsK_Stanowisko = txtStanowisko.Text
        'OsK_Kompetencje = txtKompetencje.Text
        'OsK_Telefon = txtTelefon.Text
        'OsK_eMail = txteMail.Text
        'OsK_Uwagi = txtUwagi.Text
    End Sub


    '*******************************************
    '** procedury obslugi edycji
    '*******************************************

    '---------------------------------------------------------------------
    'Osoby Kontaktowe
    '---------------------------------------------------------------------
    'Public OsK_UniqId As String              'UNIQID           Text, 40
    'Public OsK_IdentKlienta As String        'IDENTKLIENTA     Text, 10
    'Public OsK_Imie As String                'IMIE             Text, 100
    'Public OsK_Nazwisko As String            'NAZWISKO         Text, 100
    'Public OsK_Stanowisko As String          'STANOWISKO       Text, 100
    'Public OsK_Kompetencje As String         'KOMPETENCJE      Text, 100
    'Public OsK_Telefon As String             'TELEFON          Text, 100
    'Public OsK_eMail As String               'EMAIL            Text, 100
    'Public OsK_VIP As Boolean                'VIP              boolean
    'Public OsK_UprawnionyOF As Boolean       'UPRAWNIONYOF     boolean
    'Public OsK_Uwagi As String               'UWAGI            Memo
    'Public OsK_Wersja As Integer             'WERSJA           Integer

    Private Sub txtNazwisko_TextChanged(sender As Object, e As EventArgs) Handles txtNazwisko.TextChanged
        TestLen(txtNazwisko, "Nazwisko", 100)
    End Sub
    Private Sub txtNazwisko_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNazwisko.GotFocus
        StartEdycjiTxt(txtNazwisko)
    End Sub
    Private Sub txtNazwisko_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNazwisko.LostFocus
        KoniecEdycjiTxt(txtNazwisko)
    End Sub

    Private Sub txtImiona_TextChanged(sender As Object, e As EventArgs) Handles txtImiona.TextChanged
        TestLen(txtImiona, "Imiona", 100)
    End Sub
    Private Sub txtImiona_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtImiona.GotFocus
        StartEdycjiTxt(txtImiona)
    End Sub
    Private Sub txtImiona_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtImiona.LostFocus
        KoniecEdycjiTxt(txtImiona)
    End Sub

    Private Sub txtStanowisko_TextChanged(sender As Object, e As EventArgs) Handles txtStanowisko.TextChanged
        TestLen(txtStanowisko, "Stanowisko", 100)
    End Sub
    Private Sub txtStanowisko_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStanowisko.GotFocus
        StartEdycjiTxt(txtStanowisko)
    End Sub
    Private Sub txtStanowisko_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStanowisko.LostFocus
        KoniecEdycjiTxt(txtStanowisko)
    End Sub

    Private Sub txtKompetencje_TextChanged(sender As Object, e As EventArgs) Handles txtKompetencje.TextChanged
        TestLen(txtKompetencje, "Kompetencje", 100)
    End Sub
    Private Sub txtKompetencje_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKompetencje.GotFocus
        StartEdycjiTxt(txtKompetencje)
    End Sub
    Private Sub txtKompetencje_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKompetencje.LostFocus
        KoniecEdycjiTxt(txtKompetencje)
    End Sub

    Private Sub txtTelefon_TextChanged(sender As Object, e As EventArgs) Handles txtTelefon.TextChanged
        TestLen(txtTelefon, "Telefon", 100)
    End Sub
    Private Sub txtTelefon_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTelefon.GotFocus
        StartEdycjiTxt(txtTelefon)
    End Sub
    Private Sub txtTelefon_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTelefon.LostFocus
        KoniecEdycjiTxt(txtTelefon)
    End Sub

    Private Sub txteMail_TextChanged(sender As Object, e As EventArgs) Handles txteMail.TextChanged
        TestLen(txteMail, "eMail", 100)
    End Sub
    Private Sub txteMail_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txteMail.GotFocus
        StartEdycjiTxt(txteMail)
    End Sub
    Private Sub txteMail_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txteMail.LostFocus
        KoniecEdycjiTxt(txteMail)
    End Sub

    Private Sub txtUwagi_TextChanged(sender As Object, e As EventArgs) Handles txtUwagi.TextChanged
        'TestLen(txtUwagi, "Uwagi", 100)
    End Sub
    Private Sub txtUwagi_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.GotFocus
        StartEdycjiTxt(txtUwagi)
    End Sub
    Private Sub txtUwagi_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagi.LostFocus
        KoniecEdycjiTxt(txtUwagi)
    End Sub


    '*******************************************
    '** klawisze wyboru
    '*******************************************

    Private Sub cmdKlient_Click(sender As Object, e As EventArgs) Handles cmdKlient.Click
        'oCoMamRobic = "WYBIERAJ"
        'oKlient = Trim(lblIdent.Text)
        'Dim form As New frmKatalogKlientow
        'Using form
        '    form.ShowDialog()
        'End Using
        'form = Nothing

        'If Len(Trim(oKlient)) > 0 Then
        '    lblIdent.Text = oKlient
        '    IdentKlienta(oKlient)
        '    lblNazwa1.Text = K_Nazwa1Kli
        '    lblNazwa2.Text = K_Nazwa2Kli
        '    lblAdres1.Text = K_KodPoczKli & " " & K_MiejscKli
        '    lblAdres2.Text = K_UlicaKli
        'End If
    End Sub

    Private Sub cmdeMail_Click(sender As Object, e As EventArgs) Handles cmdeMail.Click
        'If Len(txteMail.Text) Then
        '    If TestujParametryeMail() Then

        '        'cbxeMailService.Items.Add("Wbudowany")
        '        'cbxeMailService.Items.Add("MS Outlook")
        '        'cbxeMailService.Items.Add("Mozilla Thunderbird")
        '        'cbxeMailService.Items.Add("Google GMail")
        '        'cbxeMailService.Items.Add("Poczta wp.pl")
        '        'cbxeMailService.Items.Add("Poczta home.pl")
        '        'cbxeMailService.Items.Add("Poczta onet.pl")
        '        'cbxeMailService.Items.Add("Poczta interia.pl")
        '        'cbxeMailService.Items.Add("Poczta o2.pl")

        '        Select Case ParL_eMailService
        '            Case "Wbudowany"
        '                Dim FormW As New _frmSendeMail(txteMail.Text, "", "", "")
        '                FormW.Show()
        '                FormW = Nothing
        '            Case "MS Outlook"
        '                If SendByOutlook(txteMail.Text, "", "", "", "", "", "POJEDYNCZA") Then
        '                End If

        '            Case "Mozilla Thunderbird"
        '                If SendByThunderbird(txteMail.Text, "", "", "", "", "", "POJEDYNCZA") Then
        '                End If

        '            Case "Google GMail", "Poczta wp.pl", "Poczta home.pl", "Poczta onet.pl", "Poczta interia.pl", "Poczta o2.pl"
        '                Dim FormG As New _frmSendeMail(txteMail.Text, "", "", "")
        '                FormG.Show()
        '                FormG = Nothing
        '            Case Else
        '        End Select
        '    End If
        'End If
    End Sub

    Private Sub cmdWybierzNumer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierzNumer.Click
        'If Len(txtTelefon.Text) > 0 Then
        '    If PhoneCall(Trim(txtTelefon.Text), "", "", "") >= 0 Then
        '        'MessageBox.Show("£¹czê z numerem " & vbCrLf & Trim(txtTelefon.Text), "OK", _
        '        '    System.Windows.Forms.MessageBoxButtons.OK, _
        '        '    MessageBoxIcon.Information, _
        '        '    MessageBoxDefaultButton.Button1)
        '    Else
        '        MessageBox.Show("B³¹d po³¹czenia z numerem " & vbCrLf & Trim(txtTelefon.Text), "Uwaga :", _
        '            System.Windows.Forms.MessageBoxButtons.OK, _
        '            MessageBoxIcon.Exclamation, _
        '            MessageBoxDefaultButton.Button1)
        '    End If
        'End If
    End Sub

End Class
