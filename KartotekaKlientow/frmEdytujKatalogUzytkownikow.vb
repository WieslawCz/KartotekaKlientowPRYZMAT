Public Class frmEdytujKatalogUzytkownikow
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents cmdWycofajSie As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents txtIdent As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtNazwa As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents tabPracownik As System.Windows.Forms.TabPage
    Friend WithEvents txtOpis As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TabUprawnienia As System.Windows.Forms.TabPage
    Friend WithEvents txtHaslo As System.Windows.Forms.TextBox
    Friend WithEvents lblHaslo As System.Windows.Forms.Label
    Friend WithEvents cmdPokazHaslo As System.Windows.Forms.Button
    Friend WithEvents txtHaslo2 As System.Windows.Forms.TextBox
    Friend WithEvents lblHaslo2 As System.Windows.Forms.Label
    Friend WithEvents rbtPracownik As System.Windows.Forms.RadioButton
    Friend WithEvents rbtSzef As System.Windows.Forms.RadioButton
    Friend WithEvents rbtAdministrator As System.Windows.Forms.RadioButton
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents rbtPracownikUprzywilejowany As System.Windows.Forms.RadioButton
    Friend WithEvents txtFunkcja As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEdytujKatalogUzytkownikow))
        Me.cmdWycofajSie = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.txtHaslo2 = New System.Windows.Forms.TextBox()
        Me.lblHaslo2 = New System.Windows.Forms.Label()
        Me.cmdPokazHaslo = New System.Windows.Forms.Button()
        Me.txtHaslo = New System.Windows.Forms.TextBox()
        Me.lblHaslo = New System.Windows.Forms.Label()
        Me.txtOpis = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtFunkcja = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtNazwa = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtIdent = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.tabPracownik = New System.Windows.Forms.TabPage()
        Me.TabUprawnienia = New System.Windows.Forms.TabPage()
        Me.rbtPracownikUprzywilejowany = New System.Windows.Forms.RadioButton()
        Me.rbtPracownik = New System.Windows.Forms.RadioButton()
        Me.rbtSzef = New System.Windows.Forms.RadioButton()
        Me.rbtAdministrator = New System.Windows.Forms.RadioButton()
        Me.Label7 = New System.Windows.Forms.Label()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.tabPracownik.SuspendLayout()
        Me.TabUprawnienia.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdWycofajSie
        '
        Me.cmdWycofajSie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWycofajSie.Image = CType(resources.GetObject("cmdWycofajSie.Image"), System.Drawing.Image)
        Me.cmdWycofajSie.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWycofajSie.Location = New System.Drawing.Point(637, 406)
        Me.cmdWycofajSie.Name = "cmdWycofajSie"
        Me.cmdWycofajSie.Size = New System.Drawing.Size(104, 23)
        Me.cmdWycofajSie.TabIndex = 5
        Me.cmdWycofajSie.Text = "Wycofaj siê"
        Me.cmdWycofajSie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(747, 406)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 23)
        Me.cmdPowrot.TabIndex = 4
        Me.cmdPowrot.Text = "Zapisz"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.Location = New System.Drawing.Point(16, 407)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(88, 25)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 16
        Me.picLogo.TabStop = False
        '
        'txtHaslo2
        '
        Me.txtHaslo2.ForeColor = System.Drawing.Color.Purple
        Me.txtHaslo2.Location = New System.Drawing.Point(432, 37)
        Me.txtHaslo2.Name = "txtHaslo2"
        Me.txtHaslo2.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtHaslo2.Size = New System.Drawing.Size(144, 20)
        Me.txtHaslo2.TabIndex = 209
        Me.txtHaslo2.Text = "Data"
        '
        'lblHaslo2
        '
        Me.lblHaslo2.ForeColor = System.Drawing.Color.Navy
        Me.lblHaslo2.Location = New System.Drawing.Point(314, 40)
        Me.lblHaslo2.Name = "lblHaslo2"
        Me.lblHaslo2.Size = New System.Drawing.Size(147, 16)
        Me.lblHaslo2.TabIndex = 210
        Me.lblHaslo2.Text = "Powtórz Has³o . . . . . . . . ."
        '
        'cmdPokazHaslo
        '
        Me.cmdPokazHaslo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPokazHaslo.Location = New System.Drawing.Point(582, 13)
        Me.cmdPokazHaslo.Name = "cmdPokazHaslo"
        Me.cmdPokazHaslo.Size = New System.Drawing.Size(80, 23)
        Me.cmdPokazHaslo.TabIndex = 208
        Me.cmdPokazHaslo.Text = "Poka¿ has³o"
        '
        'txtHaslo
        '
        Me.txtHaslo.ForeColor = System.Drawing.Color.Purple
        Me.txtHaslo.Location = New System.Drawing.Point(432, 15)
        Me.txtHaslo.Name = "txtHaslo"
        Me.txtHaslo.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtHaslo.Size = New System.Drawing.Size(144, 20)
        Me.txtHaslo.TabIndex = 13
        Me.txtHaslo.Text = "Data"
        '
        'lblHaslo
        '
        Me.lblHaslo.ForeColor = System.Drawing.Color.Navy
        Me.lblHaslo.Location = New System.Drawing.Point(314, 18)
        Me.lblHaslo.Name = "lblHaslo"
        Me.lblHaslo.Size = New System.Drawing.Size(147, 16)
        Me.lblHaslo.TabIndex = 14
        Me.lblHaslo.Text = "Has³o u¿ytkownika . . . . . . . . ."
        '
        'txtOpis
        '
        Me.txtOpis.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOpis.ForeColor = System.Drawing.Color.Purple
        Me.txtOpis.Location = New System.Drawing.Point(131, 103)
        Me.txtOpis.Multiline = True
        Me.txtOpis.Name = "txtOpis"
        Me.txtOpis.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtOpis.Size = New System.Drawing.Size(670, 247)
        Me.txtOpis.TabIndex = 11
        Me.txtOpis.Text = "Data"
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(13, 106)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(132, 16)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Uwagi . . . . . . . . . . . . . . . . . ."
        '
        'txtFunkcja
        '
        Me.txtFunkcja.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFunkcja.ForeColor = System.Drawing.Color.Purple
        Me.txtFunkcja.Location = New System.Drawing.Point(131, 81)
        Me.txtFunkcja.Name = "txtFunkcja"
        Me.txtFunkcja.Size = New System.Drawing.Size(670, 20)
        Me.txtFunkcja.TabIndex = 3
        Me.txtFunkcja.Text = "Data"
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(13, 84)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(132, 16)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Stanowisko/Funkcja . . . . . . . . ."
        '
        'txtNazwa
        '
        Me.txtNazwa.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNazwa.ForeColor = System.Drawing.Color.Purple
        Me.txtNazwa.Location = New System.Drawing.Point(131, 59)
        Me.txtNazwa.Name = "txtNazwa"
        Me.txtNazwa.Size = New System.Drawing.Size(670, 20)
        Me.txtNazwa.TabIndex = 2
        Me.txtNazwa.Text = "Data"
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(13, 62)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(141, 16)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Imiê i Nazwisko . . . . . . . . ."
        '
        'txtIdent
        '
        Me.txtIdent.ForeColor = System.Drawing.Color.Purple
        Me.txtIdent.Location = New System.Drawing.Point(131, 15)
        Me.txtIdent.Name = "txtIdent"
        Me.txtIdent.Size = New System.Drawing.Size(144, 20)
        Me.txtIdent.TabIndex = 1
        Me.txtIdent.Text = "Data"
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(13, 18)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(132, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Ident. u¿ytkownika . . . . . . . . ."
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.tabPracownik)
        Me.TabControl1.Controls.Add(Me.TabUprawnienia)
        Me.TabControl1.Location = New System.Drawing.Point(8, 8)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(819, 392)
        Me.TabControl1.TabIndex = 18
        '
        'tabPracownik
        '
        Me.tabPracownik.BackColor = System.Drawing.SystemColors.Control
        Me.tabPracownik.Controls.Add(Me.txtHaslo2)
        Me.tabPracownik.Controls.Add(Me.txtHaslo)
        Me.tabPracownik.Controls.Add(Me.txtIdent)
        Me.tabPracownik.Controls.Add(Me.txtNazwa)
        Me.tabPracownik.Controls.Add(Me.txtFunkcja)
        Me.tabPracownik.Controls.Add(Me.txtOpis)
        Me.tabPracownik.Controls.Add(Me.Label4)
        Me.tabPracownik.Controls.Add(Me.Label1)
        Me.tabPracownik.Controls.Add(Me.lblHaslo2)
        Me.tabPracownik.Controls.Add(Me.Label2)
        Me.tabPracownik.Controls.Add(Me.cmdPokazHaslo)
        Me.tabPracownik.Controls.Add(Me.Label3)
        Me.tabPracownik.Controls.Add(Me.lblHaslo)
        Me.tabPracownik.Location = New System.Drawing.Point(4, 22)
        Me.tabPracownik.Name = "tabPracownik"
        Me.tabPracownik.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPracownik.Size = New System.Drawing.Size(811, 366)
        Me.tabPracownik.TabIndex = 0
        Me.tabPracownik.Text = "Pracownik"
        '
        'TabUprawnienia
        '
        Me.TabUprawnienia.BackColor = System.Drawing.SystemColors.Control
        Me.TabUprawnienia.Controls.Add(Me.rbtPracownikUprzywilejowany)
        Me.TabUprawnienia.Controls.Add(Me.rbtPracownik)
        Me.TabUprawnienia.Controls.Add(Me.rbtSzef)
        Me.TabUprawnienia.Controls.Add(Me.rbtAdministrator)
        Me.TabUprawnienia.Controls.Add(Me.Label7)
        Me.TabUprawnienia.Location = New System.Drawing.Point(4, 22)
        Me.TabUprawnienia.Name = "TabUprawnienia"
        Me.TabUprawnienia.Size = New System.Drawing.Size(811, 366)
        Me.TabUprawnienia.TabIndex = 2
        Me.TabUprawnienia.Text = "Uprawnienia"
        '
        'rbtPracownikUprzywilejowany
        '
        Me.rbtPracownikUprzywilejowany.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.rbtPracownikUprzywilejowany.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtPracownikUprzywilejowany.ForeColor = System.Drawing.Color.Navy
        Me.rbtPracownikUprzywilejowany.Location = New System.Drawing.Point(14, 70)
        Me.rbtPracownikUprzywilejowany.Name = "rbtPracownikUprzywilejowany"
        Me.rbtPracownikUprzywilejowany.Size = New System.Drawing.Size(731, 20)
        Me.rbtPracownikUprzywilejowany.TabIndex = 235
        Me.rbtPracownikUprzywilejowany.Text = "PRACOWNIK UPRZYWILEJOWANY - posiada dostêp do funkcji u¿ytkowych, archiwizowania " &
    "i pobierania danych zTOFI . . ." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.rbtPracownikUprzywilejowany.UseVisualStyleBackColor = True
        '
        'rbtPracownik
        '
        Me.rbtPracownik.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.rbtPracownik.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtPracownik.ForeColor = System.Drawing.Color.Navy
        Me.rbtPracownik.Location = New System.Drawing.Point(14, 90)
        Me.rbtPracownik.Name = "rbtPracownik"
        Me.rbtPracownik.Size = New System.Drawing.Size(731, 20)
        Me.rbtPracownik.TabIndex = 234
        Me.rbtPracownik.Text = "PRACOWNIK - posiada dostêp tylko do funkcji u¿ytkowych systemu (menu KLIENCI) . ." &
    " ."
        Me.rbtPracownik.UseVisualStyleBackColor = True
        '
        'rbtSzef
        '
        Me.rbtSzef.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.rbtSzef.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtSzef.ForeColor = System.Drawing.Color.Navy
        Me.rbtSzef.Location = New System.Drawing.Point(14, 50)
        Me.rbtSzef.Name = "rbtSzef"
        Me.rbtSzef.Size = New System.Drawing.Size(731, 20)
        Me.rbtSzef.TabIndex = 232
        Me.rbtSzef.Text = "KIEROWNIK - posiada uprawnienia umo¿liwiaj¹ce dostêp do wszystkich funkcji system" &
    "u oprócz edycji uprawnieñ . . ."
        Me.rbtSzef.UseVisualStyleBackColor = True
        '
        'rbtAdministrator
        '
        Me.rbtAdministrator.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.rbtAdministrator.Checked = True
        Me.rbtAdministrator.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtAdministrator.ForeColor = System.Drawing.Color.Navy
        Me.rbtAdministrator.Location = New System.Drawing.Point(14, 30)
        Me.rbtAdministrator.Name = "rbtAdministrator"
        Me.rbtAdministrator.Size = New System.Drawing.Size(731, 20)
        Me.rbtAdministrator.TabIndex = 231
        Me.rbtAdministrator.TabStop = True
        Me.rbtAdministrator.Text = "ADMINISTRATOR - posiada najszersze uprawnienia w systemie, mo¿e dopisywaæ i edyto" &
    "waæ uprawnienia u¿ytkowników . . ."
        Me.rbtAdministrator.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(14, 13)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(632, 14)
        Me.Label7.TabIndex = 230
        Me.Label7.Text = "Proszê okreœliæ podstawow¹ rolê tego u¿ytkownika w systemie :"
        '
        'frmEdytujKatalogUzytkownikow
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(834, 435)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.cmdWycofajSie)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.picLogo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(8, 8)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEdytujKatalogUzytkownikow"
        Me.Text = "Edycja Informacji o U¿ytkowniku"
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.tabPracownik.ResumeLayout(False)
        Me.tabPracownik.PerformLayout()
        Me.TabUprawnienia.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    '----------------------------------------------------------
    'Katalog Uzytkownicy
    '----------------------------------------------------------
    'Public oIdentUzytkownika As String             'IDENT  Text    20
    'Public oNazwaUzytkownika As String             'NAZWA             Text    100
    'Public oFunkcjaUzytkownika As String           'FUNKCJA           Text    100
    'Public oHasloUzytkownika As String             'HASLO             Text    100
    'Public oUprawnieniaUzytkownika As String       'UPRAWNIENIA       Text    100
    'Public oUwagiUzytkownika As String   'UWAGI          Text
    'Public oWersjaUzytkownika As Integer 'WERSJA         Integer


    Private Sub frmEdytujKatalogUzytkownikow_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.picLogo.Image = My.Resources.logomini
        PobierzDane()
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
            txtIdent.Enabled = False
        Else
            txtIdent.Enabled = True
            txtIdent.Focus()
        End If
    End Sub

    Private Sub frmEdytujKatalogUzytkownikow_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If _MaUprawnieniaAdministratora() Then
            'ADMINISTATOR - mo¿e wszystko
        ElseIf (Trim(oIdentUzytkownika) = Program_IdUzytkownika) Then
            'w³aœciciel konta - to nie jest ADMNINSTATOR ale to jego konto...
            'mo¿e edytowaæ opisy - bez uprawnien
            If TabControl1.TabCount > 1 Then TabControl1.TabPages.RemoveAt(1)
        Else
            'ani ADMIN ani wlasciciel konta - tylko przeglada
            'NieEdytowac(Me.Panel1)
            'NieEdytowac(Me.Panel2)
            'NieEdytowac(Me.Panel3)
            'NieEdytowac(Me.Panel4)
            'NieEdytowac(Me.Panel5)
            'NieEdytowac(Me.Panel6)

            If TabControl1.TabCount > 0 Then NieEdytowac(TabControl1.TabPages(0))
            If TabControl1.TabCount > 1 Then TabControl1.TabPages.RemoveAt(1)
            '----------------------------
            'nie wyswietlaj....
            'txtHaslo.Visible = False
            'txtUwagi.Visible = False
            'cmdPokazHaslo.Visible = False
            'lblHaslo.Visible = False
            'lblUwagi.Visible = False
        End If
    End Sub

    Private Sub frmEdytujKatalogUzytkownikow_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub





    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        If Len(txtIdent.Text) = 0 Then
            MessageBox.Show("Proszê wpisaæ identyfikator u¿ytkownika", _
                "Uwaga:", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            If ZapiszDane() Then
                oAktualizuj = True
                Me.Close()
            End If
        End If
    End Sub


    Private Sub cmdWycofajSie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWycofajSie.Click
        oAktualizuj = False
        Me.Close()
    End Sub


    '***************************************************
    '* procedury pobierania/zapisywania danych
    '***************************************************

    Private Sub PobierzDane()
        txtIdent.Text = oIdentUzytkownika
        txtHaslo.Text = oHasloUzytkownika
        txtHaslo2.Text = oHasloUzytkownika
        txtNazwa.Text = oNazwaUzytkownika
        txtFunkcja.Text = oFunkcjaUzytkownika
        txtOpis.Text = oUwagiUzytkownika
        '-----------------------------
        If ((Trim(oIdentUzytkownika) = Program_IdUzytkownika) Or _MaUprawnieniaAdministratora()) Then
        Else
            lblHaslo.Visible = False
            txtHaslo.Visible = False
            lblHaslo2.Visible = False
            txtHaslo2.Visible = False
            cmdPokazHaslo.Visible = False
        End If
        '----------------------
        If Len(Trim(oUprawnieniaUzytkownika)) = 0 Then
            rbtPracownik.Checked = True
        ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaPracownik Then
            rbtPracownik.Checked = True
        ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaPracownikUprzywilejowany Then
            rbtPracownikUprzywilejowany.Checked = True
        ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaSzef Then
            rbtSzef.Checked = True
        ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaAdministrator Then
            rbtAdministrator.Checked = True
        Else
            rbtPracownik.Checked = True
        End If

    End Sub

    Private Function ZapiszDane() As Boolean
        If txtHaslo.Text = txtHaslo2.Text Then
            oIdentUzytkownika = txtIdent.Text
            oHasloUzytkownika = txtHaslo.Text
            oNazwaUzytkownika = txtNazwa.Text
            oFunkcjaUzytkownika = txtFunkcja.Text
            oUwagiUzytkownika = txtOpis.Text

            If rbtPracownik.Checked Then
                oUprawnieniaUzytkownika = ZapiszUprawnieniaUzytkownika(Program_RolaPracownik)
            ElseIf rbtPracownikUprzywilejowany.Checked Then
                oUprawnieniaUzytkownika = ZapiszUprawnieniaUzytkownika(Program_RolaPracownikUprzywilejowany)
            ElseIf rbtSzef.Checked Then
                oUprawnieniaUzytkownika = ZapiszUprawnieniaUzytkownika(Program_RolaSzef)
            Else    'If rbtAdministrator.Checked Then
                oUprawnieniaUzytkownika = ZapiszUprawnieniaUzytkownika(Program_RolaAdministrator)
            End If

            Return (True)
        Else
            MessageBox.Show("Has³o i powtórzone has³o s¹ ró¿ne..." & vbCrLf & _
                            "Proszê poprawnie wpisaæ oba has³a.", _
                "Uwaga - Nie zapisujê danych do katalogu.", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
            Return (False)
        End If
    End Function




    '**************************************************************
    '** Identyfikacja u¿ytkownika
    '**************************************************************

    Private Sub txtIdent_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIdent.TextChanged
        TestLen(txtIdent, "IDENTYFIKATOR", 10)
    End Sub
    Private Sub txtIdent_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdent.GotFocus
        StartEdycjiTxt(txtIdent)
    End Sub
    Private Sub txtIdent_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdent.LostFocus
        KoniecEdycjiTxt(txtIdent)
    End Sub
    '-------------------------------------------
    Private Sub txtHaslo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHaslo.TextChanged
        TestLen(txtHaslo, "HASLO", 100)
    End Sub
    Private Sub txtHaslo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHaslo.GotFocus
        StartEdycjiTxt(txtHaslo)
    End Sub
    Private Sub txtHaslo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHaslo.LostFocus
        KoniecEdycjiTxt(txtHaslo)
    End Sub
    '-------------------------------------------
    Private Sub txtFunkcja_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFunkcja.TextChanged
        TestLen(txtFunkcja, "STANOWISKO", 100)
    End Sub
    Private Sub txtFunkcja_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFunkcja.GotFocus
        StartEdycjiTxt(txtFunkcja)
    End Sub
    Private Sub txtFunkcja_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFunkcja.LostFocus
        KoniecEdycjiTxt(txtFunkcja)
    End Sub
    '-------------------------------------------
    Private Sub txtNazwa_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNazwa.TextChanged
        TestLen(txtNazwa, "NAZWISKO", 100)
    End Sub
    Private Sub txtNazwa_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNazwa.GotFocus
        StartEdycjiTxt(txtNazwa)
    End Sub
    Private Sub txtNazwa_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNazwa.LostFocus
        KoniecEdycjiTxt(txtNazwa)
    End Sub
    '-------------------------------------------
    Private Sub txtOpis_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOpis.TextChanged
        'TestLen(txtOpis, "STANOWISKO", 100)
    End Sub
    Private Sub txtOpis_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpis.GotFocus
        StartEdycjiTxt(txtOpis)
    End Sub
    Private Sub txtOpis_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpis.LostFocus
        KoniecEdycjiTxt(txtOpis)
    End Sub
    '-------------------------------------------


    Private Sub cmdPokazHaslo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPokazHaslo.Click
        If txtHaslo.PasswordChar = "*" Then
            txtHaslo.PasswordChar = ""
            txtHaslo2.PasswordChar = ""
            cmdPokazHaslo.Text = "Ukryj has³o"
        Else
            txtHaslo.PasswordChar = "*"
            txtHaslo2.PasswordChar = "*"
            cmdPokazHaslo.Text = "Poka¿ has³o"
        End If
    End Sub






    '-------------------------------
    ' ograniczenia dostêpu
    '-----------------------------------------

End Class
