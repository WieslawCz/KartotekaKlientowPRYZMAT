Public Class frmIdentyfikacjaUzytkownika
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Friend WithEvents lblUprawnienia As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmdPokazHaslo As System.Windows.Forms.Button
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    'Fields
    'Constructors
    'Events
    'Methods



    '====================================================
    'zmienne do œledzenia JAK ZAMKNIETO FORME
    Private _closeClick As Boolean
    'Private components As System.ComponentModel.Container
    Public Const SC_CLOSE As Integer = 61536
    Public Const WM_SYSCOMMAND As Integer = 274
    '====================================================

    Public Sub New()
        MyBase.New()
        'inicjowanie zmiennej do wskazania sposobu zamkniecia formy
        _closeClick = False
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call

    End Sub

    '====================================================
    ' metoda WndProc wykorzystywana do oznaczenia sposobu zamkniecia formy
    Protected Overloads Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = WM_SYSCOMMAND AndAlso m.WParam.ToInt32 = SC_CLOSE Then
            'zamkniêto naciskaj¹c krzyzyk w prawym gornym rogu...
            Me._closeClick = True
        End If
        MyBase.WndProc(m)
    End Sub
    '====================================================

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
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtLogin As System.Windows.Forms.TextBox
    Friend WithEvents txtHaslo As System.Windows.Forms.TextBox
    Friend WithEvents lblData As System.Windows.Forms.Label
    Friend WithEvents lblKomputer As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdSprawdz As System.Windows.Forms.Button
    Friend WithEvents cmdZakoncz As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmIdentyfikacjaUzytkownika))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdPokazHaslo = New System.Windows.Forms.Button()
        Me.lblUprawnienia = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmdSprawdz = New System.Windows.Forms.Button()
        Me.lblKomputer = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblData = New System.Windows.Forms.Label()
        Me.txtHaslo = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtLogin = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdZakoncz = New System.Windows.Forms.Button()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cmdPokazHaslo)
        Me.Panel1.Controls.Add(Me.lblUprawnienia)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.cmdSprawdz)
        Me.Panel1.Controls.Add(Me.lblKomputer)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.lblData)
        Me.Panel1.Controls.Add(Me.txtHaslo)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.txtLogin)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(360, 226)
        Me.Panel1.TabIndex = 0
        '
        'cmdPokazHaslo
        '
        Me.cmdPokazHaslo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPokazHaslo.Location = New System.Drawing.Point(264, 78)
        Me.cmdPokazHaslo.Name = "cmdPokazHaslo"
        Me.cmdPokazHaslo.Size = New System.Drawing.Size(80, 23)
        Me.cmdPokazHaslo.TabIndex = 207
        Me.cmdPokazHaslo.Text = "Poka¿ has³o"
        '
        'lblUprawnienia
        '
        Me.lblUprawnienia.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblUprawnienia.ForeColor = System.Drawing.Color.Navy
        Me.lblUprawnienia.Location = New System.Drawing.Point(112, 104)
        Me.lblUprawnienia.Name = "lblUprawnienia"
        Me.lblUprawnienia.Size = New System.Drawing.Size(232, 110)
        Me.lblUprawnienia.TabIndex = 206
        Me.lblUprawnienia.Text = "MójKomp"
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(8, 105)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 16)
        Me.Label5.TabIndex = 205
        Me.Label5.Text = "Uprawnienia. . . . . . . . . . . . ."
        '
        'cmdSprawdz
        '
        Me.cmdSprawdz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSprawdz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSprawdz.Location = New System.Drawing.Point(264, 54)
        Me.cmdSprawdz.Name = "cmdSprawdz"
        Me.cmdSprawdz.Size = New System.Drawing.Size(80, 23)
        Me.cmdSprawdz.TabIndex = 201
        Me.cmdSprawdz.Text = "SprawdŸ"
        Me.cmdSprawdz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblKomputer
        '
        Me.lblKomputer.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKomputer.ForeColor = System.Drawing.Color.Navy
        Me.lblKomputer.Location = New System.Drawing.Point(112, 32)
        Me.lblKomputer.Name = "lblKomputer"
        Me.lblKomputer.Size = New System.Drawing.Size(146, 16)
        Me.lblKomputer.TabIndex = 18
        Me.lblKomputer.Text = "MójKomp"
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(8, 32)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 16)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "Nazwa komputera. . . . . . . . . . . . ."
        '
        'lblData
        '
        Me.lblData.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblData.ForeColor = System.Drawing.Color.Navy
        Me.lblData.Location = New System.Drawing.Point(112, 8)
        Me.lblData.Name = "lblData"
        Me.lblData.Size = New System.Drawing.Size(146, 16)
        Me.lblData.TabIndex = 16
        Me.lblData.Text = "2005-11-16"
        '
        'txtHaslo
        '
        Me.txtHaslo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtHaslo.ForeColor = System.Drawing.Color.Purple
        Me.txtHaslo.Location = New System.Drawing.Point(112, 80)
        Me.txtHaslo.Name = "txtHaslo"
        Me.txtHaslo.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtHaslo.Size = New System.Drawing.Size(146, 20)
        Me.txtHaslo.TabIndex = 14
        Me.txtHaslo.Text = "Data"
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 16)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Has³o. . . . . . . . . . . . ."
        '
        'txtLogin
        '
        Me.txtLogin.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtLogin.ForeColor = System.Drawing.Color.Purple
        Me.txtLogin.Location = New System.Drawing.Point(112, 56)
        Me.txtLogin.Name = "txtLogin"
        Me.txtLogin.Size = New System.Drawing.Size(146, 20)
        Me.txtLogin.TabIndex = 12
        Me.txtLogin.Text = "Data"
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 16)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Nazwa u¿ytkownika. . . . . . . . . . . . ."
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Data i godzina . . . . . . . . . . . . ."
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(277, -2)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(56, 56)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 204
        Me.PictureBox1.TabStop = False
        '
        'cmdZakoncz
        '
        Me.cmdZakoncz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdZakoncz.Image = CType(resources.GetObject("cmdZakoncz.Image"), System.Drawing.Image)
        Me.cmdZakoncz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdZakoncz.Location = New System.Drawing.Point(288, 242)
        Me.cmdZakoncz.Name = "cmdZakoncz"
        Me.cmdZakoncz.Size = New System.Drawing.Size(80, 23)
        Me.cmdZakoncz.TabIndex = 202
        Me.cmdZakoncz.Text = "Dalej"
        Me.cmdZakoncz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.Location = New System.Drawing.Point(8, 241)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(88, 24)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 203
        Me.picLogo.TabStop = False
        '
        'frmIdentyfikacjaUzytkownika
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 272)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.cmdZakoncz)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmIdentyfikacjaUzytkownika"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Identyfikacja U¿ytkownika"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    '====================================================
    'UWAGA - aby to zadzialalo trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    '====================================================
    Private Sub IdentyfikacjaUzytkownika_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me._closeClick Then
            e.Cancel = True     'nie pozwalaj zamkn¹c forme
            MessageBox.Show("Proszê zamkn¹æ formê klawiszem DALEJ", _
                "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
            Me._closeClick = False
        Else
        End If
    End Sub


    Private Sub IdentyfikacjaUzytkownika_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.picLogo.Image = My.Resources.logomini
        '--------------------------------
        'Get IP Address
        Dim HostName As String
        Dim IPAddress As String
        HostName = System.Net.Dns.GetHostName()
        IPAddress = System.Net.Dns.GetHostEntry(HostName).AddressList(0).ToString()

        lblData.Text = Now
        lblKomputer.Text = HostName
        txtLogin.Text = ""
        txtHaslo.Text = ""
        lblUprawnienia.Text = "U¿ytkownik niezidentyfikowany." & vbCrLf & "Brak uprawnieñ w programie..."
        Program_STOP = True
    End Sub

    Private Sub frmIdentyfikacjaUzytkownika_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub

    Private Sub txtLogin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLogin.TextChanged
        TestLen(txtLogin, "NAZWA U¯YTKOWNIKA", 50)
    End Sub
    Private Sub txtLogin_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLogin.GotFocus
        StartEdycjiTxt(txtLogin)
    End Sub
    Private Sub txtLogin_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLogin.LostFocus
        KoniecEdycjiTxt(txtLogin)
    End Sub
    Private Sub txtLogin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLogin.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                txtHaslo.Select()
            Case Keys.Insert, Keys.Add
            Case Keys.Delete, Keys.Subtract
            Case Else
        End Select
    End Sub

    Private Sub txtHaslo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHaslo.TextChanged
        TestLen(txtHaslo, "HAS£O", 50)
        '-----------------------------
        'przejdz w tryb ukrytego hasla
        txtHaslo.PasswordChar = "*"
        cmdPokazHaslo.Text = "Poka¿ has³o"
    End Sub
    Private Sub txtHaslo_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHaslo.GotFocus
        StartEdycjiTxt(txtHaslo)
    End Sub
    Private Sub txtHaslo_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHaslo.LostFocus
        KoniecEdycjiTxt(txtHaslo)
        '---------------------
        cmdSprawdz_Click(sender, e)
    End Sub
    Private Sub txtHaslo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtHaslo.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                KoniecEdycjiTxt(txtHaslo)
                cmdZakoncz_Click(Me, e)
            Case Keys.Insert, Keys.Add
            Case Keys.Delete, Keys.Subtract
            Case Else
        End Select
    End Sub

    Private Sub cmdZakoncz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdZakoncz.Click
        'Program_STOP = True
        If SprawdzUprawnieniaUzytkownika() Then
            Me.Close()
        Else
        End If
    End Sub

    Private Sub cmdSprawdz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSprawdz.Click
        SprawdzUprawnieniaUzytkownika()
    End Sub


    Private Function SprawdzUprawnieniaUzytkownika() As Boolean
        Dim aktUser As String = txtLogin.Text
        Dim aktPass As String = txtHaslo.Text
        Dim NieTrzebaUzupelnic As Boolean = True

        'na razie wszystkich przepuszczaj.........................aktUser="" and aktPass=""
        'aktUser = UCase(Program_SupervisorLogin)
        'aktPass = UCase(Program_SupervisorHaslo)

        If SprawdzUzytkownika(aktUser, aktPass) Then
            lblUprawnienia.Text = "U¿ytkownik zidentyfikowany." & vbCrLf
            Program_STOP = False
            NieTrzebaUzupelnic = True

            'Public Const Program_RolaAdministrator As String = "A"
            'Public Const Program_RolaSzef As String = "S"
            'Public Const Program_RolaPracownikUprzywilejowany As String = "U"
            'Public Const Program_RolaPrzedstawicielHandlowy As String = "H"
            'Public Const Program_RolaKsiegowosc As String = "X"
            'Public Const Program_RolaMarketing As String = "C"
            'Public Const Program_RolaZaopatrzenie As String = "Z"
            'Public Const Program_RolaFakturowanie As String = "F"
            'Public Const Program_RolaProdukcja As String = "P"
            'Public Const Program_RolaObslugaMagazynu As String = "M"

            If _MaUprawnieniaAdministratora() Then
                lblUprawnienia.Text &= "Administrator" & vbCrLf
            ElseIf _MaUprawnieniaSzefa() Then
                lblUprawnienia.Text &= "Szef" & vbCrLf
                'ElseIf _MaUprawnieniaPracownikaUprzywilejowanego() Then
                '    lblUprawnienia.Text &= "Pracownik Uprzywilejowany" & vbCrLf
            ElseIf _MaUprawnieniaPracownika() Then
                lblUprawnienia.Text &= "Pracownik" & vbCrLf
            End If


        Else
            If (Len(Trim(aktUser)) = 0) And (Len(Trim(aktPass)) = 0) Then

                lblUprawnienia.Text = "U¿ytkownik niezidentyfikowany." & vbCrLf & "Brak uprawnieñ w programie..."

                If MessageBox.Show("Czy chcesz zakoñczyæ pracê programu ?", _
                "Nie podano nazwy i has³a u¿ytkownika.", _
                System.Windows.Forms.MessageBoxButtons.YesNo, _
                MessageBoxIcon.Question, _
                MessageBoxDefaultButton.Button2) = DialogResult.Yes Then

                    Program_STOP = True
                    NieTrzebaUzupelnic = True
                Else
                    Program_STOP = False
                    NieTrzebaUzupelnic = False
                End If
            Else
                lblUprawnienia.Text = "U¿ytkownik niezidentyfikowany." & vbCrLf & "Brak uprawnieñ w programie..."
                Program_STOP = True
                NieTrzebaUzupelnic = True
            End If
        End If
        Return NieTrzebaUzupelnic
    End Function




    Private Sub cmdPokazHaslo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPokazHaslo.Click
        If txtHaslo.PasswordChar = "*" Then
            txtHaslo.PasswordChar = ""
            cmdPokazHaslo.Text = "Ukryj has³o"
        Else
            txtHaslo.PasswordChar = "*"
            cmdPokazHaslo.Text = "Poka¿ has³o"
        End If
    End Sub




    Public Function SprawdzUzytkownika(ByVal UserLogin As String, ByVal UserPass As String) As Boolean
        Dim Result As Boolean = False
        'najpierw sprawdz w³aœciciela komputera - DEFINICJA W ZMIENNYCH SYSTEMOWYCH
        Dim WlascicielKomputera As String = Environment.GetEnvironmentVariable("WLASCICIEL_KOMPUTERA")


        'najpierw sprawdz czy SuperOperator
        If ((UCase(UserLogin) = UCase(Program_AdministratorLogin)) And (UCase(UserPass) = UCase(Program_AdministratorHaslo))) Or _
           ((UCase(UserLogin) = UCase(Program_AdministratorLogin)) And (UCase(UserPass) = UCase(Program_AdministratorHaslo2))) Or _
           ((UCase(UserLogin) = UCase(Program_AdministratorLogin)) And (UCase(UserPass) = "") And (WlascicielKomputera = "WIESLAW CZAJKOWSKI")) Or
           ((UCase(UserLogin) = "") And (UCase(UserPass) = "") And (WlascicielKomputera = "WIESLAW CZAJKOWSKI")) Then
            'to jest SUPEROPERATOR (SUPERVISOR) - wszystko mo¿e
            Program_IdUzytkownika = Program_AdministratorLogin
            Program_NazwaUzytkownika = Program_AdministratorNazwa
            Program_FunkcjaUzytkownika = "SuperAdministrator programu"
            Program_HasloUzytkownika = UCase(Program_AdministratorHaslo)
            'Program_UprawnieniaUzytkownika = ZapiszUprawnieniaUzytkownika(Program_RolaAdministrator)

            Result = True

        ElseIf (Len(Trim(UserLogin)) = 0) And (Len(Trim(UserPass)) = 0) And PozwalajNaNiekontrolowanyDostep() Then

            Program_IdUzytkownika = Program_AdministratorLogin
            Program_NazwaUzytkownika = Program_AdministratorNazwa
            Program_FunkcjaUzytkownika = "SuperAdministrator programu"
            Program_HasloUzytkownika = UCase(Program_AdministratorHaslo)
            'Program_UprawnieniaUzytkownika = ZapiszUprawnieniaUzytkownika(Program_RolaAdministrator)

            Result = True


        Else
            'teraz sprawdz wsrod uzytkownikow....
            If IdentUzytkownika(UserLogin) Then
                'sprawdz podane haslo

                'Public U_IdentUzytkownika As String           'IDENT  Text    20
                'Public U_NazwaUzytkownika As String           'NAZWA             Text    100
                'Public U_FunkcjaUzytkownika As String         'FUNKCJA           Text    100
                'Public U_HasloUzytkownika As String           'HASLO             Text    100
                'Public U_UprawnieniaUzytkownika As String     'UPRAWNIENIA          Text    100
                'Public U_UwagiUzytkownika As String           'UWAGI          Text
                'Public U_WersjaUzytkownika As Integer         'WERSJA         Integer

                If UserPass = oHasloUzytkownika Then
                    Program_IdUzytkownika = oIdentUzytkownika
                    Program_NazwaUzytkownika = oNazwaUzytkownika
                    Program_FunkcjaUzytkownika = oFunkcjaUzytkownika
                    Program_HasloUzytkownika = oHasloUzytkownika

                    'definicja roli pracownika
                    If Len(Trim(oUprawnieniaUzytkownika)) = 0 Then
                        Program_UprawnieniaUzytkownika = ZapiszUprawnieniaUzytkownika(Program_RolaPracownik)
                    ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaPracownik Then
                        Program_UprawnieniaUzytkownika = ZapiszUprawnieniaUzytkownika(Program_RolaPracownik)
                        'ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaPracownikUprzywilejowany Then
                        '    Program_UprawnieniaUzytkownika = ZapiszUprawnieniaUzytkownika(Program_RolaPracownikUprzywilejowany)
                    ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaSzef Then
                        Program_UprawnieniaUzytkownika = ZapiszUprawnieniaUzytkownika(Program_RolaSzef)
                    ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaAdministrator Then
                        Program_UprawnieniaUzytkownika = ZapiszUprawnieniaUzytkownika(Program_RolaAdministrator)
                    Else
                        Program_UprawnieniaUzytkownika = ZapiszUprawnieniaUzytkownika(Program_RolaPracownik)
                    End If

                    Result = True
                Else
                    Result = False
                End If
            Else
                Result = False
            End If
        End If
        Return (Result)
    End Function



    '-----------------------------------------------------------------
    'DEFINICJA UZYTKOWNIKÓW KTÓRZY MUSZA SIE LOGOWAC INDYWIDUALNIE
    '-----------------------------------------------------------------------------

    Private Function PozwalajNaNiekontrolowanyDostep() As Boolean
        Return False
    End Function

End Class
