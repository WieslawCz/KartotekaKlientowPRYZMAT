Public Class ArchiwizacjaDanych
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
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents lblKatalogMagazyn As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblPlikArchiwum As System.Windows.Forms.Label
    Friend WithEvents lblKatalogArchiwum As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblInstalacja As System.Windows.Forms.Label
    Friend WithEvents dlgOpenFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents txtSlad As System.Windows.Forms.TextBox
    Friend WithEvents cmdArchwizuj As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents cmdWybierz As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ArchiwizacjaDanych))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdWybierz = New System.Windows.Forms.Button()
        Me.txtSlad = New System.Windows.Forms.TextBox()
        Me.lblInstalacja = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblKatalogArchiwum = New System.Windows.Forms.Label()
        Me.lblPlikArchiwum = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblKatalogMagazyn = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.dlgOpenFolder = New System.Windows.Forms.FolderBrowserDialog()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.cmdArchwizuj = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cmdWybierz)
        Me.Panel1.Controls.Add(Me.txtSlad)
        Me.Panel1.Controls.Add(Me.lblInstalacja)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.lblKatalogArchiwum)
        Me.Panel1.Controls.Add(Me.lblPlikArchiwum)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.lblKatalogMagazyn)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(528, 136)
        Me.Panel1.TabIndex = 0
        '
        'cmdWybierz
        '
        Me.cmdWybierz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierz.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierz.Image = CType(resources.GetObject("cmdWybierz.Image"), System.Drawing.Image)
        Me.cmdWybierz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWybierz.Location = New System.Drawing.Point(392, 104)
        Me.cmdWybierz.Name = "cmdWybierz"
        Me.cmdWybierz.Size = New System.Drawing.Size(120, 24)
        Me.cmdWybierz.TabIndex = 54
        Me.cmdWybierz.Text = "Wybierz katalog"
        Me.cmdWybierz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtSlad
        '
        Me.txtSlad.BackColor = System.Drawing.SystemColors.Control
        Me.txtSlad.ForeColor = System.Drawing.Color.Purple
        Me.txtSlad.Location = New System.Drawing.Point(8, 104)
        Me.txtSlad.Multiline = True
        Me.txtSlad.Name = "txtSlad"
        Me.txtSlad.Size = New System.Drawing.Size(160, 24)
        Me.txtSlad.TabIndex = 53
        Me.txtSlad.Text = "TextBox1"
        Me.txtSlad.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtSlad.Visible = False
        '
        'lblInstalacja
        '
        Me.lblInstalacja.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblInstalacja.ForeColor = System.Drawing.Color.Purple
        Me.lblInstalacja.Location = New System.Drawing.Point(128, 8)
        Me.lblInstalacja.Name = "lblInstalacja"
        Me.lblInstalacja.Size = New System.Drawing.Size(384, 16)
        Me.lblInstalacja.TabIndex = 52
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(136, 16)
        Me.Label3.TabIndex = 51
        Me.Label3.Text = "Instalacja . . . . . . . . . . . . . . "
        '
        'lblKatalogArchiwum
        '
        Me.lblKatalogArchiwum.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKatalogArchiwum.ForeColor = System.Drawing.Color.Purple
        Me.lblKatalogArchiwum.Location = New System.Drawing.Point(128, 80)
        Me.lblKatalogArchiwum.Name = "lblKatalogArchiwum"
        Me.lblKatalogArchiwum.Size = New System.Drawing.Size(384, 16)
        Me.lblKatalogArchiwum.TabIndex = 50
        '
        'lblPlikArchiwum
        '
        Me.lblPlikArchiwum.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPlikArchiwum.ForeColor = System.Drawing.Color.Purple
        Me.lblPlikArchiwum.Location = New System.Drawing.Point(128, 56)
        Me.lblPlikArchiwum.Name = "lblPlikArchiwum"
        Me.lblPlikArchiwum.Size = New System.Drawing.Size(384, 16)
        Me.lblPlikArchiwum.TabIndex = 49
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 16)
        Me.Label2.TabIndex = 48
        Me.Label2.Text = "Katalog archiwum . . . . . . . . ."
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 16)
        Me.Label1.TabIndex = 47
        Me.Label1.Text = "Nazwa pliku archiwum . . . "
        '
        'lblKatalogMagazyn
        '
        Me.lblKatalogMagazyn.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKatalogMagazyn.ForeColor = System.Drawing.Color.Purple
        Me.lblKatalogMagazyn.Location = New System.Drawing.Point(128, 32)
        Me.lblKatalogMagazyn.Name = "lblKatalogMagazyn"
        Me.lblKatalogMagazyn.Size = New System.Drawing.Size(384, 16)
        Me.lblKatalogMagazyn.TabIndex = 45
        '
        'Label18
        '
        Me.Label18.ForeColor = System.Drawing.Color.Navy
        Me.Label18.Location = New System.Drawing.Point(8, 32)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(136, 16)
        Me.Label18.TabIndex = 44
        Me.Label18.Text = "Baza danych . . . . . . . . ."
        '
        'lblStatus
        '
        Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStatus.ForeColor = System.Drawing.Color.Purple
        Me.lblStatus.Location = New System.Drawing.Point(8, 152)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(528, 16)
        Me.lblStatus.TabIndex = 53
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdArchwizuj
        '
        Me.cmdArchwizuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdArchwizuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdArchwizuj.ForeColor = System.Drawing.Color.Black
        Me.cmdArchwizuj.Image = CType(resources.GetObject("cmdArchwizuj.Image"), System.Drawing.Image)
        Me.cmdArchwizuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdArchwizuj.Location = New System.Drawing.Point(338, 178)
        Me.cmdArchwizuj.Name = "cmdArchwizuj"
        Me.cmdArchwizuj.Size = New System.Drawing.Size(90, 23)
        Me.cmdArchwizuj.TabIndex = 56
        Me.cmdArchwizuj.Text = "Archiwizuj"
        Me.cmdArchwizuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 178)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 55
        Me.PictureBox2.TabStop = False
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPowrot.ForeColor = System.Drawing.Color.Black
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(437, 178)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(88, 23)
        Me.cmdPowrot.TabIndex = 54
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ArchiwizacjaDanych
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(541, 206)
        Me.Controls.Add(Me.cmdArchwizuj)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.Panel1)
        Me.ForeColor = System.Drawing.Color.Purple
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ArchiwizacjaDanych"
        Me.ShowInTaskbar = False
        Me.Text = "Archiwizacja danych"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdArchwizuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdArchwizuj.Click
        Dim oSrcFile As String
        Dim oDstFile As String
        Dim oZipFile As String
        Dim oKatalogAplikacji As String
        Dim Slad As String = ""

        lblStatus.Text = "Archiwizuje dane programu Kartoteka Klientów..."

        txtSlad.Left = Panel1.Left
        txtSlad.Top = Panel1.Top
        txtSlad.Width = Panel1.Width
        txtSlad.Height = Panel1.Height
        txtSlad.Visible = True

        'oKatalogAplikacji = System.Windows.Forms.Application.CommonAppDataPath
        oKatalogAplikacji = Trim(System.Windows.Forms.Application.StartupPath)   'katalog pliku EXE

        'tutaj zostanie przepisany skompresowany  ZIP plik z Baza Danych
        If Mid(Trim(oKatalogAplikacji), Len(Trim(oKatalogAplikacji)), 1) = "\" Then
            oDstFile = Trim(oKatalogAplikacji) & Trim(lblPlikArchiwum.Text)
        Else
            oDstFile = Trim(oKatalogAplikacji) & "\" & Trim(lblPlikArchiwum.Text)
        End If

        'tutaj zostanie przepisany plik ZIP po zakonczeniu kompresji...
        If Mid(Trim(lblKatalogArchiwum.Text), Len(Trim(lblKatalogArchiwum.Text)), 1) = "\" Then
            oZipFile = Trim(lblKatalogArchiwum.Text) & Trim(lblPlikArchiwum.Text)
        Else
            oZipFile = Trim(lblKatalogArchiwum.Text) & "\" & Trim(lblPlikArchiwum.Text)
        End If

        Slad = "Tworze archiwum w " & oDstFile & vbCrLf
        txtSlad.Text = Slad

        'USUN POPRZEDNIE ARCHIWUM...
        If (System.IO.File.Exists(oDstFile)) Then Kill(oDstFile)
        Dim ListaPlikowDoZipowania As String = ""

        'archiwizujemy wszystkie pliki *MDB z katalogu bazy danych...
        oSrcFile = Dir(Trim(ParL_KatZbiorow) & "\" & "*.MDB")
        Do While Len(oSrcFile) > 0
            Slad &= "Archiwizuje plik " & oSrcFile & vbCrLf
            txtSlad.Text = Slad
            System.Windows.Forms.Application.DoEvents()
            '-----------------------------------------------
            ListaPlikowDoZipowania &= Trim(ParL_KatZbiorow) & "\" & oSrcFile & "|"
            'ZipFile(Trim(ParL_KatZbiorow) & "\" & oSrcFile, oDstFile)
            '-----------------------------------------------
            oSrcFile = Dir()
        Loop

        'dodatkowo kompresujemy wszystkie pliki CFG z katalogu aplikacji...
        oSrcFile = Dir(Trim(oKatalogAplikacji) & "\" & "*.CFG")
        Do While Len(oSrcFile) > 0
            Slad &= "Archiwizuje plik " & oSrcFile & vbCrLf
            txtSlad.Text = Slad
            System.Windows.Forms.Application.DoEvents()
            '-----------------------------------------------
            ListaPlikowDoZipowania &= Trim(oKatalogAplikacji) & "\" & oSrcFile & "|"
            'ZipFile(Trim(oKatalogAplikacji) & "\" & oSrcFile, oDstFile)
            '-----------------------------------------------
            oSrcFile = Dir()
        Loop

        ZipFile(ListaPlikowDoZipowania, oDstFile)

        Slad &= "Kopiuje archiwum do " & oZipFile & vbCrLf
        txtSlad.Text = Slad
        FileCopy(oDstFile, oZipFile)
        Slad &= "Skoñczy³em." & vbCrLf
        txtSlad.Text = Slad

        txtSlad.Visible = False
        lblStatus.Text = "OK, skoñczy³em archiwizowanie danych..."
    End Sub

    Private Sub cmdWybierz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierz.Click
        With dlgOpenFolder
            .Description = "Proszê wybraæ katalog dyskowy do którego wpisaæ plik archiwum..."
            '.RootFolder = lblKatalogArchiwum.Text
            .ShowNewFolderButton = True

            'Me.Visible = False
            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                lblKatalogArchiwum.Text = .SelectedPath
            End If
        End With
    End Sub

    Private Sub ArchiwizacjaDanych_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        lblInstalacja.Text = ParL_OpisInstalacji
        lblKatalogMagazyn.Text = ParL_KatZbiorow
        lblPlikArchiwum.Text = ParL_NazwaArchiwum & ".Zip"
        lblKatalogArchiwum.Text = ParL_KatalogArchiwum
        lblKatalogArchiwum.Text = ParL_KatalogArchiwumZIP
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub ArchiwizacjaDanych_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub
End Class
