Public Class OdtwarzanieDanych
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
    Friend WithEvents txtSlad As System.Windows.Forms.TextBox
    Friend WithEvents lblInstalacja As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblPlikArchiwum As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdWybierz As System.Windows.Forms.Button
    Friend WithEvents lblKatalogMagazyn As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lblKatalogAplikacji As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdPokaz As System.Windows.Forms.Button
    Friend WithEvents cmdOdtwarzaj As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(OdtwarzanieDanych))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdPokaz = New System.Windows.Forms.Button()
        Me.lblKatalogAplikacji = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtSlad = New System.Windows.Forms.TextBox()
        Me.lblInstalacja = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblPlikArchiwum = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdWybierz = New System.Windows.Forms.Button()
        Me.lblKatalogMagazyn = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.cmdOdtwarzaj = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
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
        Me.Panel1.Controls.Add(Me.cmdPokaz)
        Me.Panel1.Controls.Add(Me.lblKatalogAplikacji)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.txtSlad)
        Me.Panel1.Controls.Add(Me.lblInstalacja)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.lblPlikArchiwum)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.cmdWybierz)
        Me.Panel1.Controls.Add(Me.lblKatalogMagazyn)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(527, 145)
        Me.Panel1.TabIndex = 1
        '
        'cmdPokaz
        '
        Me.cmdPokaz.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokaz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPokaz.ForeColor = System.Drawing.Color.Black
        Me.cmdPokaz.Image = CType(resources.GetObject("cmdPokaz.Image"), System.Drawing.Image)
        Me.cmdPokaz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPokaz.Location = New System.Drawing.Point(306, 104)
        Me.cmdPokaz.Name = "cmdPokaz"
        Me.cmdPokaz.Size = New System.Drawing.Size(94, 24)
        Me.cmdPokaz.TabIndex = 56
        Me.cmdPokaz.Text = "Poka¿ .ZIP"
        Me.cmdPokaz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblKatalogAplikacji
        '
        Me.lblKatalogAplikacji.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblKatalogAplikacji.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKatalogAplikacji.ForeColor = System.Drawing.Color.Purple
        Me.lblKatalogAplikacji.Location = New System.Drawing.Point(128, 56)
        Me.lblKatalogAplikacji.Name = "lblKatalogAplikacji"
        Me.lblKatalogAplikacji.Size = New System.Drawing.Size(383, 16)
        Me.lblKatalogAplikacji.TabIndex = 55
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 57)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 16)
        Me.Label4.TabIndex = 54
        Me.Label4.Text = "Katalog programu . . . . . . . . "
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
        Me.lblInstalacja.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblInstalacja.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblInstalacja.ForeColor = System.Drawing.Color.Purple
        Me.lblInstalacja.Location = New System.Drawing.Point(128, 8)
        Me.lblInstalacja.Name = "lblInstalacja"
        Me.lblInstalacja.Size = New System.Drawing.Size(383, 16)
        Me.lblInstalacja.TabIndex = 52
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(136, 16)
        Me.Label3.TabIndex = 51
        Me.Label3.Text = "Instalacja . . . . . . . . . . . . . . "
        '
        'lblPlikArchiwum
        '
        Me.lblPlikArchiwum.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPlikArchiwum.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPlikArchiwum.ForeColor = System.Drawing.Color.Purple
        Me.lblPlikArchiwum.Location = New System.Drawing.Point(128, 80)
        Me.lblPlikArchiwum.Name = "lblPlikArchiwum"
        Me.lblPlikArchiwum.Size = New System.Drawing.Size(383, 16)
        Me.lblPlikArchiwum.TabIndex = 49
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 81)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 16)
        Me.Label1.TabIndex = 47
        Me.Label1.Text = "Nazwa pliku archiwum . . . "
        '
        'cmdWybierz
        '
        Me.cmdWybierz.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierz.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierz.Image = CType(resources.GetObject("cmdWybierz.Image"), System.Drawing.Image)
        Me.cmdWybierz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWybierz.Location = New System.Drawing.Point(405, 104)
        Me.cmdWybierz.Name = "cmdWybierz"
        Me.cmdWybierz.Size = New System.Drawing.Size(105, 24)
        Me.cmdWybierz.TabIndex = 46
        Me.cmdWybierz.Text = "Wybierz plik"
        Me.cmdWybierz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblKatalogMagazyn
        '
        Me.lblKatalogMagazyn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblKatalogMagazyn.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKatalogMagazyn.ForeColor = System.Drawing.Color.Purple
        Me.lblKatalogMagazyn.Location = New System.Drawing.Point(128, 32)
        Me.lblKatalogMagazyn.Name = "lblKatalogMagazyn"
        Me.lblKatalogMagazyn.Size = New System.Drawing.Size(383, 16)
        Me.lblKatalogMagazyn.TabIndex = 45
        '
        'Label18
        '
        Me.Label18.ForeColor = System.Drawing.Color.Navy
        Me.Label18.Location = New System.Drawing.Point(8, 33)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(136, 16)
        Me.Label18.TabIndex = 44
        Me.Label18.Text = "Katalog Bazy Danych . . . . . . . . ."
        '
        'lblStatus
        '
        Me.lblStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStatus.ForeColor = System.Drawing.Color.Purple
        Me.lblStatus.Location = New System.Drawing.Point(8, 161)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(527, 10)
        Me.lblStatus.TabIndex = 57
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdOdtwarzaj
        '
        Me.cmdOdtwarzaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOdtwarzaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdOdtwarzaj.ForeColor = System.Drawing.Color.Black
        Me.cmdOdtwarzaj.Image = CType(resources.GetObject("cmdOdtwarzaj.Image"), System.Drawing.Image)
        Me.cmdOdtwarzaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOdtwarzaj.Location = New System.Drawing.Point(339, 177)
        Me.cmdOdtwarzaj.Name = "cmdOdtwarzaj"
        Me.cmdOdtwarzaj.Size = New System.Drawing.Size(92, 24)
        Me.cmdOdtwarzaj.TabIndex = 56
        Me.cmdOdtwarzaj.Text = "Odtwarzaj"
        Me.cmdOdtwarzaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 177)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 25)
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
        Me.cmdPowrot.Location = New System.Drawing.Point(447, 177)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(88, 24)
        Me.cmdPowrot.TabIndex = 54
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'OdtwarzanieDanych
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(541, 206)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.cmdOdtwarzaj)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "OdtwarzanieDanych"
        Me.ShowInTaskbar = False
        Me.Text = "Odtwarzanie danych z kopii archiwalnej"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdWybierz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierz.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik z archiwum programu..."
            .InitialDirectory = "c:\"
            .DefaultExt = "zip"
            .Filter = "Archiwum programu Kartoteka Klientów (*.zip)|*.zip|Wszystkie pliki|*.*"
            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                lblPlikArchiwum.Text = .FileName
            End If
        End With
    End Sub

    Private Sub OdtwarzanieDanych_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub OdtwarzanieDanych_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        lblInstalacja.Text = ParL_OpisInstalacji
        lblKatalogMagazyn.Text = ParL_KatZbiorow
        lblKatalogAplikacji.Text = Trim(System.Windows.Forms.Application.StartupPath)   'katalog pliku EXE
        lblPlikArchiwum.Text = ""
        txtSlad.Text = ""
        lblStatus.Text = ""
    End Sub

    Private Sub cmdOdtwarzaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOdtwarzaj.Click
        Dim oSrcFile As String
        Dim oDstFile As String
        Dim oZipFile As String
        Dim oKatalogAplikacji As String
        Dim Slad As String = ""

        oZipFile = Trim(lblPlikArchiwum.Text)
        If Len(oZipFile) = 0 Then
            MessageBox.Show("Brak definicji pliku archiwum", _
                "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            If System.IO.File.Exists(oZipFile) Then
                Dim ListaPlikowDoOdZipowania As String = ""

                lblStatus.Text = "Odtwarzam dane z archiwum programu Kartoteka Klientów..."
                Slad = "Odtwarzam dane z archiwum programu Kartoteka Klientów" & vbCrLf
                txtSlad.Text = Slad

                txtSlad.Left = Panel1.Left
                txtSlad.Top = Panel1.Top
                txtSlad.Width = Panel1.Width
                txtSlad.Height = Panel1.Height
                txtSlad.Visible = True

                oKatalogAplikacji = Trim(System.Windows.Forms.Application.StartupPath)   'katalog pliku EXE

                'przed odtworzeniem - zmien nazwe plików aktualnych .MDB na .OLDMDB- aby nie zapisaæ
                oSrcFile = Dir(Trim(ParL_KatZbiorow) & "\" & "*.MDB")
                Do While Len(oSrcFile) > 0
                    oDstFile = Mid(oSrcFile, 1, InStr(UCase(oSrcFile), ".MDB") - 1) & ".oldmdb"
                    Slad &= "Zmieniam nazwe pliku " & oSrcFile & " na " & oDstFile & vbCrLf
                    txtSlad.Text = Slad
                    System.Windows.Forms.Application.DoEvents()
                    '-----------------------------------------------
                    If Len(Dir(Trim(ParL_KatZbiorow) & "\" & oDstFile)) > 0 Then Kill(Trim(ParL_KatZbiorow) & "\" & oDstFile)
                    Rename(Trim(ParL_KatZbiorow) & "\" & oSrcFile, Trim(ParL_KatZbiorow) & "\" & oDstFile)
                    ListaPlikowDoOdZipowania &= Trim(ParL_KatZbiorow) & "\" & oSrcFile & "|"
                    '-----------------------------------------------
                    oSrcFile = Dir(Trim(ParL_KatZbiorow) & "\" & "*.MDB")
                Loop

                'przed odtworzeniem - zmien nazwe plików aktualnych .CFG na .OLDCFG - aby nie zapisaæ
                oSrcFile = Dir(Trim(oKatalogAplikacji) & "\" & "*.CFG")
                Do While Len(oSrcFile) > 0
                    oDstFile = Mid(oSrcFile, 1, InStr(UCase(oSrcFile), ".CFG") - 1) & ".oldcfg"
                    Slad &= "Zmieniam nazwe pliku " & oSrcFile & " na " & oDstFile & vbCrLf
                    txtSlad.Text = Slad
                    System.Windows.Forms.Application.DoEvents()
                    '-----------------------------------------------
                    If Len(Dir(Trim(oKatalogAplikacji) & "\" & oDstFile)) > 0 Then Kill(Trim(oKatalogAplikacji) & "\" & oDstFile)
                    Rename(Trim(oKatalogAplikacji) & "\" & oSrcFile, Trim(oKatalogAplikacji) & "\" & oDstFile)
                    ListaPlikowDoOdZipowania &= Trim(oKatalogAplikacji) & "\" & oSrcFile & "|"
                    '-----------------------------------------------
                    oSrcFile = Dir(Trim(oKatalogAplikacji) & "\" & "*.CFG")
                Loop

                Slad &= "Rozpakowuje archiwum " & Trim(lblPlikArchiwum.Text) & vbCrLf
                txtSlad.Text = Slad

                UnzipFile(ListaPlikowDoOdZipowania, oZipFile)

                Slad &= "Skoñczy³em." & vbCrLf
                txtSlad.Text = Slad

                txtSlad.Visible = False
                lblStatus.Text = "OK, skoñczy³em odtwarzanie danych..."
            Else
                MessageBox.Show("Nie znalaz³em na dysku pliku archiwum" & vbCrLf & oZipFile & vbCrLf, _
                    "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            End If
        End If
    End Sub

    Private Sub cmdPokaz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPokaz.Click
        Dim oZipFile As String

        oZipFile = Trim(lblPlikArchiwum.Text)
        If Len(oZipFile) = 0 Then
            MessageBox.Show("Brak definicji pliku archiwum", _
                "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            If System.IO.File.Exists(oZipFile) Then
                Dim Form As New PokazPlikZIP(oZipFile)
                Form.ShowDialog()
            Else
                MessageBox.Show("Nie znalaz³em na dysku pliku archiwum" & vbCrLf & oZipFile & vbCrLf, _
                    "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            End If
        End If
    End Sub
End Class
