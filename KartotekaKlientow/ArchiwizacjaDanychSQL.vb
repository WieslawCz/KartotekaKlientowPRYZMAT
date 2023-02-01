Imports System.IO

Public Class ArchiwizacjaDanychSQL
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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblKatalogArchiwum As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblInstalacja As System.Windows.Forms.Label
    Friend WithEvents dlgOpenFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents cmdArchwizuj As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents cmdWybierz As System.Windows.Forms.Button
    Friend WithEvents cmdWybierzZIP As System.Windows.Forms.Button
    Friend WithEvents lblKatalogZIP As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblPlikDB As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ArchiwizacjaDanychSQL))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdWybierzZIP = New System.Windows.Forms.Button()
        Me.lblKatalogZIP = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblInstalacja = New System.Windows.Forms.Label()
        Me.cmdWybierz = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblKatalogArchiwum = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblPlikDB = New System.Windows.Forms.Label()
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
        Me.Panel1.Controls.Add(Me.cmdWybierzZIP)
        Me.Panel1.Controls.Add(Me.lblKatalogZIP)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.lblInstalacja)
        Me.Panel1.Controls.Add(Me.cmdWybierz)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.lblKatalogArchiwum)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.lblPlikDB)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(528, 112)
        Me.Panel1.TabIndex = 0
        '
        'cmdWybierzZIP
        '
        Me.cmdWybierzZIP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierzZIP.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierzZIP.Image = CType(resources.GetObject("cmdWybierzZIP.Image"), System.Drawing.Image)
        Me.cmdWybierzZIP.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWybierzZIP.Location = New System.Drawing.Point(470, 74)
        Me.cmdWybierzZIP.Name = "cmdWybierzZIP"
        Me.cmdWybierzZIP.Size = New System.Drawing.Size(42, 24)
        Me.cmdWybierzZIP.TabIndex = 57
        Me.cmdWybierzZIP.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblKatalogZIP
        '
        Me.lblKatalogZIP.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKatalogZIP.ForeColor = System.Drawing.Color.Purple
        Me.lblKatalogZIP.Location = New System.Drawing.Point(109, 78)
        Me.lblKatalogZIP.Name = "lblKatalogZIP"
        Me.lblKatalogZIP.Size = New System.Drawing.Size(403, 16)
        Me.lblKatalogZIP.TabIndex = 56
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(8, 79)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(136, 16)
        Me.Label5.TabIndex = 55
        Me.Label5.Text = "Katalog ZIP . . . . . . . . ."
        '
        'lblInstalacja
        '
        Me.lblInstalacja.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblInstalacja.ForeColor = System.Drawing.Color.Purple
        Me.lblInstalacja.Location = New System.Drawing.Point(245, 8)
        Me.lblInstalacja.Name = "lblInstalacja"
        Me.lblInstalacja.Size = New System.Drawing.Size(267, 16)
        Me.lblInstalacja.TabIndex = 52
        '
        'cmdWybierz
        '
        Me.cmdWybierz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierz.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierz.Image = CType(resources.GetObject("cmdWybierz.Image"), System.Drawing.Image)
        Me.cmdWybierz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWybierz.Location = New System.Drawing.Point(470, 51)
        Me.cmdWybierz.Name = "cmdWybierz"
        Me.cmdWybierz.Size = New System.Drawing.Size(42, 24)
        Me.cmdWybierz.TabIndex = 54
        Me.cmdWybierz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(288, 16)
        Me.Label3.TabIndex = 51
        Me.Label3.Text = "Instalacja . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . " &
    ". ."
        '
        'lblKatalogArchiwum
        '
        Me.lblKatalogArchiwum.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKatalogArchiwum.ForeColor = System.Drawing.Color.Purple
        Me.lblKatalogArchiwum.Location = New System.Drawing.Point(109, 55)
        Me.lblKatalogArchiwum.Name = "lblKatalogArchiwum"
        Me.lblKatalogArchiwum.Size = New System.Drawing.Size(403, 16)
        Me.lblKatalogArchiwum.TabIndex = 50
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 16)
        Me.Label2.TabIndex = 48
        Me.Label2.Text = "Katalog archiwum . . . . . . . . ."
        '
        'lblPlikDB
        '
        Me.lblPlikDB.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPlikDB.ForeColor = System.Drawing.Color.Purple
        Me.lblPlikDB.Location = New System.Drawing.Point(245, 32)
        Me.lblPlikDB.Name = "lblPlikDB"
        Me.lblPlikDB.Size = New System.Drawing.Size(267, 16)
        Me.lblPlikDB.TabIndex = 45
        '
        'Label18
        '
        Me.Label18.ForeColor = System.Drawing.Color.Navy
        Me.Label18.Location = New System.Drawing.Point(8, 32)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(344, 16)
        Me.Label18.TabIndex = 44
        Me.Label18.Text = "Nazwa pliku z kopi¹ archiwaln¹ bazy danych . . . . . . . . ."
        '
        'lblStatus
        '
        Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStatus.ForeColor = System.Drawing.Color.Purple
        Me.lblStatus.Location = New System.Drawing.Point(8, 130)
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
        Me.cmdArchwizuj.Location = New System.Drawing.Point(349, 156)
        Me.cmdArchwizuj.Name = "cmdArchwizuj"
        Me.cmdArchwizuj.Size = New System.Drawing.Size(90, 22)
        Me.cmdArchwizuj.TabIndex = 56
        Me.cmdArchwizuj.Text = "Archiwizuj"
        Me.cmdArchwizuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 156)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 23)
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
        Me.cmdPowrot.Location = New System.Drawing.Point(447, 156)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(87, 22)
        Me.cmdPowrot.TabIndex = 54
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ArchiwizacjaDanychSQL
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(543, 184)
        Me.Controls.Add(Me.cmdArchwizuj)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.Panel1)
        Me.ForeColor = System.Drawing.Color.Purple
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ArchiwizacjaDanychSQL"
        Me.ShowInTaskbar = False
        Me.Text = "Archiwizacja danych (Backup)"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdArchwizuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdArchwizuj.Click
        Dim oDstFileDB As String = ""
        Dim oDstFileLOG As String = ""
        Dim oDstFileZip As String = ""

        Dim sConnectionStr As String = ConnectionStrings()
        Dim dbConnection As New SqlClient.SqlConnection(sConnectionStr)
        'Dim dbConnection As New OleDb.OleDbConnection(sConnectionStr)
        Dim dbCommand As SqlClient.SqlCommand = New SqlClient.SqlCommand
        'Dim dbCommand As OleDb.OleDbCommand = New OleDb.OleDbCommand

        Dim sql As String
        Dim Wynik As Integer

        Dim APath As String = Trim(lblKatalogArchiwum.Text)
        If Mid(APath, Len(APath), 1) = "\" Then
            oDstFileDB = Trim(lblKatalogArchiwum.Text) & Trim(lblPlikDB.Text)
            'oDstFileLOG = Trim(lblKatalogArchiwum.Text) & Trim(lblPlikLOG.Text)
        Else
            oDstFileDB = Trim(lblKatalogArchiwum.Text) & "\" & Trim(lblPlikDB.Text)
            'oDstFileLOG = Trim(lblKatalogArchiwum.Text) & "\" & Trim(lblPlikLOG.Text)
        End If

        Dim NazwaPliku As String = Trim(ParL_NazwaArchiwum) & "_" & _
                  Microsoft.VisualBasic.Right("00" & Trim(Now.Year.ToString()), 2) & _
                  Microsoft.VisualBasic.Right("00" & Trim(Now.Month.ToString()), 2) & _
                  Microsoft.VisualBasic.Right("00" & Trim(Now.Day.ToString()), 2) & "_" & _
                  Microsoft.VisualBasic.Right("00" & Trim(Now.Hour.ToString()), 2) & _
                  Microsoft.VisualBasic.Right("00" & Trim(Now.Minute.ToString()), 2) & _
                  Microsoft.VisualBasic.Right("00" & Trim(Now.Second.ToString()), 2) & ".zip"

        APath = Trim(lblKatalogZIP.Text)
        If Mid(APath, Len(APath), 1) = "\" Then
            oDstFileZip = Trim(lblKatalogZIP.Text) & NazwaPliku
        Else
            oDstFileZip = Trim(lblKatalogZIP.Text) & "\" & NazwaPliku
        End If
        If File.Exists(oDstFileZip) Then File.Delete(oDstFileZip)
        Dim ListaPlikowDoZipowania As String = ""
        '--------------------------------------
        lblStatus.Text = "Archiwizuje baze danych..."
        System.Windows.Forms.Application.DoEvents()

        'BACKUP DATABASE
        'INIT           - nadpisz poprzedni Backup
        'NOUNLOAD       - nie trzeba rozladowywaæ taœmy...
        'NAME=          - nazwa tego BackUpuj
        'NOSKIP         - sprawdz waznosc (expirations) poprzednich backupow przed nadpisaniem
        'STATS=10       - wyswietl statystyki co 10 %
        'NOFORMAT       - Nie formatuj przed backupem
        sql = "BACKUP DATABASE " & ParL_DataBase & " TO DISK=N'" & Trim(oDstFileDB) & _
              "' WITH  INIT ,  NOUNLOAD ,  NAME = N'KartotekaKlientow database backup',  NOSKIP ,  STATS = 10,  NOFORMAT "
        dbCommand.Connection = dbConnection
        dbCommand.CommandType = CommandType.Text
        dbCommand.CommandText = sql
        dbConnection.Open()
        Try
            Wynik = dbCommand.ExecuteNonQuery()
            If Wynik = -1 Then
            Else
                MessageBox.Show("Nie uda³o sie wykonaæ poprawnie kopii archiwalnej bazy danych...")
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        ListaPlikowDoZipowania &= oDstFileDB & "|"
        '--------------------------------------
        ''lblStatus.Text = "Archiwizuje log transakcji..."
        ''System.Windows.Forms.Application.DoEvents()

        ' ''BACKUP DATABASE
        ' ''INIT           - nadpisz poprzedni Backup
        ' ''NOUNLOAD       - nie trzeba rozladowywaæ taœmy...
        ' ''NAME=          - nazwa tego BackUpuj
        ' ''NOSKIP         - sprawdz waznosc (expirations) poprzednich backupow przed nadpisaniem
        ' ''STATS=10       - wyswietl statystyki co 10 %
        ' ''NOFORMAT       - Nie formatuj przed backupem
        ''sql = "BACKUP LOG " & ParL_DataBase & " TO DISK=N'" & Trim(oDstFileLOG) & _
        ''      "' WITH  NOINIT ,  NOUNLOAD ,  NAME = N'KartotekaKlientow Log backup',  NOSKIP ,  STATS = 10,  NOFORMAT "
        ''dbCommand.Connection = dbConnection
        ''dbCommand.CommandType = CommandType.Text
        ''dbCommand.CommandText = sql
        ' ''dbConnection.Open()
        ''Try
        ''    Wynik = dbCommand.ExecuteNonQuery()
        ''    dbConnection.Close()
        ''    If Wynik = -1 Then
        ''    Else
        ''        MessageBox.Show("Nie uda³o sie wykonaæ poprawnie kopii archiwalnej logu transakcji...")
        ''        Exit Sub
        ''    End If
        ''Catch ex As Exception
        ''    MessageBox.Show(ex.Message)
        ''End Try
        ''ListaPlikowDoZipowania &= oDstFileLog & "|"
        '--------------------------------------
        lblStatus.Text = "Wykonujê kompresje archiwum..."
        System.Windows.Forms.Application.DoEvents()

        ZipFile(ListaPlikowDoZipowania, Trim(oDstFileZip))
        '--------------------------------------
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

    Private Sub ArchiwizacjaDanychSQL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        lblInstalacja.Text = ParL_OpisInstalacji
        lblPlikDB.Text = Trim(ParL_NazwaArchiwum) & "_Database.BAK"
        'lblPlikLOG.Text = Trim(ParL_NazwaArchiwum) & "_Log.BAK"
        lblKatalogArchiwum.Text = Trim(ParL_KatalogArchiwum)
        lblKatalogZIP.Text = Trim(ParL_KatalogArchiwumZIP)
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub ArchiwizacjaDanychSQL_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdWybierzZIP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierzZIP.Click
        With dlgOpenFolder
            .Description = "Proszê wybraæ katalog dyskowy do którego wpisaæ plik archiwum..."
            '.RootFolder = lblKatalogArchiwum.Text
            .ShowNewFolderButton = True

            'Me.Visible = False
            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                lblKatalogZIP.Text = .SelectedPath
            End If
        End With

    End Sub
End Class
