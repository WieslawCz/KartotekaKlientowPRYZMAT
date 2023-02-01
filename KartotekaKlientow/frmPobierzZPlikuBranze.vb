Imports System.IO

Public Class frmPobierzZPlikuBranze
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "

    Private CoPobieramy As String = ""
    Private _Dataset As DataSet = Nothing



    Public Sub New(ByVal parCoPobieramy As String,
                   ByRef parDataSet As DataSet)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        CoPobieramy = parCoPobieramy
        _Dataset = parDataSet
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
    Friend WithEvents dlgOpenFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents cmdPobierz As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdWybierzPlik As Button
    Friend WithEvents txtPlikTXT As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPobierzZPlikuBranze))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdWybierzPlik = New System.Windows.Forms.Button()
        Me.txtPlikTXT = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dlgOpenFolder = New System.Windows.Forms.FolderBrowserDialog()
        Me.cmdPobierz = New System.Windows.Forms.Button()
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
        Me.Panel1.Controls.Add(Me.cmdWybierzPlik)
        Me.Panel1.Controls.Add(Me.txtPlikTXT)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(636, 36)
        Me.Panel1.TabIndex = 0
        '
        'cmdWybierzPlik
        '
        Me.cmdWybierzPlik.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzPlik.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierzPlik.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierzPlik.Image = CType(resources.GetObject("cmdWybierzPlik.Image"), System.Drawing.Image)
        Me.cmdWybierzPlik.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWybierzPlik.Location = New System.Drawing.Point(587, 5)
        Me.cmdWybierzPlik.Name = "cmdWybierzPlik"
        Me.cmdWybierzPlik.Size = New System.Drawing.Size(37, 22)
        Me.cmdWybierzPlik.TabIndex = 77
        Me.cmdWybierzPlik.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPlikTXT
        '
        Me.txtPlikTXT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPlikTXT.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtPlikTXT.Location = New System.Drawing.Point(176, 6)
        Me.txtPlikTXT.Name = "txtPlikTXT"
        Me.txtPlikTXT.Size = New System.Drawing.Size(448, 20)
        Me.txtPlikTXT.TabIndex = 78
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(12, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(190, 16)
        Me.Label2.TabIndex = 76
        Me.Label2.Text = "Plik .txt z wykazem Bran¿ . . . . . . . . ."
        '
        'cmdPobierz
        '
        Me.cmdPobierz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPobierz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPobierz.ForeColor = System.Drawing.Color.Black
        Me.cmdPobierz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPobierz.Location = New System.Drawing.Point(430, 50)
        Me.cmdPobierz.Name = "cmdPobierz"
        Me.cmdPobierz.Size = New System.Drawing.Size(119, 22)
        Me.cmdPobierz.TabIndex = 56
        Me.cmdPobierz.Text = "Pobierz dane z pliku"
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 50)
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
        Me.cmdPowrot.Location = New System.Drawing.Point(557, 50)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(87, 22)
        Me.cmdPowrot.TabIndex = 54
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'frmPobierzZPlikuBranze
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(653, 78)
        Me.Controls.Add(Me.cmdPobierz)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.Panel1)
        Me.ForeColor = System.Drawing.Color.Purple
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPobierzZPlikuBranze"
        Me.ShowInTaskbar = False
        Me.Text = "Pobieranie danych z pliku"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim oPlik As String = ""

    Private Sub frmPobierzZPlikuBranze_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        oPlik = ""
        Select Case CoPobieramy
            Case "BRANZE"
                Label2.Text = "Plik .txt z wykazem Bran¿ . . . . . . . . ."
            Case "PODBRANZE"
                Label2.Text = "Plik .txt z wykazem Podbran¿ . . . . . . . . ."
        End Select
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub frmPobierzZPlikuBranze_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdWybierzPlik_Click(sender As Object, e As EventArgs) Handles cmdWybierzPlik.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik tekstowy z Bran¿ami"
            .InitialDirectory = Trim("C:\")
            .DefaultExt = "txt"
            .Filter = "Plik tekstowy (*.txt)|*.txt|Wszystkie pliki|*.*"
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                txtPlikTXT.Text = .FileName
            End If
        End With
    End Sub


    Private Sub cmdPobierz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPobierz.Click
        Dim FileNum As Integer
        Dim NextLine As String
        Dim PozSeparator As Integer = 0
        Dim robBranza As String = ""
        Dim robPodBranza As String = ""

        Dim findBranze(0) As Object
        Dim findPodbranze(1) As Object
        Dim foundRow As DataRow
        Dim NewRow As DataRow

        If Len(txtPlikTXT.Text) = 0 Then
            MessageBox.Show("Proszê zdefiniowaæ plik do pobrania...",
                                        "Nie wiem co pobraæ",
                                        System.Windows.Forms.MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Question,
                                        MessageBoxDefaultButton.Button1)
        Else
            If Not IO.File.Exists(txtPlikTXT.Text) Then
                MessageBox.Show("Nie znalaz³em pliku" & vbCrLf & txtPlikTXT.Text,
                                        "Przykro mi...",
                                        System.Windows.Forms.MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Question,
                                        MessageBoxDefaultButton.Button1)
            Else
                Select Case CoPobieramy
                    Case "BRANZE"
                        If MessageBox.Show("Czy mam pobraæ BRAN¯E z pliku tekstowego ?",
                            "Proszê o potwierdzenie :",
                            System.Windows.Forms.MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question,
                            MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then


                            FileNum = FreeFile()    'kolejny nr otwartego zbioru
                            FileOpen(FileNum, txtPlikTXT.Text, OpenMode.Input)
                            If Not EOF(FileNum) Then
                                NextLine = LineInput(FileNum)
                                'Branza|Podbranza
                                Do While (Len(NextLine) > 0) And (Not EOF(FileNum))
                                    PozSeparator = InStr(NextLine, "|")
                                    robBranza = Mid(NextLine, 1, PozSeparator - 1)
                                    robPodBranza = Mid(NextLine, PozSeparator + 1)

                                    ' Create an array for the key values to find.
                                    findBranze(0) = robBranza
                                    foundRow = _Dataset.Tables(0).Rows.Find(findBranze)
                                    ' sprawdz czy jest...
                                    If Not (foundRow Is Nothing) Then
                                    Else
                                        'stworz nowy rekord
                                        NewRow = _Dataset.Tables(0).NewRow()
                                        NewRow("IDBRANZY") = Mid(robBranza, 1, 100)
                                        NewRow("WERSJA") = 1                     'Wersja         Integer, 2
                                        _Dataset.Tables(0).Rows.Add(NewRow)
                                    End If
                                    NextLine = LineInput(FileNum)
                                Loop
                                FileClose(FileNum)
                            End If

                            MessageBox.Show("Pobra³em BRAN¯E z pliku tekstowego...",
                                        "OK, skoñczy³em...",
                                        System.Windows.Forms.MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Question,
                                        MessageBoxDefaultButton.Button1)

                        Else
                        End If




                    Case "PODBRANZE"
                        If MessageBox.Show("Czy mam pobraæ PODBRAN¯E z pliku tekstowego ?",
                            "Proszê o potwierdzenie :",
                            System.Windows.Forms.MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question,
                            MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then

                            FileNum = FreeFile()    'kolejny nr otwartego zbioru
                            FileOpen(FileNum, txtPlikTXT.Text, OpenMode.Input)
                            If Not EOF(FileNum) Then
                                NextLine = LineInput(FileNum)
                                'Branza|Podbranza
                                Do While (Len(NextLine) > 0) And (Not EOF(FileNum))
                                    PozSeparator = InStr(NextLine, "|")
                                    robBranza = Mid(NextLine, 1, PozSeparator - 1)
                                    robPodBranza = Mid(NextLine, PozSeparator + 1)

                                    ' Create an array for the key values to find.
                                    findPodbranze(0) = robBranza
                                    findPodbranze(1) = robPodBranza
                                    foundRow = _Dataset.Tables(0).Rows.Find(findPodbranze)
                                    ' sprawdz czy jest...
                                    If Not (foundRow Is Nothing) Then
                                    Else
                                        'stworz nowy rekord
                                        NewRow = _Dataset.Tables(0).NewRow()
                                        NewRow("IDBRANZY") = Mid(robBranza, 1, 100)
                                        NewRow("IDPODBRANZY") = Mid(robPodBranza, 1, 100)
                                        NewRow("WERSJA") = 1                     'Wersja         Integer, 2
                                        _Dataset.Tables(0).Rows.Add(NewRow)
                                    End If
                                    NextLine = LineInput(FileNum)
                                Loop
                                FileClose(FileNum)
                            End If

                            MessageBox.Show("Pobra³em PODBRAN¯E z pliku tekstowego...",
                                        "OK, skoñczy³em...",
                                        System.Windows.Forms.MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Question,
                                        MessageBoxDefaultButton.Button1)

                        Else
                        End If




                End Select
            End If
        End If
    End Sub

End Class
