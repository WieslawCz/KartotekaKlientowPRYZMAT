Imports System
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Threading
Imports System.ComponentModel

Public Class SendeMail
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm


#Region " Windows Form Designer generated code "

    Dim _eMailZalacznik As String
    Dim _eMailTemat As String
    Dim _eMailAdres As String
    Dim _eMailTresc As String

    Public Sub New(ByVal pAdres As String, _
                   ByVal pTemat As String, _
                   ByVal pTresc As String, _
                   ByVal pZalacznik As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _eMailAdres = pAdres
        _eMailTemat = pTemat
        _eMailTresc = pTresc
        _eMailZalacznik = pZalacznik
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
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtTresc As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dlgOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents lblZalacznik As System.Windows.Forms.Label
    Friend WithEvents cmdWybierz As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents cmdWysylaj As System.Windows.Forms.Button
    Friend WithEvents txtAdres As System.Windows.Forms.TextBox
    Friend WithEvents txtTemat As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SendeMail))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtTemat = New System.Windows.Forms.TextBox()
        Me.txtAdres = New System.Windows.Forms.TextBox()
        Me.cmdWybierz = New System.Windows.Forms.Button()
        Me.lblZalacznik = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtTresc = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dlgOpen = New System.Windows.Forms.OpenFileDialog()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.cmdWysylaj = New System.Windows.Forms.Button()
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
        Me.Panel1.Controls.Add(Me.txtTemat)
        Me.Panel1.Controls.Add(Me.txtAdres)
        Me.Panel1.Controls.Add(Me.cmdWybierz)
        Me.Panel1.Controls.Add(Me.lblZalacznik)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.txtTresc)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(480, 303)
        Me.Panel1.TabIndex = 7
        '
        'txtTemat
        '
        Me.txtTemat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTemat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtTemat.ForeColor = System.Drawing.Color.Purple
        Me.txtTemat.Location = New System.Drawing.Point(68, 40)
        Me.txtTemat.Name = "txtTemat"
        Me.txtTemat.Size = New System.Drawing.Size(400, 20)
        Me.txtTemat.TabIndex = 17
        '
        'txtAdres
        '
        Me.txtAdres.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAdres.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtAdres.ForeColor = System.Drawing.Color.Purple
        Me.txtAdres.Location = New System.Drawing.Point(68, 16)
        Me.txtAdres.Name = "txtAdres"
        Me.txtAdres.Size = New System.Drawing.Size(400, 20)
        Me.txtAdres.TabIndex = 16
        '
        'cmdWybierz
        '
        Me.cmdWybierz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierz.Image = CType(resources.GetObject("cmdWybierz.Image"), System.Drawing.Image)
        Me.cmdWybierz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWybierz.Location = New System.Drawing.Point(359, 271)
        Me.cmdWybierz.Name = "cmdWybierz"
        Me.cmdWybierz.Size = New System.Drawing.Size(105, 25)
        Me.cmdWybierz.TabIndex = 15
        Me.cmdWybierz.Text = "Wybierz plik"
        Me.cmdWybierz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblZalacznik
        '
        Me.lblZalacznik.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblZalacznik.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblZalacznik.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblZalacznik.ForeColor = System.Drawing.Color.Purple
        Me.lblZalacznik.Location = New System.Drawing.Point(68, 247)
        Me.lblZalacznik.Name = "lblZalacznik"
        Me.lblZalacznik.Size = New System.Drawing.Size(396, 17)
        Me.lblZalacznik.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(12, 247)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 17)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Za³¹cznik"
        '
        'txtTresc
        '
        Me.txtTresc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTresc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtTresc.ForeColor = System.Drawing.Color.Purple
        Me.txtTresc.Location = New System.Drawing.Point(68, 64)
        Me.txtTresc.Multiline = True
        Me.txtTresc.Name = "txtTresc"
        Me.txtTresc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtTresc.Size = New System.Drawing.Size(400, 167)
        Me.txtTresc.TabIndex = 12
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(12, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 16)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Treœæ . . . . . . . . . . "
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(12, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Temat . . . . . . . . "
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(12, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Adres . . . . . . "
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 319)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 12
        Me.PictureBox2.TabStop = False
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(399, 319)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 23)
        Me.cmdPowrot.TabIndex = 13
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdWysylaj
        '
        Me.cmdWysylaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWysylaj.Image = CType(resources.GetObject("cmdWysylaj.Image"), System.Drawing.Image)
        Me.cmdWysylaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWysylaj.Location = New System.Drawing.Point(304, 319)
        Me.cmdWysylaj.Name = "cmdWysylaj"
        Me.cmdWysylaj.Size = New System.Drawing.Size(80, 23)
        Me.cmdWysylaj.TabIndex = 15
        Me.cmdWysylaj.Text = "Wyœlij"
        Me.cmdWysylaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'SendeMail
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(494, 348)
        Me.Controls.Add(Me.cmdWysylaj)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Panel1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SendeMail"
        Me.ShowInTaskbar = False
        Me.Text = "Poczta elektroniczna"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub cmdWybierz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierz.Click
        Dim sNewFile As String
        With dlgOpen
            .Filter = "Pliki do³¹czane: (*.*)|*.*"
            .AddExtension = True
            .CheckFileExists = True
            .CheckPathExists = True
            .InitialDirectory = "c:\"
            .Multiselect = False

            'Me.Visible = False
            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                sNewFile = .FileName
                If IO.File.Exists(sNewFile) Then
                    lblZalacznik.Text = sNewFile
                Else
                    MessageBox.Show("Nie znalaz³em na dysku wybranego pliku", "Przykro mi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1)
                End If
            End If
            'Me.Visible = True
        End With
    End Sub

    Private Sub SendeMail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        lblZalacznik.Text = _eMailZalacznik
        txtTemat.Text = _eMailTemat
        txtAdres.Text = _eMailAdres
        txtTresc.Text = _eMailTresc
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub cmdWysylaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWysylaj.Click
        If Len(Trim(ParL_eMailAdres)) = 0 Then
            MessageBox.Show("Brak adresu nadawcy (w parametrach programu)...", _
                            "Uwaga :", _
                            MessageBoxButtons.OK, _
                            MessageBoxIcon.Exclamation)
        Else
            If Len(Trim(txtAdres.Text)) = 0 Then
                MessageBox.Show("Brak adresu odbiorcy...", _
                                "Uwaga :", _
                                MessageBoxButtons.OK, _
                                MessageBoxIcon.Exclamation)
            Else
                If Len(Trim(txtTemat.Text)) = 0 Then
                    MessageBox.Show("Brak tematu przesy³ki...", _
                                    "Uwaga :", _
                                    MessageBoxButtons.OK, _
                                    MessageBoxIcon.Exclamation)
                Else
                    If Len(Trim(txtTresc.Text)) = 0 Then
                        MessageBox.Show("Brak treœci przesy³ki...", _
                                        "Uwaga :", _
                                        MessageBoxButtons.OK, _
                                        MessageBoxIcon.Exclamation)
                    Else
                        WyslijeMail2(Trim(txtAdres.Text), txtTemat.Text, txtTresc.Text, lblZalacznik.Text)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub SendeMail_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub txtAdres_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAdres.GotFocus
        StartEdycjiTxt(txtAdres)
    End Sub
    Private Sub txtAdres_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAdres.LostFocus
        KoniecEdycjiTxt(txtAdres)
    End Sub
    Private Sub txtAdres_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAdres.TextChanged
    End Sub

    Private Sub txtTemat_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTemat.GotFocus
        StartEdycjiTxt(txtTemat)
    End Sub
    Private Sub txtTemat_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTemat.LostFocus
        KoniecEdycjiTxt(txtTemat)
    End Sub
    Private Sub txtTemat_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTemat.TextChanged
    End Sub

    Private Sub txtTresc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTresc.GotFocus
        StartEdycjiTxt(txtTresc)
    End Sub
    Private Sub txtTresc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTresc.LostFocus
        KoniecEdycjiTxt(txtTresc)
    End Sub
    Private Sub txtTresc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTresc.TextChanged
    End Sub
End Class
