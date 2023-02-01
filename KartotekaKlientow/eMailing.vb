Public Class eMailing
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

    Private _dView As DataView
    Private _eMailAdres As String
    Private _eMailKlient As String

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef dv As DataView, _
                   ByVal eAdres As String, _
                   ByVal eKlient As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _dView = dv
        _eMailAdres = eAdres
        _eMailKlient = eAdres
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
    Friend WithEvents cmdWysylaj As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmdWybierz As System.Windows.Forms.Button
    Friend WithEvents lblZalacznik As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtTresc As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblAdreseMail As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtTemat As System.Windows.Forms.TextBox
    Friend WithEvents dlgOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(eMailing))
        Me.cmdWysylaj = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtTemat = New System.Windows.Forms.TextBox()
        Me.cmdWybierz = New System.Windows.Forms.Button()
        Me.lblZalacznik = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtTresc = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblAdreseMail = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dlgOpen = New System.Windows.Forms.OpenFileDialog()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdWysylaj
        '
        Me.cmdWysylaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWysylaj.Image = CType(resources.GetObject("cmdWysylaj.Image"), System.Drawing.Image)
        Me.cmdWysylaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWysylaj.Location = New System.Drawing.Point(314, 358)
        Me.cmdWysylaj.Name = "cmdWysylaj"
        Me.cmdWysylaj.Size = New System.Drawing.Size(80, 22)
        Me.cmdWysylaj.TabIndex = 19
        Me.cmdWysylaj.Text = "Wyœlij"
        Me.cmdWysylaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(411, 358)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 22)
        Me.cmdPowrot.TabIndex = 18
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 358)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 23)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 17
        Me.PictureBox2.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.txtTemat)
        Me.Panel1.Controls.Add(Me.cmdWybierz)
        Me.Panel1.Controls.Add(Me.lblZalacznik)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.txtTresc)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.lblAdreseMail)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(483, 341)
        Me.Panel1.TabIndex = 16
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label5.Location = New System.Drawing.Point(72, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(392, 56)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = resources.GetString("Label5.Text")
        '
        'txtTemat
        '
        Me.txtTemat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTemat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtTemat.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtTemat.Location = New System.Drawing.Point(72, 96)
        Me.txtTemat.Name = "txtTemat"
        Me.txtTemat.Size = New System.Drawing.Size(394, 20)
        Me.txtTemat.TabIndex = 16
        '
        'cmdWybierz
        '
        Me.cmdWybierz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierz.Image = CType(resources.GetObject("cmdWybierz.Image"), System.Drawing.Image)
        Me.cmdWybierz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWybierz.Location = New System.Drawing.Point(362, 309)
        Me.cmdWybierz.Name = "cmdWybierz"
        Me.cmdWybierz.Size = New System.Drawing.Size(104, 25)
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
        Me.lblZalacznik.Location = New System.Drawing.Point(72, 286)
        Me.lblZalacznik.Name = "lblZalacznik"
        Me.lblZalacznik.Size = New System.Drawing.Size(394, 15)
        Me.lblZalacznik.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(12, 286)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 15)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Za³¹cznik"
        '
        'txtTresc
        '
        Me.txtTresc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTresc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtTresc.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtTresc.Location = New System.Drawing.Point(72, 120)
        Me.txtTresc.Multiline = True
        Me.txtTresc.Name = "txtTresc"
        Me.txtTresc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtTresc.Size = New System.Drawing.Size(394, 157)
        Me.txtTresc.TabIndex = 12
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 16)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Treœæ:"
        '
        'lblAdreseMail
        '
        Me.lblAdreseMail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAdreseMail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAdreseMail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblAdreseMail.ForeColor = System.Drawing.Color.Purple
        Me.lblAdreseMail.Location = New System.Drawing.Point(72, 72)
        Me.lblAdreseMail.Name = "lblAdreseMail"
        Me.lblAdreseMail.Size = New System.Drawing.Size(394, 16)
        Me.lblAdreseMail.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Temat:"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 72)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Adresat:"
        '
        'eMailing
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(498, 387)
        Me.Controls.Add(Me.cmdWysylaj)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "eMailing"
        Me.ShowInTaskbar = False
        Me.Text = "Rozsy³anie poczty elektronicznej do klientów"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub eMaileing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        If Len(_eMailAdres) = 0 Then
            lblAdreseMail.Text = "...adresy poszczególnych klientów..."
        Else
            lblAdreseMail.Text = _eMailAdres
        End If
        txtTemat.Text = ""
        txtTresc.Text = ""
        lblZalacznik.Text = ""
    End Sub

    Private Sub cmdWysylaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWysylaj.Click
        If TestujParametryeMail() Then
            If Len(Trim(txtTemat.Text)) = 0 Then
                MessageBox.Show("Proszê zdefiniowaæ TEMAT listu", _
                                "Nie mogê rozes³aæ poczty...", _
                                MessageBoxButtons.OK, _
                                MessageBoxIcon.Exclamation, _
                                MessageBoxDefaultButton.Button1)
            Else
                If Len(Trim(txtTresc.Text)) = 0 Then
                    MessageBox.Show("Proszê zdefiniowaæ TREŒÆ listu", _
                                    "Nie mogê rozes³aæ poczty...", _
                                    MessageBoxButtons.OK, _
                                    MessageBoxIcon.Exclamation, _
                                    MessageBoxDefaultButton.Button1)
                Else
                    If Len(_eMailAdres) = 0 Then
                        Dim Form As New WysylanieeMaili(_dView, txtTemat.Text, txtTresc.Text, lblZalacznik.Text)
                        Form.ShowDialog()
                    Else
                        Dim Form As New WyslijeMail(_eMailAdres, _eMailKlient, txtTemat.Text, txtTresc.Text, lblZalacznik.Text, False)
                        Form.ShowDialog()
                    End If
                End If
            End If
        End If
    End Sub

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
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                sNewFile = .FileName
                If IO.File.Exists(sNewFile) Then
                    lblZalacznik.Text = sNewFile
                Else
                    MessageBox.Show("Nie znalaz³em na dysku wybranego pliku", _
                                    "Przykro mi", _
                                    MessageBoxButtons.OK, _
                                    MessageBoxIcon.Exclamation, _
                                    MessageBoxDefaultButton.Button1)
                End If
            End If
            'Me.Visible = True
        End With
    End Sub

    Private Sub eMailing_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub
End Class
