Public Class EdytujKatalogPodbranze
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
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdWycofajSie As System.Windows.Forms.Button
    Friend WithEvents cmdPrzywroc As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents txtPodbranza As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents cmdWybierzBranze As Button
    Friend WithEvents txtBranza As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EdytujKatalogPodbranze))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdWybierzBranze = New System.Windows.Forms.Button()
        Me.txtPodbranza = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtBranza = New System.Windows.Forms.TextBox()
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
        Me.Panel1.Controls.Add(Me.cmdWybierzBranze)
        Me.Panel1.Controls.Add(Me.txtPodbranza)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.txtBranza)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(589, 66)
        Me.Panel1.TabIndex = 2
        '
        'cmdWybierzBranze
        '
        Me.cmdWybierzBranze.Image = CType(resources.GetObject("cmdWybierzBranze.Image"), System.Drawing.Image)
        Me.cmdWybierzBranze.Location = New System.Drawing.Point(534, 7)
        Me.cmdWybierzBranze.Name = "cmdWybierzBranze"
        Me.cmdWybierzBranze.Size = New System.Drawing.Size(32, 23)
        Me.cmdWybierzBranze.TabIndex = 90
        '
        'txtPodbranza
        '
        Me.txtPodbranza.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPodbranza.ForeColor = System.Drawing.Color.Purple
        Me.txtPodbranza.Location = New System.Drawing.Point(68, 31)
        Me.txtPodbranza.Name = "txtPodbranza"
        Me.txtPodbranza.Size = New System.Drawing.Size(498, 20)
        Me.txtPodbranza.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 16)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Podbranza . . . . . . . . ."
        '
        'txtBranza
        '
        Me.txtBranza.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBranza.ForeColor = System.Drawing.Color.Purple
        Me.txtBranza.Location = New System.Drawing.Point(68, 9)
        Me.txtBranza.Name = "txtBranza"
        Me.txtBranza.ReadOnly = True
        Me.txtBranza.Size = New System.Drawing.Size(498, 20)
        Me.txtBranza.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Branza . . . . . . . . ."
        '
        'cmdWycofajSie
        '
        Me.cmdWycofajSie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWycofajSie.Image = CType(resources.GetObject("cmdWycofajSie.Image"), System.Drawing.Image)
        Me.cmdWycofajSie.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWycofajSie.Location = New System.Drawing.Point(409, 81)
        Me.cmdWycofajSie.Name = "cmdWycofajSie"
        Me.cmdWycofajSie.Size = New System.Drawing.Size(104, 25)
        Me.cmdWycofajSie.TabIndex = 12
        Me.cmdWycofajSie.Text = "Wycofaj siê"
        Me.cmdWycofajSie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPrzywroc
        '
        Me.cmdPrzywroc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPrzywroc.Image = CType(resources.GetObject("cmdPrzywroc.Image"), System.Drawing.Image)
        Me.cmdPrzywroc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPrzywroc.Location = New System.Drawing.Point(312, 81)
        Me.cmdPrzywroc.Name = "cmdPrzywroc"
        Me.cmdPrzywroc.Size = New System.Drawing.Size(91, 25)
        Me.cmdPrzywroc.TabIndex = 13
        Me.cmdPrzywroc.Text = "Przywróæ"
        Me.cmdPrzywroc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(519, 81)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 25)
        Me.cmdPowrot.TabIndex = 11
        Me.cmdPowrot.Text = "Zapisz"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 81)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(88, 25)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 21
        Me.PictureBox2.TabStop = False
        '
        'EdytujKatalogPodbranze
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(604, 112)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdWycofajSie)
        Me.Controls.Add(Me.cmdPrzywroc)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "EdytujKatalogPodbranze"
        Me.ShowInTaskbar = False
        Me.Text = "Bran¿a"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub EdytujKatalogPodbranze_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        PobierzDane()
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
            txtBranza.Enabled = False
            cmdWybierzBranze.Enabled = False
            txtPodbranza.Enabled = False
        Else
            txtBranza.Enabled = True
            cmdWybierzBranze.Enabled = True
            txtPodbranza.Enabled = True
        End If
    End Sub

    Private Sub EdytujKatalogPodbranze_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If oOperacja = "EDYTUJ" Then
        Else
            txtBranza.Focus()
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

    Private Sub EdytujKatalogPodbranze_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub


    '====================================================
    'UWAGA - trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    Private Sub EdytujKatalogPodbranze_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me.ClosedByControlBox Then
            'e.Cancel = True     'nie pozwalaj zamkn¹c forme
            If MessageBox.Show("Zdecydowa³eœ opuœciæ formê bez zapisania wprowadzonych zmian." & vbCrLf &
                "Czy mam zapisaæ zmiany w Bazie Danych ?",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
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
        txtBranza.Text = pbrIdBranzy
        txtPodbranza.Text = pbrIdPodBranzy
    End Sub

    Private Sub ZapiszDane()
        pbrIdBranzy = txtBranza.Text
        pbrIdPodBranzy = txtPodbranza.Text
    End Sub

    '*******************************************
    '** procedury obslugi edycji
    '*******************************************

    Private Sub txtBranza_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBranza.TextChanged
        TestLen(txtBranza, "BRAN¯A", 100)
    End Sub
    Private Sub txtBranza_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBranza.GotFocus
        StartEdycjiTxt(txtBranza)
    End Sub
    Private Sub txtBranza_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBranza.LostFocus
        KoniecEdycjiTxt(txtBranza)
    End Sub



    Private Sub txtPodBranza_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPodbranza.TextChanged
        TestLen(txtPodbranza, "PODBRAN¯A", 100)
    End Sub
    Private Sub txtPodBranza_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPodbranza.GotFocus
        StartEdycjiTxt(txtPodbranza)
    End Sub
    Private Sub txtPodBranza_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPodbranza.LostFocus
        KoniecEdycjiTxt(txtPodbranza)
    End Sub



    Private Sub cmdWybierzBranze_Click(sender As Object, e As EventArgs) Handles cmdWybierzBranze.Click
        oCoMamRobic = "WYBIERAJ"
        oIdBranzy = Trim(txtBranza.Text)
        Dim form As New KatalogBranze
        form.ShowDialog()
        If Len(Trim(oIdBranzy)) > 0 Then
            txtBranza.Text = oIdBranzy
        End If
    End Sub


End Class
