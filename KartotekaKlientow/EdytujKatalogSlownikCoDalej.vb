Public Class EdytujKatalogSlownikCoDalej
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
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtIdent As System.Windows.Forms.TextBox
    Friend WithEvents txtOpis As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EdytujKatalogSlownikCoDalej))
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtOpis = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtIdent = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(427, 195)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 24)
        Me.cmdPowrot.TabIndex = 13
        Me.cmdPowrot.Text = "Zapisz"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 195)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(88, 25)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 16
        Me.PictureBox2.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.txtOpis)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.txtIdent)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(499, 181)
        Me.Panel1.TabIndex = 17
        '
        'txtOpis
        '
        Me.txtOpis.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOpis.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtOpis.ForeColor = System.Drawing.Color.Purple
        Me.txtOpis.Location = New System.Drawing.Point(91, 31)
        Me.txtOpis.Multiline = True
        Me.txtOpis.Name = "txtOpis"
        Me.txtOpis.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtOpis.Size = New System.Drawing.Size(390, 136)
        Me.txtOpis.TabIndex = 164
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(11, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(124, 16)
        Me.Label2.TabIndex = 165
        Me.Label2.Text = "Kiedy stosowaæ . . . . . . . . ."
        '
        'txtIdent
        '
        Me.txtIdent.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtIdent.ForeColor = System.Drawing.Color.Purple
        Me.txtIdent.Location = New System.Drawing.Point(91, 8)
        Me.txtIdent.Name = "txtIdent"
        Me.txtIdent.Size = New System.Drawing.Size(390, 20)
        Me.txtIdent.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(11, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(124, 16)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Identyfikator . . . . . . . . ."
        '
        'EdytujKatalogSlownikCoDalej
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(515, 220)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Location = New System.Drawing.Point(8, 8)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "EdytujKatalogSlownikCoDalej"
        Me.ShowInTaskbar = False
        Me.Text = "Co Dalej"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub EdytujKatalogSlownikCoDalej_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        PobierzDane()
        If oOperacja = "EDYTUJ" Then
            txtIdent.Enabled = False
        Else
        End If
    End Sub

    Private Sub EdytujKatalogSlownikCoDalej_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If oOperacja = "EDYTUJ" Then
            txtOpis.Focus()
        Else
            txtIdent.Focus()
        End If
    End Sub

    Private Sub EdytujKatalogSlownikCoDalej_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        ZapiszDane()
        oAktualizuj = True
        Me.Close()
    End Sub

    '====================================================
    'UWAGA - trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    Private Sub EdytujKatalogSlownikCoDalej_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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



    '*******************************************
    '** procedury obslugi edycji
    '*******************************************

    Private Sub txtRodzajKontaktu_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIdent.TextChanged
        TestLen(txtIdent, "IDENTYFIKATOR", 20)
    End Sub
    Private Sub txtNazwa_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdent.GotFocus
        StartEdycjiTxt(txtIdent)
    End Sub
    Private Sub txtNazwa_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIdent.LostFocus
        KoniecEdycjiTxt(txtIdent)
    End Sub


    Private Sub txtOPIS_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOpis.TextChanged
        'TestLen(txtOpis, "KIEDY STOSOWAC", 20)
    End Sub
    Private Sub txtOPIS_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpis.GotFocus
        StartEdycjiTxt(txtOpis)
    End Sub
    Private Sub txtOPIS_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpis.LostFocus
        KoniecEdycjiTxt(txtOpis)
    End Sub



    '***************************************************
    '* procedury pobierania/zapisywania danych
    '***************************************************

    Private Sub PobierzDane()
        txtIdent.Text = scdIDENT
        txtOpis.Text = scdOPIS
    End Sub

    Private Sub ZapiszDane()
        scdIDENT = txtIdent.Text
        scdOPIS = txtOpis.Text
    End Sub


End Class
