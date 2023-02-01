Public Class EdytujKatalogAkcjeSpec
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "


    'Fields
    'Constructors
    'Events
    'Methods

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
    Friend WithEvents cmdKlient As System.Windows.Forms.Button
    Friend WithEvents cmdWycofajSie As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtAkcja As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtKlient As System.Windows.Forms.TextBox
    Friend WithEvents lblNazwa As System.Windows.Forms.Label
    Friend WithEvents lblMiejscowosc As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblAdres As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EdytujKatalogAkcjeSpec))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdKlient = New System.Windows.Forms.Button()
        Me.lblAdres = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblMiejscowosc = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblNazwa = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtKlient = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtAkcja = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdWycofajSie = New System.Windows.Forms.Button()
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
        Me.Panel1.Controls.Add(Me.cmdKlient)
        Me.Panel1.Controls.Add(Me.lblAdres)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.lblMiejscowosc)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.lblNazwa)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.txtKlient)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.txtAkcja)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(471, 144)
        Me.Panel1.TabIndex = 17
        '
        'cmdKlient
        '
        Me.cmdKlient.Image = CType(resources.GetObject("cmdKlient.Image"), System.Drawing.Image)
        Me.cmdKlient.Location = New System.Drawing.Point(197, 38)
        Me.cmdKlient.Name = "cmdKlient"
        Me.cmdKlient.Size = New System.Drawing.Size(32, 23)
        Me.cmdKlient.TabIndex = 16
        '
        'lblAdres
        '
        Me.lblAdres.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAdres.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblAdres.ForeColor = System.Drawing.Color.Navy
        Me.lblAdres.Location = New System.Drawing.Point(128, 112)
        Me.lblAdres.Name = "lblAdres"
        Me.lblAdres.Size = New System.Drawing.Size(328, 16)
        Me.lblAdres.TabIndex = 15
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(16, 112)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(140, 16)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Adres klienta . . . . . . . . ."
        '
        'lblMiejscowosc
        '
        Me.lblMiejscowosc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblMiejscowosc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblMiejscowosc.ForeColor = System.Drawing.Color.Navy
        Me.lblMiejscowosc.Location = New System.Drawing.Point(128, 88)
        Me.lblMiejscowosc.Name = "lblMiejscowosc"
        Me.lblMiejscowosc.Size = New System.Drawing.Size(328, 16)
        Me.lblMiejscowosc.TabIndex = 13
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(16, 88)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(140, 16)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Miejscowoœæ . . . . . . . . ."
        '
        'lblNazwa
        '
        Me.lblNazwa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa.ForeColor = System.Drawing.Color.Navy
        Me.lblNazwa.Location = New System.Drawing.Point(128, 64)
        Me.lblNazwa.Name = "lblNazwa"
        Me.lblNazwa.Size = New System.Drawing.Size(328, 16)
        Me.lblNazwa.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(16, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(140, 16)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Nazwa klienta . . . . . . . . ."
        '
        'txtKlient
        '
        Me.txtKlient.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtKlient.ForeColor = System.Drawing.Color.Purple
        Me.txtKlient.Location = New System.Drawing.Point(128, 40)
        Me.txtKlient.Name = "txtKlient"
        Me.txtKlient.Size = New System.Drawing.Size(72, 20)
        Me.txtKlient.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(16, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(140, 16)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Identyfikator klienta . . . . . . . . ."
        '
        'txtAkcja
        '
        Me.txtAkcja.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtAkcja.ForeColor = System.Drawing.Color.Purple
        Me.txtAkcja.Location = New System.Drawing.Point(128, 16)
        Me.txtAkcja.Name = "txtAkcja"
        Me.txtAkcja.Size = New System.Drawing.Size(144, 20)
        Me.txtAkcja.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(16, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(140, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Identyfikator Akcji . . . . . . . . ."
        '
        'cmdWycofajSie
        '
        Me.cmdWycofajSie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWycofajSie.Image = CType(resources.GetObject("cmdWycofajSie.Image"), System.Drawing.Image)
        Me.cmdWycofajSie.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWycofajSie.Location = New System.Drawing.Point(289, 160)
        Me.cmdWycofajSie.Name = "cmdWycofajSie"
        Me.cmdWycofajSie.Size = New System.Drawing.Size(104, 23)
        Me.cmdWycofajSie.TabIndex = 14
        Me.cmdWycofajSie.Text = "Wycofaj siê"
        Me.cmdWycofajSie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(399, 160)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 23)
        Me.cmdPowrot.TabIndex = 13
        Me.cmdPowrot.Text = "Zapisz"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 160)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 16
        Me.PictureBox2.TabStop = False
        '
        'EdytujKatalogAkcjeSpec
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(488, 188)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdWycofajSie)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(8, 8)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "EdytujKatalogAkcjeSpec"
        Me.ShowInTaskbar = False
        Me.Text = "Edycja Informacji o uczestniku akcji"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub EdytujKatalogAkcjeSpec_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        PobierzDane()
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
            txtAkcja.ReadOnly = True
            txtKlient.ReadOnly = True
        Else
            txtAkcja.ReadOnly = True
            txtKlient.ReadOnly = False
            txtKlient.Focus()
        End If
    End Sub

    Private Sub EdytujKatalogAkcjeSpec_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If oOperacja = "EDYTUJ" Then
            cmdPowrot.Focus()
        Else
            txtKlient.Focus()
        End If
    End Sub

    Private Sub EdytujKatalogAkcjeSpec_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        ZapiszDane()
        If Len(Trim(asIdentKlienta)) > 0 Then
            oAktualizuj = True
        Else
            oAktualizuj = False
        End If
        Me.Close()
    End Sub


    Private Sub cmdWycofajSie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWycofajSie.Click
        oAktualizuj = False
        Me.Close()
    End Sub


    '====================================================
    'UWAGA - trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    Private Sub EdytujKatalogAkcjeSpec_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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

    Private Sub txtKlient_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKlient.TextChanged
        TestLen(txtKlient, "IDENTYFIKATOR KLIENTA", 6)
        '-------------------
        IdentIdKlienta(txtKlient.Text)
        lblNazwa.Text = o2Nazwa1Kli
        lblMiejscowosc.Text = o2KodPoczKli & "  " & o2MiejscKli
        lblAdres.Text = o2UlicaKli
    End Sub
    Private Sub txtKlient_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKlient.GotFocus
        StartEdycjiTxt(txtKlient)
    End Sub
    Private Sub txtKlient_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKlient.LostFocus
        KoniecEdycjiTxt(txtKlient)
    End Sub

    Private Sub txtIdent_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAkcja.TextChanged
        TestLen(txtAkcja, "IDENTYFIKATOR AKCJI", 20)
    End Sub
    Private Sub txtIdent_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAkcja.GotFocus
        StartEdycjiTxt(txtAkcja)
    End Sub
    Private Sub txtIdent_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAkcja.LostFocus
        KoniecEdycjiTxt(txtAkcja)
    End Sub


    '***************************************************
    '* procedury pobierania/zapisywania danych
    '***************************************************

    Private Sub PobierzDane()
        txtKlient.Text = asIdentKlienta
        txtAkcja.Text = asIdentAkcji

        IdentIdKlienta(txtKlient.Text)
        lblNazwa.Text = o2Nazwa1Kli
        lblMiejscowosc.Text = o2KodPoczKli & "  " & o2MiejscKli
        lblAdres.Text = o2UlicaKli
    End Sub

    Private Sub ZapiszDane()
        asIdentAkcji = txtAkcja.Text
        asIdentKlienta = txtKlient.Text
    End Sub



    Private Sub cmdKlient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKlient.Click
        oCoMamRobic = "WYBIERAJ"
        oKlient = Trim(txtKlient.Text)
        Dim form As New KatalogKlientow
        form.ShowDialog()
        If Len(Trim(oKlient)) > 0 Then
            txtKlient.Text = oKlient
        End If
    End Sub

End Class