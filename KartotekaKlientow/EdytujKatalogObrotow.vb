Public Class EdytujKatalogObrotow
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
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblNazwa2 As System.Windows.Forms.Label
    Friend WithEvents lblNazwa1 As System.Windows.Forms.Label
    Friend WithEvents lblNazwaHandlowa As System.Windows.Forms.Label
    Friend WithEvents lblIdent As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdWycofajSie As System.Windows.Forms.Button
    Friend WithEvents cmdPrzywroc As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents txtAWartSprzed As System.Windows.Forms.TextBox
    Friend WithEvents txtAIloSprzed As System.Windows.Forms.TextBox
    Friend WithEvents txtLWartSprzed As System.Windows.Forms.TextBox
    Friend WithEvents txtLIloSprzed As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents cmdPKontakt As System.Windows.Forms.Button
    Friend WithEvents txtLoIloSprzed As System.Windows.Forms.TextBox
    Friend WithEvents txtAoWartSprzed As System.Windows.Forms.TextBox
    Friend WithEvents txtAoIloSprzed As System.Windows.Forms.TextBox
    Friend WithEvents txtLoWartSprzed As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As Panel
    Friend WithEvents txtSUMAIlo As TextBox
    Friend WithEvents Label14 As Label
    Friend WithEvents txtSUMAWart As TextBox
    Friend WithEvents txtSKUPIloSprzed As TextBox
    Friend WithEvents Label13 As Label
    Friend WithEvents txtSKUPWartSprzed As TextBox
    Friend WithEvents txtDRUKARKIIloSprzed As TextBox
    Friend WithEvents Label12 As Label
    Friend WithEvents txtDRUKARKIWartSprzed As TextBox
    Friend WithEvents txtSTRONYIloSprzed As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents txtSTRONYWartSprzed As TextBox
    Friend WithEvents txtNAJEMIloSprzed As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtNAJEMWartSprzed As TextBox
    Friend WithEvents TxtData As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EdytujKatalogObrotow))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.txtSUMAIlo = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtSUMAWart = New System.Windows.Forms.TextBox()
        Me.txtSKUPIloSprzed = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtSKUPWartSprzed = New System.Windows.Forms.TextBox()
        Me.txtDRUKARKIIloSprzed = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtDRUKARKIWartSprzed = New System.Windows.Forms.TextBox()
        Me.txtSTRONYIloSprzed = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtSTRONYWartSprzed = New System.Windows.Forms.TextBox()
        Me.txtNAJEMIloSprzed = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtNAJEMWartSprzed = New System.Windows.Forms.TextBox()
        Me.txtLIloSprzed = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtLoIloSprzed = New System.Windows.Forms.TextBox()
        Me.txtAoWartSprzed = New System.Windows.Forms.TextBox()
        Me.txtAoIloSprzed = New System.Windows.Forms.TextBox()
        Me.txtLoWartSprzed = New System.Windows.Forms.TextBox()
        Me.cmdPKontakt = New System.Windows.Forms.Button()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblNazwa2 = New System.Windows.Forms.Label()
        Me.lblNazwa1 = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa = New System.Windows.Forms.Label()
        Me.lblIdent = New System.Windows.Forms.Label()
        Me.txtAWartSprzed = New System.Windows.Forms.TextBox()
        Me.txtAIloSprzed = New System.Windows.Forms.TextBox()
        Me.txtLWartSprzed = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TxtData = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
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
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Panel5)
        Me.Panel1.Controls.Add(Me.txtSUMAIlo)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Controls.Add(Me.txtSUMAWart)
        Me.Panel1.Controls.Add(Me.txtSKUPIloSprzed)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.txtSKUPWartSprzed)
        Me.Panel1.Controls.Add(Me.txtDRUKARKIIloSprzed)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.txtDRUKARKIWartSprzed)
        Me.Panel1.Controls.Add(Me.txtSTRONYIloSprzed)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.txtSTRONYWartSprzed)
        Me.Panel1.Controls.Add(Me.txtNAJEMIloSprzed)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.txtNAJEMWartSprzed)
        Me.Panel1.Controls.Add(Me.txtLIloSprzed)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.txtLoIloSprzed)
        Me.Panel1.Controls.Add(Me.txtAoWartSprzed)
        Me.Panel1.Controls.Add(Me.txtAoIloSprzed)
        Me.Panel1.Controls.Add(Me.txtLoWartSprzed)
        Me.Panel1.Controls.Add(Me.cmdPKontakt)
        Me.Panel1.Controls.Add(Me.Panel6)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.lblNazwa2)
        Me.Panel1.Controls.Add(Me.lblNazwa1)
        Me.Panel1.Controls.Add(Me.lblNazwaHandlowa)
        Me.Panel1.Controls.Add(Me.lblIdent)
        Me.Panel1.Controls.Add(Me.txtAWartSprzed)
        Me.Panel1.Controls.Add(Me.txtAIloSprzed)
        Me.Panel1.Controls.Add(Me.txtLWartSprzed)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.TxtData)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Location = New System.Drawing.Point(9, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(710, 347)
        Me.Panel1.TabIndex = 27
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Location = New System.Drawing.Point(170, 300)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(420, 1)
        Me.Panel3.TabIndex = 72
        '
        'Panel5
        '
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Location = New System.Drawing.Point(338, 96)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(1, 240)
        Me.Panel5.TabIndex = 34
        '
        'txtSUMAIlo
        '
        Me.txtSUMAIlo.ForeColor = System.Drawing.Color.Purple
        Me.txtSUMAIlo.Location = New System.Drawing.Point(346, 306)
        Me.txtSUMAIlo.Name = "txtSUMAIlo"
        Me.txtSUMAIlo.Size = New System.Drawing.Size(104, 20)
        Me.txtSUMAIlo.TabIndex = 71
        Me.txtSUMAIlo.Text = "0"
        Me.txtSUMAIlo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label14
        '
        Me.Label14.ForeColor = System.Drawing.Color.Navy
        Me.Label14.Location = New System.Drawing.Point(187, 309)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(153, 16)
        Me.Label14.TabIndex = 70
        Me.Label14.Text = "SUMA . . . . . . . . . . . . . . . . . . ."
        '
        'txtSUMAWart
        '
        Me.txtSUMAWart.ForeColor = System.Drawing.Color.Purple
        Me.txtSUMAWart.Location = New System.Drawing.Point(476, 306)
        Me.txtSUMAWart.Name = "txtSUMAWart"
        Me.txtSUMAWart.Size = New System.Drawing.Size(104, 20)
        Me.txtSUMAWart.TabIndex = 69
        Me.txtSUMAWart.Text = "0"
        Me.txtSUMAWart.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSKUPIloSprzed
        '
        Me.txtSKUPIloSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtSKUPIloSprzed.Location = New System.Drawing.Point(346, 274)
        Me.txtSKUPIloSprzed.Name = "txtSKUPIloSprzed"
        Me.txtSKUPIloSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtSKUPIloSprzed.TabIndex = 68
        Me.txtSKUPIloSprzed.Text = "0"
        Me.txtSKUPIloSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label13
        '
        Me.Label13.ForeColor = System.Drawing.Color.Navy
        Me.Label13.Location = New System.Drawing.Point(187, 277)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(153, 16)
        Me.Label13.TabIndex = 67
        Me.Label13.Text = "SKUP . . . . . . . . . . . . . . . . . . . . . ."
        '
        'txtSKUPWartSprzed
        '
        Me.txtSKUPWartSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtSKUPWartSprzed.Location = New System.Drawing.Point(476, 274)
        Me.txtSKUPWartSprzed.Name = "txtSKUPWartSprzed"
        Me.txtSKUPWartSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtSKUPWartSprzed.TabIndex = 66
        Me.txtSKUPWartSprzed.Text = "0"
        Me.txtSKUPWartSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtDRUKARKIIloSprzed
        '
        Me.txtDRUKARKIIloSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtDRUKARKIIloSprzed.Location = New System.Drawing.Point(346, 252)
        Me.txtDRUKARKIIloSprzed.Name = "txtDRUKARKIIloSprzed"
        Me.txtDRUKARKIIloSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtDRUKARKIIloSprzed.TabIndex = 65
        Me.txtDRUKARKIIloSprzed.Text = "0"
        Me.txtDRUKARKIIloSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label12
        '
        Me.Label12.ForeColor = System.Drawing.Color.Navy
        Me.Label12.Location = New System.Drawing.Point(187, 255)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(153, 16)
        Me.Label12.TabIndex = 64
        Me.Label12.Text = "DRUKARKI . . . . . . . . . . . . . . . . . ."
        '
        'txtDRUKARKIWartSprzed
        '
        Me.txtDRUKARKIWartSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtDRUKARKIWartSprzed.Location = New System.Drawing.Point(476, 252)
        Me.txtDRUKARKIWartSprzed.Name = "txtDRUKARKIWartSprzed"
        Me.txtDRUKARKIWartSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtDRUKARKIWartSprzed.TabIndex = 63
        Me.txtDRUKARKIWartSprzed.Text = "0"
        Me.txtDRUKARKIWartSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSTRONYIloSprzed
        '
        Me.txtSTRONYIloSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtSTRONYIloSprzed.Location = New System.Drawing.Point(346, 230)
        Me.txtSTRONYIloSprzed.Name = "txtSTRONYIloSprzed"
        Me.txtSTRONYIloSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtSTRONYIloSprzed.TabIndex = 62
        Me.txtSTRONYIloSprzed.Text = "0"
        Me.txtSTRONYIloSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(187, 233)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(153, 16)
        Me.Label6.TabIndex = 61
        Me.Label6.Text = "STRONY . . . . . . . . . . . . . . . . . . . ."
        '
        'txtSTRONYWartSprzed
        '
        Me.txtSTRONYWartSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtSTRONYWartSprzed.Location = New System.Drawing.Point(476, 230)
        Me.txtSTRONYWartSprzed.Name = "txtSTRONYWartSprzed"
        Me.txtSTRONYWartSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtSTRONYWartSprzed.TabIndex = 60
        Me.txtSTRONYWartSprzed.Text = "0"
        Me.txtSTRONYWartSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtNAJEMIloSprzed
        '
        Me.txtNAJEMIloSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtNAJEMIloSprzed.Location = New System.Drawing.Point(346, 208)
        Me.txtNAJEMIloSprzed.Name = "txtNAJEMIloSprzed"
        Me.txtNAJEMIloSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtNAJEMIloSprzed.TabIndex = 59
        Me.txtNAJEMIloSprzed.Text = "0"
        Me.txtNAJEMIloSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(187, 211)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(153, 16)
        Me.Label2.TabIndex = 58
        Me.Label2.Text = "NAJEM . . . . . . . . . . . . . . . . . . . ."
        '
        'txtNAJEMWartSprzed
        '
        Me.txtNAJEMWartSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtNAJEMWartSprzed.Location = New System.Drawing.Point(476, 208)
        Me.txtNAJEMWartSprzed.Name = "txtNAJEMWartSprzed"
        Me.txtNAJEMWartSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtNAJEMWartSprzed.TabIndex = 57
        Me.txtNAJEMWartSprzed.Text = "0"
        Me.txtNAJEMWartSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtLIloSprzed
        '
        Me.txtLIloSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtLIloSprzed.Location = New System.Drawing.Point(346, 142)
        Me.txtLIloSprzed.Name = "txtLIloSprzed"
        Me.txtLIloSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtLIloSprzed.TabIndex = 7
        Me.txtLIloSprzed.Text = "0"
        Me.txtLIloSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Green
        Me.Label18.Location = New System.Drawing.Point(346, 94)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(104, 16)
        Me.Label18.TabIndex = 56
        Me.Label18.Text = "ILOŒÆ"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Green
        Me.Label17.Location = New System.Drawing.Point(474, 94)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(104, 16)
        Me.Label17.TabIndex = 55
        Me.Label17.Text = "WARTOŒÆ"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtLoIloSprzed
        '
        Me.txtLoIloSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtLoIloSprzed.Location = New System.Drawing.Point(346, 120)
        Me.txtLoIloSprzed.Name = "txtLoIloSprzed"
        Me.txtLoIloSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtLoIloSprzed.TabIndex = 41
        Me.txtLoIloSprzed.Text = "0"
        Me.txtLoIloSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAoWartSprzed
        '
        Me.txtAoWartSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtAoWartSprzed.Location = New System.Drawing.Point(476, 164)
        Me.txtAoWartSprzed.Name = "txtAoWartSprzed"
        Me.txtAoWartSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtAoWartSprzed.TabIndex = 44
        Me.txtAoWartSprzed.Text = "0"
        Me.txtAoWartSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAoIloSprzed
        '
        Me.txtAoIloSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtAoIloSprzed.Location = New System.Drawing.Point(346, 164)
        Me.txtAoIloSprzed.Name = "txtAoIloSprzed"
        Me.txtAoIloSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtAoIloSprzed.TabIndex = 43
        Me.txtAoIloSprzed.Text = "0"
        Me.txtAoIloSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtLoWartSprzed
        '
        Me.txtLoWartSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtLoWartSprzed.Location = New System.Drawing.Point(476, 120)
        Me.txtLoWartSprzed.Name = "txtLoWartSprzed"
        Me.txtLoWartSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtLoWartSprzed.TabIndex = 42
        Me.txtLoWartSprzed.Text = "0"
        Me.txtLoWartSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cmdPKontakt
        '
        Me.cmdPKontakt.Image = CType(resources.GetObject("cmdPKontakt.Image"), System.Drawing.Image)
        Me.cmdPKontakt.Location = New System.Drawing.Point(114, 95)
        Me.cmdPKontakt.Name = "cmdPKontakt"
        Me.cmdPKontakt.Size = New System.Drawing.Size(24, 22)
        Me.cmdPKontakt.TabIndex = 36
        '
        'Panel6
        '
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel6.Location = New System.Drawing.Point(463, 96)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(1, 240)
        Me.Panel6.TabIndex = 35
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Location = New System.Drawing.Point(170, 114)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(420, 1)
        Me.Panel4.TabIndex = 33
        '
        'Label10
        '
        Me.Label10.ForeColor = System.Drawing.Color.Navy
        Me.Label10.Location = New System.Drawing.Point(187, 189)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(153, 16)
        Me.Label10.TabIndex = 31
        Me.Label10.Text = "ATRAMENT Pryzmat . . . . . . . . . . . ."
        '
        'Label11
        '
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(187, 167)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(153, 16)
        Me.Label11.TabIndex = 30
        Me.Label11.Text = "ATRAMENT Oryginalne . . . . . . . . . . . ."
        '
        'Label9
        '
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(187, 145)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(153, 16)
        Me.Label9.TabIndex = 29
        Me.Label9.Text = "LASEROWE Pryzmat . . . . . . . . . . . ."
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(0, 88)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(710, 1)
        Me.Panel2.TabIndex = 23
        '
        'Label8
        '
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(8, 48)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 16)
        Me.Label8.TabIndex = 22
        Me.Label8.Text = "Adres klienta . . . . . . . . ."
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(8, 32)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 16)
        Me.Label7.TabIndex = 21
        Me.Label7.Text = "Nazwa klienta . . . . . . . . ."
        '
        'lblNazwa2
        '
        Me.lblNazwa2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa2.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa2.Location = New System.Drawing.Point(120, 64)
        Me.lblNazwa2.Name = "lblNazwa2"
        Me.lblNazwa2.Size = New System.Drawing.Size(574, 16)
        Me.lblNazwa2.TabIndex = 20
        '
        'lblNazwa1
        '
        Me.lblNazwa1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa1.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa1.Location = New System.Drawing.Point(120, 48)
        Me.lblNazwa1.Name = "lblNazwa1"
        Me.lblNazwa1.Size = New System.Drawing.Size(574, 16)
        Me.lblNazwa1.TabIndex = 19
        '
        'lblNazwaHandlowa
        '
        Me.lblNazwaHandlowa.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwaHandlowa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwaHandlowa.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwaHandlowa.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwaHandlowa.Location = New System.Drawing.Point(120, 32)
        Me.lblNazwaHandlowa.Name = "lblNazwaHandlowa"
        Me.lblNazwaHandlowa.Size = New System.Drawing.Size(574, 16)
        Me.lblNazwaHandlowa.TabIndex = 18
        '
        'lblIdent
        '
        Me.lblIdent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdent.ForeColor = System.Drawing.Color.Purple
        Me.lblIdent.Location = New System.Drawing.Point(120, 8)
        Me.lblIdent.Name = "lblIdent"
        Me.lblIdent.Size = New System.Drawing.Size(100, 24)
        Me.lblIdent.TabIndex = 17
        Me.lblIdent.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtAWartSprzed
        '
        Me.txtAWartSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtAWartSprzed.Location = New System.Drawing.Point(476, 186)
        Me.txtAWartSprzed.Name = "txtAWartSprzed"
        Me.txtAWartSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtAWartSprzed.TabIndex = 10
        Me.txtAWartSprzed.Text = "0"
        Me.txtAWartSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtAIloSprzed
        '
        Me.txtAIloSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtAIloSprzed.Location = New System.Drawing.Point(346, 186)
        Me.txtAIloSprzed.Name = "txtAIloSprzed"
        Me.txtAIloSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtAIloSprzed.TabIndex = 9
        Me.txtAIloSprzed.Text = "0"
        Me.txtAIloSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtLWartSprzed
        '
        Me.txtLWartSprzed.ForeColor = System.Drawing.Color.Purple
        Me.txtLWartSprzed.Location = New System.Drawing.Point(476, 142)
        Me.txtLWartSprzed.Name = "txtLWartSprzed"
        Me.txtLWartSprzed.Size = New System.Drawing.Size(104, 20)
        Me.txtLWartSprzed.TabIndex = 8
        Me.txtLWartSprzed.Text = "0"
        Me.txtLWartSprzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(190, 96)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(142, 16)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "SPRZEDA¯"
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(187, 123)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(153, 16)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "LASEROWE Oryginalne. . . . . . . . . . . ."
        '
        'TxtData
        '
        Me.TxtData.ForeColor = System.Drawing.Color.Purple
        Me.TxtData.Location = New System.Drawing.Point(40, 96)
        Me.TxtData.Name = "TxtData"
        Me.TxtData.Size = New System.Drawing.Size(88, 20)
        Me.TxtData.TabIndex = 1
        Me.TxtData.Text = "2004"
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 96)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 16)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Data . . . . . . . . ."
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Ident. klienta . . . . . . . . ."
        '
        'cmdWycofajSie
        '
        Me.cmdWycofajSie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWycofajSie.Image = CType(resources.GetObject("cmdWycofajSie.Image"), System.Drawing.Image)
        Me.cmdWycofajSie.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWycofajSie.Location = New System.Drawing.Point(529, 363)
        Me.cmdWycofajSie.Name = "cmdWycofajSie"
        Me.cmdWycofajSie.Size = New System.Drawing.Size(104, 24)
        Me.cmdWycofajSie.TabIndex = 12
        Me.cmdWycofajSie.Text = "Wycofaj siê"
        Me.cmdWycofajSie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPrzywroc
        '
        Me.cmdPrzywroc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPrzywroc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPrzywroc.Location = New System.Drawing.Point(432, 363)
        Me.cmdPrzywroc.Name = "cmdPrzywroc"
        Me.cmdPrzywroc.Size = New System.Drawing.Size(91, 24)
        Me.cmdPrzywroc.TabIndex = 13
        Me.cmdPrzywroc.Text = "Przywróæ"
        Me.cmdPrzywroc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(639, 363)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 24)
        Me.cmdPowrot.TabIndex = 11
        Me.cmdPowrot.Text = "Zapisz"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(9, 363)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(88, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 26
        Me.PictureBox2.TabStop = False
        '
        'EdytujKatalogObrotow
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(728, 390)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdWycofajSie)
        Me.Controls.Add(Me.cmdPrzywroc)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "EdytujKatalogObrotow"
        Me.ShowInTaskbar = False
        Me.Text = "Edytuj Katalog Obrotów"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub EdytujKatalogObrotow_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        PobierzDane()
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
            TxtData.Enabled = False
            cmdPKontakt.Enabled = False
        Else
            TxtData.Enabled = True
            cmdPKontakt.Enabled = True
        End If
        '---------------------
        IdentIdKlienta(oIdentObr)
        lblIdent.Text = o2IdentKli
        'lblTOFI.Text = o2NrTOFITxtKli
        lblNazwaHandlowa.Text = o2Nazwa1Kli
        lblNazwa1.Text = o2UlicaKli
        lblNazwa2.Text = o2KodPoczKli & "  " & o2MiejscKli
    End Sub

    Private Sub EdytujKatalogObrotow_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub EdytujKatalogObrotow_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If oOperacja = "EDYTUJ" Then
            'nie pozwalaj na edycja PrimaryKey
            txtLIloSprzed.Focus()
        Else
            TxtData.Focus()
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

    Private Sub cmdPKontakt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPKontakt.Click
        oData = TxtData.Text
        Dim Proc As New WybierzDate
        Proc.ShowDialog()
        TxtData.Text = oData
    End Sub

    '====================================================
    'UWAGA - trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    Private Sub EdytujKatalogObrotow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
        TxtData.Text = oDataObr

        txtAIloSprzed.Text = oAIloSprzedObr.ToString("F0")
        txtAWartSprzed.Text = oAWartSprzedObr.ToString("F2")

        txtLIloSprzed.Text = oLIloSprzedObr.ToString("F0")
        txtLWartSprzed.Text = oLWartSprzedObr.ToString("F2")

        txtAoIloSprzed.Text = oAOIloSprzedObr.ToString("F0")
        txtAoWartSprzed.Text = oAOWartSprzedObr.ToString("F2")

        txtLoIloSprzed.Text = oLOIloSprzedObr.ToString("F0")
        txtLoWartSprzed.Text = oLOWartSprzedObr.ToString("F2")

        txtNAJEMIloSprzed.Text = oNAJEMIloSprzedObr.ToString("F0")
        txtNAJEMWartSprzed.Text = oNAJEMWartSprzedObr.ToString("F2")

        txtSTRONYIloSprzed.Text = oSTRONYIloSprzedObr.ToString("F0")
        txtSTRONYWartSprzed.Text = oSTRONYWartSprzedObr.ToString("F2")

        txtDRUKARKIIloSprzed.Text = oDRUKARKIIloSprzedObr.ToString("F0")
        txtDRUKARKIWartSprzed.Text = oDRUKARKIWartSprzedObr.ToString("F2")

        txtSKUPIloSprzed.Text = oSKUPIloSprzedObr.ToString("F0")
        txtSKUPWartSprzed.Text = oSKUPWartSprzedObr.ToString("F2")

        txtSUMAIlo.Text = (oLIloSprzedObr + oAIloSprzedObr +
                          oLOIloSprzedObr + oAOIloSprzedObr +
                          oNAJEMIloSprzedObr + oSTRONYIloSprzedObr +
                          oDRUKARKIIloSprzedObr + oSKUPIloSprzedObr).ToString("F0")
        txtSUMAWart.Text = (oLWartSprzedObr + oAWartSprzedObr +
                           oLOWartSprzedObr + oAOWartSprzedObr +
                           oNAJEMWartSprzedObr + oSTRONYWartSprzedObr +
                           oDRUKARKIWartSprzedObr + oSKUPWartSprzedObr).ToString("F2")

    End Sub

    Private Sub WyliczSumeWartosc()
        txtSUMAWart.Text = (GetDblFromText(txtAWartSprzed.Text) + GetDblFromText(txtLWartSprzed.Text) +
                          GetDblFromText(txtAoWartSprzed.Text) + GetDblFromText(txtLoWartSprzed.Text) +
                          GetDblFromText(txtNAJEMWartSprzed.Text) + GetDblFromText(txtSTRONYWartSprzed.Text) +
                          GetDblFromText(txtDRUKARKIWartSprzed.Text) + GetDblFromText(txtSKUPWartSprzed.Text)).ToString("F2")
    End Sub

    Private Sub WyliczSumeIlosc()
        txtSUMAIlo.Text = (GetDblFromText(txtAIloSprzed.Text) + GetDblFromText(txtLIloSprzed.Text) +
                          GetDblFromText(txtAoIloSprzed.Text) + GetDblFromText(txtLoIloSprzed.Text) +
                          GetDblFromText(txtNAJEMIloSprzed.Text) + GetDblFromText(txtSTRONYIloSprzed.Text) +
                          GetDblFromText(txtDRUKARKIIloSprzed.Text) + GetDblFromText(txtSKUPIloSprzed.Text)).ToString("F0")
    End Sub

    Private Sub ZapiszDane()
        oDataObr = TxtData.Text

        oAIloSprzedObr = GetDblFromText(txtAIloSprzed.Text)
        oAWartSprzedObr = GetDblFromText(txtAWartSprzed.Text)

        oLIloSprzedObr = GetDblFromText(txtLIloSprzed.Text)
        oLWartSprzedObr = GetDblFromText(txtLWartSprzed.Text)

        oAOIloSprzedObr = GetDblFromText(txtAoIloSprzed.Text)
        oAOWartSprzedObr = GetDblFromText(txtAoWartSprzed.Text)

        oLOIloSprzedObr = GetDblFromText(txtLoIloSprzed.Text)
        oLOWartSprzedObr = GetDblFromText(txtLoWartSprzed.Text)

        oNAJEMIloSprzedObr = GetDblFromText(txtNAJEMIloSprzed.Text)
        oNAJEMWartSprzedObr = GetDblFromText(txtNAJEMWartSprzed.Text)

        oSTRONYIloSprzedObr = GetDblFromText(txtSTRONYIloSprzed.Text)
        oSTRONYWartSprzedObr = GetDblFromText(txtSTRONYWartSprzed.Text)

        oDRUKARKIIloSprzedObr = GetDblFromText(txtDRUKARKIIloSprzed.Text)
        oDRUKARKIWartSprzedObr = GetDblFromText(txtDRUKARKIWartSprzed.Text)

        oSKUPIloSprzedObr = GetDblFromText(txtSKUPIloSprzed.Text)
        oSKUPWartSprzedObr = GetDblFromText(txtSKUPWartSprzed.Text)

    End Sub















    '**************************************************************
    '** obsluga obiektu w trakcie edycji
    '**************************************************************

    Private Sub txtData_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtData.TextChanged
        TestLen(TxtData, "DATA", 10)
    End Sub
    Private Sub txtData_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtData.GotFocus
        StartEdycjiTxt(TxtData)
    End Sub
    Private Sub txtData_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtData.LostFocus
        KoniecEdycjiTxt(TxtData)
    End Sub







    Private Sub txtLIloSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLIloSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtLIloSprzed, "Sprzed L ILOŒÆ")
    End Sub
    Private Sub txtLIloSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLIloSprzed.GotFocus
        StartEdycjiTxt(txtLIloSprzed)
    End Sub
    Private Sub txtLIloSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLIloSprzed.LostFocus
        KoniecEdycjiTxt(txtLIloSprzed)
        WyliczSumeIlosc()
    End Sub

    Private Sub txtLWartSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLWartSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtLWartSprzed, "Sprzed L WARTOŒÆ")
    End Sub
    Private Sub txtLWartSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLWartSprzed.GotFocus
        StartEdycjiTxt(txtLWartSprzed)
    End Sub
    Private Sub txtLWartSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLWartSprzed.LostFocus
        KoniecEdycjiTxt(txtLWartSprzed)
        WyliczSumeWartosc()
    End Sub

    Private Sub txtAIloSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAIloSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtAIloSprzed, "Sprzed A ILOŒÆ")
    End Sub
    Private Sub txtAIloSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAIloSprzed.GotFocus
        StartEdycjiTxt(txtAIloSprzed)
    End Sub
    Private Sub txtAIloSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAIloSprzed.LostFocus
        KoniecEdycjiTxt(txtAIloSprzed)
        WyliczSumeIlosc()
    End Sub

    Private Sub txtAWartSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAWartSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtAWartSprzed, "Sprzed A WARTOŒÆ")
    End Sub
    Private Sub txtAWartSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAWartSprzed.GotFocus
        StartEdycjiTxt(txtAWartSprzed)
    End Sub
    Private Sub txtAWartSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAWartSprzed.LostFocus
        KoniecEdycjiTxt(txtAWartSprzed)
        WyliczSumeWartosc()
    End Sub







    Private Sub txtLOIloSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLoIloSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtLoIloSprzed, "Sprzed ORYGINA£Y L ILOŒÆ")
    End Sub
    Private Sub txtLOIloSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLoIloSprzed.GotFocus
        StartEdycjiTxt(txtLoIloSprzed)
    End Sub
    Private Sub txtLOIloSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLoIloSprzed.LostFocus
        KoniecEdycjiTxt(txtLoIloSprzed)
        WyliczSumeIlosc()
    End Sub

    Private Sub txtLOWartSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLoWartSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtLoWartSprzed, "Sprzed ORYGINA£Y L WARTOŒÆ")
    End Sub
    Private Sub txtLOWartSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLoWartSprzed.GotFocus
        StartEdycjiTxt(txtLoWartSprzed)
    End Sub
    Private Sub txtLOWartSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLoWartSprzed.LostFocus
        KoniecEdycjiTxt(txtLoWartSprzed)
        WyliczSumeWartosc()
    End Sub

    Private Sub txtAOIloSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAoIloSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtAoIloSprzed, "Sprzed ORYGINA£Y A ILOŒÆ")
    End Sub
    Private Sub txtAOIloSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAoIloSprzed.GotFocus
        StartEdycjiTxt(txtAoIloSprzed)
    End Sub
    Private Sub txtAOIloSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAoIloSprzed.LostFocus
        KoniecEdycjiTxt(txtAoIloSprzed)
        WyliczSumeIlosc()
    End Sub

    Private Sub txtAOWartSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAoWartSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtAoWartSprzed, "Sprzed ORYGINA£Y A WARTOŒÆ")
    End Sub
    Private Sub txtAOWartSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAoWartSprzed.GotFocus
        StartEdycjiTxt(txtAoWartSprzed)
    End Sub
    Private Sub txtAOWartSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAoWartSprzed.LostFocus
        KoniecEdycjiTxt(txtAoWartSprzed)
        WyliczSumeWartosc()
    End Sub


    Private Sub txtNAJEMIloSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNAJEMIloSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtNAJEMIloSprzed, "NAJEM ILOŒÆ")
    End Sub
    Private Sub txtNAJEMIloSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNAJEMIloSprzed.GotFocus
        StartEdycjiTxt(txtNAJEMIloSprzed)
    End Sub
    Private Sub txtNAJEMIloSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNAJEMIloSprzed.LostFocus
        KoniecEdycjiTxt(txtNAJEMIloSprzed)
        WyliczSumeIlosc()
    End Sub


    Private Sub txtNAJEMWartSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNAJEMWartSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtNAJEMWartSprzed, "NAJEM WARTOŒÆ")
    End Sub
    Private Sub txtNAJEMWartSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNAJEMWartSprzed.GotFocus
        StartEdycjiTxt(txtNAJEMWartSprzed)
    End Sub
    Private Sub txtNAJEMWartSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNAJEMWartSprzed.LostFocus
        KoniecEdycjiTxt(txtNAJEMWartSprzed)
        WyliczSumeWartosc()
    End Sub


    Private Sub txtSTRONYIloSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSTRONYIloSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtSTRONYIloSprzed, "STRONY ILOŒÆ")
    End Sub
    Private Sub txtSTRONYIloSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSTRONYIloSprzed.GotFocus
        StartEdycjiTxt(txtSTRONYIloSprzed)
    End Sub
    Private Sub txtSTRONYIloSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSTRONYIloSprzed.LostFocus
        KoniecEdycjiTxt(txtSTRONYIloSprzed)
        WyliczSumeIlosc()
    End Sub


    Private Sub txtSTRONYWartSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSTRONYWartSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtSTRONYWartSprzed, "STRONY WARTOŒÆ")
    End Sub
    Private Sub txtSTRONYWartSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSTRONYWartSprzed.GotFocus
        StartEdycjiTxt(txtSTRONYWartSprzed)
    End Sub
    Private Sub txtSTRONYWartSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSTRONYWartSprzed.LostFocus
        KoniecEdycjiTxt(txtSTRONYWartSprzed)
        WyliczSumeWartosc()
    End Sub




    Private Sub txtDRUKARKIIloSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDRUKARKIIloSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtDRUKARKIIloSprzed, "DRUKARKI ILOŒÆ")
    End Sub
    Private Sub txtDRUKARKIIloSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDRUKARKIIloSprzed.GotFocus
        StartEdycjiTxt(txtDRUKARKIIloSprzed)
    End Sub
    Private Sub txtDRUKARKIIloSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDRUKARKIIloSprzed.LostFocus
        KoniecEdycjiTxt(txtDRUKARKIIloSprzed)
        WyliczSumeIlosc()
    End Sub


    Private Sub txtDRUKARKIWartSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDRUKARKIWartSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtDRUKARKIWartSprzed, "DRUKARKI WARTOŒÆ")
    End Sub
    Private Sub txtDRUKARKIWartSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDRUKARKIWartSprzed.GotFocus
        StartEdycjiTxt(txtDRUKARKIWartSprzed)
    End Sub
    Private Sub txtDRUKARKIWartSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDRUKARKIWartSprzed.LostFocus
        KoniecEdycjiTxt(txtDRUKARKIWartSprzed)
        WyliczSumeWartosc()
    End Sub




    Private Sub txtSKUPIloSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSKUPIloSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtSKUPIloSprzed, "SKUP ILOŒÆ")
    End Sub
    Private Sub txtSKUPIloSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSKUPIloSprzed.GotFocus
        StartEdycjiTxt(txtSKUPIloSprzed)
    End Sub
    Private Sub txtSKUPIloSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSKUPIloSprzed.LostFocus
        KoniecEdycjiTxt(txtSKUPIloSprzed)
        WyliczSumeIlosc()
    End Sub


    Private Sub txtSKUPWartSprzed_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSKUPWartSprzed.TextChanged
        Dim tst As Boolean
        tst = TestNum(txtSKUPWartSprzed, "SKUP WARTOŒÆ")
    End Sub
    Private Sub txtSKUPWartSprzed_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSKUPWartSprzed.GotFocus
        StartEdycjiTxt(txtSKUPWartSprzed)
    End Sub
    Private Sub txtSKUPWartSprzed_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSKUPWartSprzed.LostFocus
        KoniecEdycjiTxt(txtSKUPWartSprzed)
        WyliczSumeWartosc()
    End Sub



End Class
