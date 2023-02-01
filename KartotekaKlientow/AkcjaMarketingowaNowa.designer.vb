<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AkcjaMarketingowaNowa
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

    Dim _dvKlenci As DataView

    Public Sub New(ByVal dv As DataView)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _dvKlenci = dv
    End Sub




    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AkcjaMarketingowaNowa))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.dtpData = New System.Windows.Forms.DateTimePicker()
        Me.txtUwagi = New System.Windows.Forms.TextBox()
        Me.txtIleZap = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtIdent = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtOpis = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.cmdZapamietaj = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.dtpData)
        Me.Panel1.Controls.Add(Me.txtUwagi)
        Me.Panel1.Controls.Add(Me.txtIleZap)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.txtIdent)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.txtOpis)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(348, 222)
        Me.Panel1.TabIndex = 0
        '
        'dtpData
        '
        Me.dtpData.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtpData.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpData.CustomFormat = "yyyy-MM-dd"
        Me.dtpData.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpData.Location = New System.Drawing.Point(80, 56)
        Me.dtpData.Name = "dtpData"
        Me.dtpData.Size = New System.Drawing.Size(109, 20)
        Me.dtpData.TabIndex = 182
        Me.dtpData.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'txtUwagi
        '
        Me.txtUwagi.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUwagi.ForeColor = System.Drawing.Color.Purple
        Me.txtUwagi.Location = New System.Drawing.Point(80, 104)
        Me.txtUwagi.Multiline = True
        Me.txtUwagi.Name = "txtUwagi"
        Me.txtUwagi.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtUwagi.Size = New System.Drawing.Size(259, 110)
        Me.txtUwagi.TabIndex = 22
        '
        'txtIleZap
        '
        Me.txtIleZap.ForeColor = System.Drawing.Color.Purple
        Me.txtIleZap.Location = New System.Drawing.Point(80, 8)
        Me.txtIleZap.Name = "txtIleZap"
        Me.txtIleZap.ReadOnly = True
        Me.txtIleZap.Size = New System.Drawing.Size(144, 20)
        Me.txtIleZap.TabIndex = 20
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(7, 11)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 18)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "Iloœæ zapisów . . . . . . . . ."
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(7, 107)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Uwagi . . . . . . . . ."
        '
        'txtIdent
        '
        Me.txtIdent.ForeColor = System.Drawing.Color.Purple
        Me.txtIdent.Location = New System.Drawing.Point(80, 32)
        Me.txtIdent.Name = "txtIdent"
        Me.txtIdent.Size = New System.Drawing.Size(144, 20)
        Me.txtIdent.TabIndex = 13
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(7, 35)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Ident. akcji . . . . . . . . ."
        '
        'txtOpis
        '
        Me.txtOpis.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOpis.ForeColor = System.Drawing.Color.Purple
        Me.txtOpis.Location = New System.Drawing.Point(80, 80)
        Me.txtOpis.Name = "txtOpis"
        Me.txtOpis.Size = New System.Drawing.Size(259, 20)
        Me.txtOpis.TabIndex = 17
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(7, 59)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 16)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Data akcji . . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(7, 83)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 16)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Opis akcji . . . . . . . . ."
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 236)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 40
        Me.PictureBox1.TabStop = False
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ForeColor = System.Drawing.Color.Black
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(274, 237)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(80, 23)
        Me.cmdStop.TabIndex = 41
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdZapamietaj
        '
        Me.cmdZapamietaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdZapamietaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdZapamietaj.ForeColor = System.Drawing.Color.Black
        Me.cmdZapamietaj.Image = CType(resources.GetObject("cmdZapamietaj.Image"), System.Drawing.Image)
        Me.cmdZapamietaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdZapamietaj.Location = New System.Drawing.Point(150, 237)
        Me.cmdZapamietaj.Name = "cmdZapamietaj"
        Me.cmdZapamietaj.Size = New System.Drawing.Size(118, 23)
        Me.cmdZapamietaj.TabIndex = 42
        Me.cmdZapamietaj.Text = "Zapamiêtaj akcje"
        Me.cmdZapamietaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'AkcjaMarketingowaNowa
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(366, 266)
        Me.Controls.Add(Me.cmdZapamietaj)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "AkcjaMarketingowaNowa"
        Me.ShowInTaskbar = False
        Me.Text = "Nowa akcja marketingowa"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents cmdZapamietaj As System.Windows.Forms.Button
    Friend WithEvents txtUwagi As System.Windows.Forms.TextBox
    Friend WithEvents txtIleZap As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtIdent As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtOpis As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpData As DateTimePicker
End Class
