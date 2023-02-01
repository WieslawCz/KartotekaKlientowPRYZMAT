<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEdytujTabele
    Inherits System.Windows.Forms.Form

    Dim _Generator As String = ""

    Public Sub New(ByRef parG As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _Generator = parG
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEdytujTabele))
        Me.dagKlienci = New System.Windows.Forms.DataGrid()
        Me.txtNazwa = New System.Windows.Forms.TextBox()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.lblSzerokosc = New System.Windows.Forms.Label()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        CType(Me.dagKlienci, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dagKlienci
        '
        Me.dagKlienci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagKlienci.DataMember = ""
        Me.dagKlienci.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dagKlienci.Location = New System.Drawing.Point(12, 61)
        Me.dagKlienci.Name = "dagKlienci"
        Me.dagKlienci.Size = New System.Drawing.Size(739, 264)
        Me.dagKlienci.TabIndex = 53
        '
        'txtNazwa
        '
        Me.txtNazwa.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNazwa.ForeColor = System.Drawing.Color.Purple
        Me.txtNazwa.Location = New System.Drawing.Point(12, 12)
        Me.txtNazwa.Name = "txtNazwa"
        Me.txtNazwa.Size = New System.Drawing.Size(739, 20)
        Me.txtNazwa.TabIndex = 63
        Me.txtNazwa.Text = "Data"
        Me.txtNazwa.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ForeColor = System.Drawing.Color.Black
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(671, 331)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(80, 23)
        Me.cmdStop.TabIndex = 64
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(12, 331)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 65
        Me.PictureBox2.TabStop = False
        '
        'lblSzerokosc
        '
        Me.lblSzerokosc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblSzerokosc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSzerokosc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblSzerokosc.ForeColor = System.Drawing.Color.Navy
        Me.lblSzerokosc.Location = New System.Drawing.Point(142, 334)
        Me.lblSzerokosc.Name = "lblSzerokosc"
        Me.lblSzerokosc.Size = New System.Drawing.Size(193, 17)
        Me.lblSzerokosc.TabIndex = 66
        Me.lblSzerokosc.Text = "Szerokoœæ wydruku : 0 mm"
        '
        'stbFiltry
        '
        Me.stbFiltry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltry.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltry.Location = New System.Drawing.Point(12, 38)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(739, 22)
        Me.stbFiltry.TabIndex = 68
        Me.stbFiltry.Text = "StatusBar1"
        '
        'frmEdytujTabele
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(763, 366)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.lblSzerokosc)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.txtNazwa)
        Me.Controls.Add(Me.dagKlienci)
        Me.Name = "frmEdytujTabele"
        Me.Text = "Edytuj tabele do wydruku"
        CType(Me.dagKlienci, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dagKlienci As System.Windows.Forms.DataGrid
    Friend WithEvents txtNazwa As System.Windows.Forms.TextBox
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents lblSzerokosc As System.Windows.Forms.Label
    Friend WithEvents stbFiltry As System.Windows.Forms.StatusBar
End Class
