<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WybierzNrKarty
    Inherits System.Windows.Forms.Form


    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(WybierzNrKarty))
        Me.cmdWybierz = New System.Windows.Forms.Button()
        Me.cbxNrKarty = New System.Windows.Forms.ComboBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdWybierz
        '
        Me.cmdWybierz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierz.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierz.Image = CType(resources.GetObject("cmdWybierz.Image"), System.Drawing.Image)
        Me.cmdWybierz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWybierz.Location = New System.Drawing.Point(115, 269)
        Me.cmdWybierz.Name = "cmdWybierz"
        Me.cmdWybierz.Size = New System.Drawing.Size(80, 23)
        Me.cmdWybierz.TabIndex = 55
        Me.cmdWybierz.Text = "Powrót"
        Me.cmdWybierz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbxNrKarty
        '
        Me.cbxNrKarty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbxNrKarty.DropDownStyle = System.Windows.Forms.ComboBoxStyle.Simple
        Me.cbxNrKarty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxNrKarty.ForeColor = System.Drawing.Color.Navy
        Me.cbxNrKarty.FormattingEnabled = True
        Me.cbxNrKarty.Location = New System.Drawing.Point(8, 12)
        Me.cbxNrKarty.Name = "cbxNrKarty"
        Me.cbxNrKarty.Size = New System.Drawing.Size(187, 251)
        Me.cbxNrKarty.TabIndex = 58
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 269)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 59
        Me.PictureBox2.TabStop = False
        '
        'WybierzNrKarty
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(207, 300)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdWybierz)
        Me.Controls.Add(Me.cbxNrKarty)
        Me.Name = "WybierzNrKarty"
        Me.ShowInTaskbar = False
        Me.Text = "Wybierz Nr Karty"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdWybierz As System.Windows.Forms.Button
    Friend WithEvents cbxNrKarty As System.Windows.Forms.ComboBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
End Class
