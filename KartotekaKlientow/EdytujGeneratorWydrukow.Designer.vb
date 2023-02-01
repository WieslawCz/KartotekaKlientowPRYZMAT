<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EdytujGeneratorWydrukow
    Inherits System.Windows.Forms.Form

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EdytujGeneratorWydrukow))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cbxWyrownanie = New System.Windows.Forms.ComboBox()
        Me.txtMaxIlWierszy = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtNaglowek = New System.Windows.Forms.TextBox()
        Me.txtRozmiar = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblTyp = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblNazwa = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
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
        Me.Panel1.Controls.Add(Me.cbxWyrownanie)
        Me.Panel1.Controls.Add(Me.txtMaxIlWierszy)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.txtNaglowek)
        Me.Panel1.Controls.Add(Me.txtRozmiar)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.lblTyp)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.lblNazwa)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(337, 169)
        Me.Panel1.TabIndex = 0
        '
        'cbxWyrownanie
        '
        Me.cbxWyrownanie.FormattingEnabled = True
        Me.cbxWyrownanie.Location = New System.Drawing.Point(93, 108)
        Me.cbxWyrownanie.Name = "cbxWyrownanie"
        Me.cbxWyrownanie.Size = New System.Drawing.Size(226, 21)
        Me.cbxWyrownanie.TabIndex = 77
        '
        'txtMaxIlWierszy
        '
        Me.txtMaxIlWierszy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMaxIlWierszy.ForeColor = System.Drawing.Color.Purple
        Me.txtMaxIlWierszy.Location = New System.Drawing.Point(93, 132)
        Me.txtMaxIlWierszy.Name = "txtMaxIlWierszy"
        Me.txtMaxIlWierszy.Size = New System.Drawing.Size(226, 20)
        Me.txtMaxIlWierszy.TabIndex = 76
        Me.txtMaxIlWierszy.Text = "1"
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(12, 135)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(130, 16)
        Me.Label4.TabIndex = 75
        Me.Label4.Text = "Max il.wierszy . . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(12, 111)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(130, 16)
        Me.Label2.TabIndex = 73
        Me.Label2.Text = "Wyrównanie . . . . . . . . ."
        '
        'txtNaglowek
        '
        Me.txtNaglowek.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNaglowek.ForeColor = System.Drawing.Color.Purple
        Me.txtNaglowek.Location = New System.Drawing.Point(93, 84)
        Me.txtNaglowek.Name = "txtNaglowek"
        Me.txtNaglowek.Size = New System.Drawing.Size(226, 20)
        Me.txtNaglowek.TabIndex = 72
        '
        'txtRozmiar
        '
        Me.txtRozmiar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRozmiar.ForeColor = System.Drawing.Color.Purple
        Me.txtRozmiar.Location = New System.Drawing.Point(93, 60)
        Me.txtRozmiar.Name = "txtRozmiar"
        Me.txtRozmiar.Size = New System.Drawing.Size(226, 20)
        Me.txtRozmiar.TabIndex = 65
        Me.txtRozmiar.Text = "0"
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(12, 87)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(130, 16)
        Me.Label7.TabIndex = 71
        Me.Label7.Text = "Nag³ówek . . . . . . . . ."
        '
        'Label5
        '
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(12, 63)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(130, 16)
        Me.Label5.TabIndex = 69
        Me.Label5.Text = "Rozmiar [mm] . . . . . . . . ."
        '
        'lblTyp
        '
        Me.lblTyp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTyp.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTyp.ForeColor = System.Drawing.Color.Purple
        Me.lblTyp.Location = New System.Drawing.Point(93, 36)
        Me.lblTyp.Name = "lblTyp"
        Me.lblTyp.Size = New System.Drawing.Size(226, 16)
        Me.lblTyp.TabIndex = 68
        Me.lblTyp.Text = "0"
        Me.lblTyp.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(12, 37)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(130, 16)
        Me.Label3.TabIndex = 67
        Me.Label3.Text = "Typ pola . . . . . . . . ."
        '
        'lblNazwa
        '
        Me.lblNazwa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa.Location = New System.Drawing.Point(93, 12)
        Me.lblNazwa.Name = "lblNazwa"
        Me.lblNazwa.Size = New System.Drawing.Size(226, 16)
        Me.lblNazwa.TabIndex = 66
        Me.lblNazwa.Text = "0"
        Me.lblNazwa.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(12, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(130, 16)
        Me.Label1.TabIndex = 64
        Me.Label1.Text = "Nazwa pola . . . . . . . . ."
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(257, 183)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 23)
        Me.cmdStop.TabIndex = 58
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 183)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 59
        Me.PictureBox1.TabStop = False
        '
        'EdytujGeneratorWydrukow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(350, 218)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "EdytujGeneratorWydrukow"
        Me.ShowInTaskbar = False
        Me.Text = "Edytuj pozycje generatora wydruków"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtRozmiar As System.Windows.Forms.TextBox
    Friend WithEvents txtNaglowek As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblTyp As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblNazwa As System.Windows.Forms.Label
    Friend WithEvents cbxWyrownanie As System.Windows.Forms.ComboBox
    Friend WithEvents txtMaxIlWierszy As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
