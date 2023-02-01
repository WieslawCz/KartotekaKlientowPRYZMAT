<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAnalizaCzestKontaktow
    Inherits System.Windows.Forms.Form

    Dim _DataView As DataView = Nothing
    Dim _IdentKlienta As String = ""
    Public Sub New(ByRef pDataView As DataView, ByVal pIdentK As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _DataView = pDataView
        _IdentKlienta = pIdentK
    End Sub




    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAnalizaCzestKontaktow))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblIDDomu = New System.Windows.Forms.Label()
        Me.lblNrDomu = New System.Windows.Forms.Label()
        Me.lblUlica = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblMiejscowosc = New System.Windows.Forms.Label()
        Me.lblKod = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa = New System.Windows.Forms.Label()
        Me.lblIdent = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.dagRaport = New System.Windows.Forms.DataGrid()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dagRaport, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lblIDDomu)
        Me.Panel1.Controls.Add(Me.lblNrDomu)
        Me.Panel1.Controls.Add(Me.lblUlica)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.lblMiejscowosc)
        Me.Panel1.Controls.Add(Me.lblKod)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.lblNazwaHandlowa)
        Me.Panel1.Controls.Add(Me.lblIdent)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(877, 71)
        Me.Panel1.TabIndex = 59
        '
        'lblIDDomu
        '
        Me.lblIDDomu.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblIDDomu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIDDomu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIDDomu.ForeColor = System.Drawing.Color.Purple
        Me.lblIDDomu.Location = New System.Drawing.Point(821, 41)
        Me.lblIDDomu.Name = "lblIDDomu"
        Me.lblIDDomu.Size = New System.Drawing.Size(45, 16)
        Me.lblIDDomu.TabIndex = 36
        '
        'lblNrDomu
        '
        Me.lblNrDomu.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNrDomu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNrDomu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNrDomu.ForeColor = System.Drawing.Color.Purple
        Me.lblNrDomu.Location = New System.Drawing.Point(778, 41)
        Me.lblNrDomu.Name = "lblNrDomu"
        Me.lblNrDomu.Size = New System.Drawing.Size(45, 16)
        Me.lblNrDomu.TabIndex = 11
        '
        'lblUlica
        '
        Me.lblUlica.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUlica.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblUlica.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblUlica.ForeColor = System.Drawing.Color.Purple
        Me.lblUlica.Location = New System.Drawing.Point(237, 41)
        Me.lblUlica.Name = "lblUlica"
        Me.lblUlica.Size = New System.Drawing.Size(535, 16)
        Me.lblUlica.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(181, 42)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 16)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Ulica . . "
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(181, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 16)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Nazwa . . . . . . . . "
        '
        'lblMiejscowosc
        '
        Me.lblMiejscowosc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblMiejscowosc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblMiejscowosc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblMiejscowosc.ForeColor = System.Drawing.Color.Purple
        Me.lblMiejscowosc.Location = New System.Drawing.Point(309, 25)
        Me.lblMiejscowosc.Name = "lblMiejscowosc"
        Me.lblMiejscowosc.Size = New System.Drawing.Size(557, 16)
        Me.lblMiejscowosc.TabIndex = 12
        '
        'lblKod
        '
        Me.lblKod.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKod.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblKod.ForeColor = System.Drawing.Color.Purple
        Me.lblKod.Location = New System.Drawing.Point(237, 25)
        Me.lblKod.Name = "lblKod"
        Me.lblKod.Size = New System.Drawing.Size(64, 16)
        Me.lblKod.TabIndex = 7
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(181, 26)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(56, 16)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Adres"
        '
        'lblNazwaHandlowa
        '
        Me.lblNazwaHandlowa.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwaHandlowa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwaHandlowa.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwaHandlowa.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwaHandlowa.Location = New System.Drawing.Point(237, 8)
        Me.lblNazwaHandlowa.Name = "lblNazwaHandlowa"
        Me.lblNazwaHandlowa.Size = New System.Drawing.Size(629, 16)
        Me.lblNazwaHandlowa.TabIndex = 2
        '
        'lblIdent
        '
        Me.lblIdent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdent.ForeColor = System.Drawing.Color.Purple
        Me.lblIdent.Location = New System.Drawing.Point(64, 8)
        Me.lblIdent.Name = "lblIdent"
        Me.lblIdent.Size = New System.Drawing.Size(80, 24)
        Me.lblIdent.TabIndex = 1
        Me.lblIdent.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Klient . . . . . . . . "
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 418)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 65
        Me.PictureBox1.TabStop = False
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ForeColor = System.Drawing.Color.Black
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(797, 419)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 23)
        Me.cmdStop.TabIndex = 66
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dagRaport
        '
        Me.dagRaport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagRaport.DataMember = ""
        Me.dagRaport.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dagRaport.Location = New System.Drawing.Point(8, 85)
        Me.dagRaport.Name = "dagRaport"
        Me.dagRaport.Size = New System.Drawing.Size(880, 327)
        Me.dagRaport.TabIndex = 67
        '
        'frmAnalizaCzestKontaktow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(897, 449)
        Me.Controls.Add(Me.dagRaport)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmAnalizaCzestKontaktow"
        Me.Text = "Analiza częstości kontaktów z klientem "
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagRaport, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblIDDomu As System.Windows.Forms.Label
    Friend WithEvents lblNrDomu As System.Windows.Forms.Label
    Friend WithEvents lblUlica As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblMiejscowosc As System.Windows.Forms.Label
    Friend WithEvents lblKod As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblNazwaHandlowa As System.Windows.Forms.Label
    Friend WithEvents lblIdent As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents dagRaport As System.Windows.Forms.DataGrid
End Class
