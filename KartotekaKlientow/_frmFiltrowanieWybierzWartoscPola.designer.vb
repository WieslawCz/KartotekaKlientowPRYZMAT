<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class _frmFiltrowanieWybierzWartoscPola
    Inherits System.Windows.Forms.Form

    Dim _NazwaPola As String = ""
    Dim _NazwaTabeli As String = ""
    Dim _Filtr As String = ""
    Dim _ConnString As String = ""
    Dim _DataView As DataView = Nothing

    'wyswietl z podanego (FILTROWANEGO) Dataview
    Public Sub New(ByRef eDataView As DataView,
                   ByVal ePole As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _NazwaPola = ePole
        _DataView = eDataView
        _NazwaTabeli = ""
        _Filtr = ""
        _ConnString = ""
    End Sub

    'dla ACCESS i SQL
    Public Sub New(ByVal ePole As String, _
                   ByVal eTabela As String, _
                   ByVal eFiltr As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _NazwaPola = ePole
        _NazwaTabeli = eTabela
        _Filtr = eFiltr
        _ConnString = ""
        _DataView = Nothing
    End Sub

    'dla innej bazy danych ni¿ ACCESS i SQL (np DBF...)
    Public Sub New(ByVal ePole As String, _
               ByVal eTabela As String, _
               ByVal eFiltr As String, _
               ByVal eConnString As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _NazwaPola = ePole
        _NazwaTabeli = eTabela
        _Filtr = eFiltr
        _ConnString = eConnString
        _DataView = Nothing
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(_frmFiltrowanieWybierzWartoscPola))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.cmdWybierz = New System.Windows.Forms.Button()
        Me.dagWartosc = New System.Windows.Forms.DataGridView()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dagWartosc, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.Image = Global.KartotekaKlientow.My.Resources.Resources.logomini
        Me.picLogo.Location = New System.Drawing.Point(8, 409)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(104, 24)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 54
        Me.picLogo.TabStop = False
        '
        'cmdWybierz
        '
        Me.cmdWybierz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierz.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierz.Image = CType(resources.GetObject("cmdWybierz.Image"), System.Drawing.Image)
        Me.cmdWybierz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWybierz.Location = New System.Drawing.Point(252, 409)
        Me.cmdWybierz.Name = "cmdWybierz"
        Me.cmdWybierz.Size = New System.Drawing.Size(80, 23)
        Me.cmdWybierz.TabIndex = 55
        Me.cmdWybierz.Text = "Wybierz"
        Me.cmdWybierz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dagWartosc
        '
        Me.dagWartosc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagWartosc.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dagWartosc.Location = New System.Drawing.Point(8, 8)
        Me.dagWartosc.Name = "dagWartosc"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagWartosc.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagWartosc.Size = New System.Drawing.Size(330, 395)
        Me.dagWartosc.TabIndex = 53
        '
        '_frmFiltrowanieWybierzWartoscPola
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(344, 437)
        Me.Controls.Add(Me.cmdWybierz)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.dagWartosc)
        Me.Name = "_frmFiltrowanieWybierzWartoscPola"
        Me.Text = "Wybierz Wartoœæ Pola"
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagWartosc, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dagWartosc As System.Windows.Forms.DataGridView
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents cmdWybierz As System.Windows.Forms.Button
End Class
