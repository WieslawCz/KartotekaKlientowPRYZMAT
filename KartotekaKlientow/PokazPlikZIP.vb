Public Class PokazPlikZIP
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "

    Dim _ZipFile As String

    Public Sub New(ByVal parZipFile As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _ZipFile = parZipFile
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
    Friend WithEvents ListView As System.Windows.Forms.ListView
    Friend WithEvents Nazwa As System.Windows.Forms.ColumnHeader
    Friend WithEvents WielkoscPrzed As System.Windows.Forms.ColumnHeader
    Friend WithEvents WielkoscPo As System.Windows.Forms.ColumnHeader
    Friend WithEvents Data As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PokazPlikZIP))
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.ListView = New System.Windows.Forms.ListView()
        Me.Nazwa = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Data = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.WielkoscPrzed = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.WielkoscPo = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(602, 240)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 22)
        Me.cmdPowrot.TabIndex = 16
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 240)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 15
        Me.PictureBox2.TabStop = False
        '
        'ListView
        '
        Me.ListView.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListView.BackColor = System.Drawing.SystemColors.Control
        Me.ListView.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Nazwa, Me.Data, Me.WielkoscPrzed, Me.WielkoscPo})
        Me.ListView.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.ListView.ForeColor = System.Drawing.Color.Purple
        Me.ListView.FullRowSelect = True
        Me.ListView.GridLines = True
        Me.ListView.HideSelection = False
        Me.ListView.Location = New System.Drawing.Point(8, 8)
        Me.ListView.Name = "ListView"
        Me.ListView.Size = New System.Drawing.Size(674, 224)
        Me.ListView.TabIndex = 14
        Me.ListView.UseCompatibleStateImageBehavior = False
        Me.ListView.View = System.Windows.Forms.View.Details
        '
        'Nazwa
        '
        Me.Nazwa.Text = "Nazwa pliku"
        Me.Nazwa.Width = 300
        '
        'Data
        '
        Me.Data.Text = "Data"
        Me.Data.Width = 150
        '
        'WielkoscPrzed
        '
        Me.WielkoscPrzed.Text = "Przed kompresj¹"
        Me.WielkoscPrzed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.WielkoscPrzed.Width = 100
        '
        'WielkoscPo
        '
        Me.WielkoscPo.Text = "Po kompresji"
        Me.WielkoscPo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.WielkoscPo.Width = 100
        '
        'PokazPlikZIP
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(690, 270)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.ListView)
        Me.Name = "PokazPlikZIP"
        Me.ShowInTaskbar = False
        Me.Text = "Zawartoœæ pliku ZIP"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub PokazPlikZIP_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        Dim i As Integer = 0
        Dim ExecutingApp As System.Reflection.Assembly
        ExecutingApp = System.Reflection.Assembly.GetExecutingAssembly()
        Dim Name As System.Reflection.AssemblyName
        Name = ExecutingApp.GetName()

        Me.Text = "Zawartœæ pliku " & _ZipFile

        ViewZipFile(_ZipFile, ListView)
        ListView.Select()
    End Sub

    Private Sub PokazPlikZIPe_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub
End Class
