Imports System.Xml
Imports System.IO
Imports System.Text

Public Class PokazPlikXML
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
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents dagEXCEL As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PokazPlikXML))
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.dagEXCEL = New System.Windows.Forms.DataGrid()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dagEXCEL, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ForeColor = System.Drawing.Color.Black
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(456, 302)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 24)
        Me.cmdStop.TabIndex = 44
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 302)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 25)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 43
        Me.PictureBox1.TabStop = False
        '
        'dagEXCEL
        '
        Me.dagEXCEL.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagEXCEL.DataMember = ""
        Me.dagEXCEL.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dagEXCEL.Location = New System.Drawing.Point(8, 8)
        Me.dagEXCEL.Name = "dagEXCEL"
        Me.dagEXCEL.Size = New System.Drawing.Size(536, 289)
        Me.dagEXCEL.TabIndex = 42
        '
        'PokazPlikXML
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(552, 334)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.dagEXCEL)
        Me.Name = "PokazPlikXML"
        Me.Text = "PokazPlikXML"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagEXCEL, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub PokazPlikXML_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        Dim EXCELds As New DataSet
        Dim EXCELdv As New DataView

        'wypelnij DATASET
        Try
            Me.Update()

            Dim fs As New FileStream(oPlikNaDysku, FileMode.Open)
            Dim xr As New XmlTextReader(fs)
            'pobierz dane do DataSet
            EXCELds.ReadXml(xr)
            xr.Close()
            fs.Close()

        Catch ex As System.IO.FileNotFoundException
            MessageBox.Show("Nie znalaz³em pliku '" + oPlikNaDysku + "'. Proszê wybraæ prawid³ow¹ nazwê pliku XML.")

        Catch ex As XmlException
        End Try


        'parametry DataGrid
        dagEXCEL.CaptionVisible = False
        dagEXCEL.CaptionText = EXCELds.Tables(0).TableName
        dagEXCEL.ColumnHeadersVisible = True
        dagEXCEL.RowHeadersVisible = True
        dagEXCEL.BackgroundColor = System.Drawing.SystemColors.Control
        dagEXCEL.BorderStyle = BorderStyle.Fixed3D
        dagEXCEL.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                Or System.Windows.Forms.AnchorStyles.Bottom) _
                                Or System.Windows.Forms.AnchorStyles.Left) _
                                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        dagEXCEL.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        dagEXCEL.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))

        'definiuj DataView
        EXCELdv = New DataView(EXCELds.Tables(0))
        'wypelnienie DataGrid
        'dagKlienci.SetDataBinding(DataSetKlienci, "TABELA_Klienci")
        dagEXCEL.DataSource = EXCELdv
        'dagEXCEL.DataSource = EXCELDs.Tables("Arkusz1")

        ' Add a GridTableStyle and set the MappingName to the name of the DataTable.
        Dim EXCELts As New DataGridTableStyle
        EXCELts.MappingName = EXCELds.Tables(0).TableName
        EXCELts.AlternatingBackColor = System.Drawing.SystemColors.Control
        EXCELts.BackColor = System.Drawing.SystemColors.ControlLight
        EXCELts.GridLineColor = System.Drawing.SystemColors.ControlDark
        EXCELts.ForeColor = System.Drawing.Color.Purple
        EXCELts.HeaderBackColor = System.Drawing.SystemColors.ScrollBar
        EXCELts.HeaderForeColor = System.Drawing.Color.Navy
        EXCELts.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
        EXCELts.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow

        Dim daCol As DataColumn
        Dim liCol As Integer = 0
        For Each daCol In EXCELds.Tables(0).Columns
            liCol += 1
            Dim TCol As New Softart.myDataGridTextBoxColumn
            TCol.MappingName = daCol.ColumnName
            TCol.HeaderText = daCol.ColumnName
            TCol.Width = 100
            TCol.Alignment = HorizontalAlignment.Left
            EXCELts.GridColumnStyles.Add(TCol)
        Next

        ' Add the DataGridTableStyle instance to the GridTableStylesCollection. 
        dagEXCEL.TableStyles.Add(EXCELts)
        'ustaw sie na pierwszym zapisie
        If EXCELds.Tables(0).Rows.Count > 0 Then
            dagEXCEL.Select(0)
        End If
    End Sub

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub

    Private Sub PokazPlikXML_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub
End Class
