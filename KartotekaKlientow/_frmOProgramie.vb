Public Class _frmOProgramie
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents dagOProgramie As System.Windows.Forms.DataGridView
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.dagOProgramie = New System.Windows.Forms.DataGridView()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dagOProgramie, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(635, 323)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 25)
        Me.cmdPowrot.TabIndex = 16
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.Location = New System.Drawing.Point(8, 323)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(104, 26)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 15
        Me.picLogo.TabStop = False
        '
        'dagOProgramie
        '
        Me.dagOProgramie.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagOProgramie.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dagOProgramie.Location = New System.Drawing.Point(8, 8)
        Me.dagOProgramie.Name = "dagOProgramie"
        Me.dagOProgramie.Size = New System.Drawing.Size(707, 309)
        Me.dagOProgramie.TabIndex = 17
        '
        '_frmOProgramie
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(723, 355)
        Me.Controls.Add(Me.dagOProgramie)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.picLogo)
        Me.Name = "_frmOProgramie"
        Me.ShowInTaskbar = False
        Me.Text = "O programie"
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagOProgramie, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub oProgramie_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.MrJones
        Me.picLogo.Image = My.Resources.logomini
        Me.cmdPowrot.Image = My.Resources._EXIT.ToBitmap
        '-------------------------------
        Dim i As Integer = 0
        Dim ExecutingApp As System.Reflection.Assembly
        ExecutingApp = System.Reflection.Assembly.GetExecutingAssembly()
        Dim Name As System.Reflection.AssemblyName
        Name = ExecutingApp.GetName()


        'Roboczy Dataset
        Dim robDataSet As New System.Data.DataSet()
        Dim robTable As New System.Data.DataTable()
        Dim robDataView As System.Data.DataView = Nothing
        Dim newRow As System.Data.DataRow = Nothing

        Dim robCol1 As DataColumn = robTable.Columns.Add("NAZWAPAR", GetType(System.String))
        Dim robCol2 As DataColumn = robTable.Columns.Add("WARTOSC", GetType(System.String))
        robDataSet.Tables.Add(robTable)
        'definiuj DataView
        robDataView = New DataView(robDataSet.Tables(0))

        '--------------------------------------------
        ' wypelnij dataset informacjami o programie...
        NewRow = robDataSet.Tables(0).NewRow()
        newRow("NAZWAPAR") = "Nazwa aplikacji"
        newRow("WARTOSC") = Name.Name
        robDataSet.Tables(0).Rows.Add(newRow)
        '--------------------------------------------
        newRow = robDataSet.Tables(0).NewRow()
        newRow("NAZWAPAR") = "Opis aplikacji"
        newRow("WARTOSC") = Name.FullName
        robDataSet.Tables(0).Rows.Add(newRow)
        '--------------------------------------------
        newRow = robDataSet.Tables(0).NewRow()
        newRow("NAZWAPAR") = "Katalog aplikacji"
        newRow("WARTOSC") = System.Windows.Forms.Application.StartupPath
        robDataSet.Tables(0).Rows.Add(newRow)
        '--------------------------------------------
        newRow = robDataSet.Tables(0).NewRow()
        newRow("NAZWAPAR") = "Program startowy"
        newRow("WARTOSC") = System.Windows.Forms.Application.ExecutablePath
        robDataSet.Tables(0).Rows.Add(newRow)
        '--------------------------------------------
        newRow = robDataSet.Tables(0).NewRow()
        newRow("NAZWAPAR") = "Kultura"
        newRow("WARTOSC") = Name.CultureInfo.DisplayName
        robDataSet.Tables(0).Rows.Add(newRow)
        '--------------------------------------------
        newRow = robDataSet.Tables(0).NewRow()
        newRow("NAZWAPAR") = "Kod kultury"
        newRow("WARTOSC") = Name.CultureInfo.ToString()
        robDataSet.Tables(0).Rows.Add(newRow)
        '--------------------------------------------
        newRow = robDataSet.Tables(0).NewRow()
        newRow("NAZWAPAR") = "lokalizacja Assembly"
        newRow("WARTOSC") = Name.CodeBase
        robDataSet.Tables(0).Rows.Add(newRow)
        '--------------------------------------------
        newRow = robDataSet.Tables(0).NewRow()
        newRow("NAZWAPAR") = "Nazwa Assembly"
        newRow("WARTOSC") = System.Windows.Forms.Application.ProductName
        robDataSet.Tables(0).Rows.Add(newRow)
        '--------------------------------------------
        newRow = robDataSet.Tables(0).NewRow()
        newRow("NAZWAPAR") = "Wersja Assembly"
        newRow("WARTOSC") = System.Windows.Forms.Application.ProductVersion
        robDataSet.Tables(0).Rows.Add(newRow)
        '--------------------------------------------
        newRow = robDataSet.Tables(0).NewRow()
        newRow("NAZWAPAR") = "Producent"
        newRow("WARTOSC") = System.Windows.Forms.Application.CompanyName
        robDataSet.Tables(0).Rows.Add(newRow)
        '--------------------------------------------
        newRow = robDataSet.Tables(0).NewRow()
        newRow("NAZWAPAR") = "Nazwa produktu"
        newRow("WARTOSC") = System.Windows.Forms.Application.ProductName
        robDataSet.Tables(0).Rows.Add(newRow)
        '--------------------------------------------
        newRow = robDataSet.Tables(0).NewRow()
        newRow("NAZWAPAR") = "Jêzyk"
        newRow("WARTOSC") = System.Windows.Forms.Application.CurrentCulture.ToString()
        robDataSet.Tables(0).Rows.Add(newRow)
        '--------------------------------------------
        newRow = robDataSet.Tables(0).NewRow()
        newRow("NAZWAPAR") = "Kod jêzyka"
        newRow("WARTOSC") = System.Windows.Forms.Application.CurrentCulture.DisplayName
        robDataSet.Tables(0).Rows.Add(newRow)
        '--------------------------------------------

        Try
            '-------------------------------------------------
            'parametry DataGridView
            '-------------------------------------------------
            ' Initialize basic DataGridView properties.
            dagOProgramie.BackgroundColor = System.Drawing.SystemColors.Control
            dagOProgramie.GridColor = System.Drawing.SystemColors.ControlDark
            dagOProgramie.ForeColor = System.Drawing.Color.Purple
            dagOProgramie.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize, _
                                                         System.Drawing.FontStyle.Regular, _
                                                         System.Drawing.GraphicsUnit.Point, _
                                                         CType(238, Byte))
            dagOProgramie.BorderStyle = BorderStyle.Fixed3D
            'dagOProgramie.Dock = DockStyle.Fill
            dagOProgramie.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                    Or System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            ' Set property values appropriate for read-only display and limited interactivity. 
            dagOProgramie.AllowUserToAddRows = False
            dagOProgramie.AllowUserToDeleteRows = False
            dagOProgramie.AllowUserToOrderColumns = True
            dagOProgramie.AllowUserToResizeColumns = True
            dagOProgramie.AllowUserToResizeRows = True
            dagOProgramie.ReadOnly = True
            dagOProgramie.SelectionMode = DataGridViewSelectionMode.CellSelect
            dagOProgramie.MultiSelect = False

            'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
            'dagOProgramie.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            dagOProgramie.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
            dagOProgramie.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            dagOProgramie.ColumnHeadersVisible = True
            dagOProgramie.ColumnHeadersHeight = 40
            dagOProgramie.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
            dagOProgramie.RowHeadersVisible = True
            dagOProgramie.RowHeadersWidth = 50
            dagOProgramie.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing


            ' Set the selection background color for all the cells.
            dagOProgramie.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
            dagOProgramie.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
            dagOProgramie.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
            dagOProgramie.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
            dagOProgramie.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize, _
                                                         System.Drawing.FontStyle.Regular, _
                                                         System.Drawing.GraphicsUnit.Point, _
                                                         CType(238, Byte))
            dagOProgramie.DefaultCellStyle.NullValue = ""
            dagOProgramie.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dagOProgramie.DefaultCellStyle.WrapMode = DataGridViewTriState.True


            ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
            ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
            dagOProgramie.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
            dagOProgramie.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

            ' Set the background color for all rows and for alternating rows.  
            ' The value for alternating rows overrides the value for all rows. 
            dagOProgramie.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
            dagOProgramie.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

            ' Set the row and column header styles.
            dagOProgramie.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
            dagOProgramie.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
            dagOProgramie.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
            dagOProgramie.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
            '-----------------------------------
            'wypelnienie DataGridView
            dagOProgramie.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
            dagOProgramie.DataMember = robDataSet.Tables(0).TableName
            dagOProgramie.DataSource = robDataView
            '-----------------------------------

            ' Add a GridColumnStyle and set the MappingName 
            ' to the name of a DataColumn in the DataTable. 
            ' Set the HeaderText and Width properties. 
            TxtColumnView(dagOProgramie, robDataSet.Tables(0).Columns(0).ColumnName, "Nazwa parametru", 150, DataGridViewContentAlignment.MiddleLeft, True)
            TxtColumnView(dagOProgramie, robDataSet.Tables(0).Columns(1).ColumnName, "Wartoœæ parametru", 500, DataGridViewContentAlignment.MiddleLeft, True)

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        End Try

    End Sub

    Private Sub oProgramie_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub
End Class
