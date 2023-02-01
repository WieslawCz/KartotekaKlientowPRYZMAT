Imports System.Drawing.Printing

Public Class KatalogKartPayBack
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
    Friend WithEvents cmdEdytuj As System.Windows.Forms.Button
    Friend WithEvents cmdDodaj As System.Windows.Forms.Button
    Friend WithEvents stbKartyPayBack As System.Windows.Forms.StatusBar
    Friend WithEvents menuKartyPayBackEdytuj As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents MenuUEEdytuj As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuUEDodaj As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuUEUsun As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents cmdFi As Button
    Friend WithEvents cmdWszystko As Button
    Friend WithEvents stbFiltry As StatusBar
    Friend WithEvents dagKartyPayBack As DataGridView
    Friend WithEvents HorizScrollTimer As Timer
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogKartPayBack))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.stbKartyPayBack = New System.Windows.Forms.StatusBar()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.menuKartyPayBackEdytuj = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.MenuUEEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuUEDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuUEUsun = New System.Windows.Forms.ToolStripMenuItem()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.dagKartyPayBack = New System.Windows.Forms.DataGridView()
        Me.HorizScrollTimer = New System.Windows.Forms.Timer(Me.components)
        Me.menuKartyPayBackEdytuj.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dagKartyPayBack, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(309, 302)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(88, 24)
        Me.cmdStop.TabIndex = 37
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'stbKartyPayBack
        '
        Me.stbKartyPayBack.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbKartyPayBack.Dock = System.Windows.Forms.DockStyle.None
        Me.stbKartyPayBack.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbKartyPayBack.Location = New System.Drawing.Point(8, 303)
        Me.stbKartyPayBack.Name = "stbKartyPayBack"
        Me.stbKartyPayBack.ShowPanels = True
        Me.stbKartyPayBack.Size = New System.Drawing.Size(295, 23)
        Me.stbKartyPayBack.TabIndex = 42
        Me.stbKartyPayBack.Text = "StatusBar1"
        '
        'cmdEdytuj
        '
        Me.cmdEdytuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytuj.Image = CType(resources.GetObject("cmdEdytuj.Image"), System.Drawing.Image)
        Me.cmdEdytuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytuj.Location = New System.Drawing.Point(309, 106)
        Me.cmdEdytuj.Name = "cmdEdytuj"
        Me.cmdEdytuj.Size = New System.Drawing.Size(88, 23)
        Me.cmdEdytuj.TabIndex = 41
        Me.cmdEdytuj.Text = "Edytuj"
        Me.cmdEdytuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdUsuñ
        '
        Me.cmdUsuñ.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUsuñ.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdUsuñ.Image = CType(resources.GetObject("cmdUsuñ.Image"), System.Drawing.Image)
        Me.cmdUsuñ.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUsuñ.Location = New System.Drawing.Point(309, 77)
        Me.cmdUsuñ.Name = "cmdUsuñ"
        Me.cmdUsuñ.Size = New System.Drawing.Size(88, 23)
        Me.cmdUsuñ.TabIndex = 40
        Me.cmdUsuñ.Text = "Usuñ"
        Me.cmdUsuñ.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdDodaj
        '
        Me.cmdDodaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodaj.Image = CType(resources.GetObject("cmdDodaj.Image"), System.Drawing.Image)
        Me.cmdDodaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodaj.Location = New System.Drawing.Point(309, 48)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(88, 23)
        Me.cmdDodaj.TabIndex = 39
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'menuKartyPayBackEdytuj
        '
        Me.menuKartyPayBackEdytuj.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuUEEdytuj, Me.MenuUEDodaj, Me.MenuUEUsun})
        Me.menuKartyPayBackEdytuj.Name = "KartyPayBackEdytuj"
        Me.menuKartyPayBackEdytuj.Size = New System.Drawing.Size(108, 70)
        '
        'MenuUEEdytuj
        '
        Me.MenuUEEdytuj.Name = "MenuUEEdytuj"
        Me.MenuUEEdytuj.Size = New System.Drawing.Size(107, 22)
        Me.MenuUEEdytuj.Text = "Edytuj"
        '
        'MenuUEDodaj
        '
        Me.MenuUEDodaj.Name = "MenuUEDodaj"
        Me.MenuUEDodaj.Size = New System.Drawing.Size(107, 22)
        Me.MenuUEDodaj.Text = "Dodaj"
        '
        'MenuUEUsun
        '
        Me.MenuUEUsun.Name = "MenuUEUsun"
        Me.MenuUEUsun.Size = New System.Drawing.Size(107, 22)
        Me.MenuUEUsun.Text = "Usun"
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(309, 8)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(88, 34)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 45
        Me.PictureBox2.TabStop = False
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(66, 8)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 183
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(88, 7)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 23)
        Me.cmdWszystko.TabIndex = 181
        '
        'stbFiltry
        '
        Me.stbFiltry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltry.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltry.Location = New System.Drawing.Point(8, 8)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(295, 22)
        Me.stbFiltry.TabIndex = 182
        Me.stbFiltry.Text = "StatusBar1"
        '
        'dagKartyPayBack
        '
        Me.dagKartyPayBack.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagKartyPayBack.ContextMenuStrip = Me.menuKartyPayBackEdytuj
        Me.dagKartyPayBack.Location = New System.Drawing.Point(8, 30)
        Me.dagKartyPayBack.Name = "dagKartyPayBack"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagKartyPayBack.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagKartyPayBack.Size = New System.Drawing.Size(295, 272)
        Me.dagKartyPayBack.TabIndex = 180
        '
        'HorizScrollTimer
        '
        '
        'KatalogKartPayBack
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(407, 333)
        Me.Controls.Add(Me.cmdFi)
        Me.Controls.Add(Me.cmdWszystko)
        Me.Controls.Add(Me.stbFiltry)
        Me.Controls.Add(Me.dagKartyPayBack)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.stbKartyPayBack)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Name = "KatalogKartPayBack"
        Me.ShowInTaskbar = False
        Me.Text = "Karty PayBack"
        Me.menuKartyPayBackEdytuj.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagKartyPayBack, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim StartFormy As Boolean = True
    Dim _OCoMamRobic As String = oCoMamRobic

    '---------------------------------------------
    'SCROLL-parametry obslugi scrollingu poziomego
    Dim StartPointInCell00 As System.Drawing.Point             'pocz¹tek ektanu przed scrollingiem
    Dim ScrollXOffset As Integer = 0            'bie¿¹cy punkt ekranu wskazywany przez kursor podczas scrollingu
    Dim HorizScrollLastTime As String = ""      'Chwila rozpoczecia scrollingu - filtry wyswietlaner bêd¹ po 1 sekundzie...
    ' HorizScrollTimer - zegar odliczaj¹cy czas od poczatku scrollingu
    '---------------------------------------------
    Dim DataSetKartyPayBack As New DataSet()
    Dim TableKartyPayBack As New DataTable()
    Dim DataViewKartyPayBack As DataView = Nothing
    Dim newRow As DataRow = Nothing
    Dim foundRow As DataRow = Nothing
    Dim j As Integer = 0

    Private Sub KatalogKartPayBack_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        'Roboczy Dataset
        Dim robCol1 As DataColumn = TableKartyPayBack.Columns.Add("NRKARTYPAYBACK", GetType(System.String))
        DataSetKartyPayBack.Tables.Add(TableKartyPayBack)
        'definiuj DataView
        DataViewKartyPayBack = New DataView(DataSetKartyPayBack.Tables(0))

        'zdefiniuj unikalny klucz wyszukiwania w tabeli
        DataSetKartyPayBack.Tables(0).PrimaryKey = New DataColumn() {DataSetKartyPayBack.Tables(0).Columns("NRKARTYPAYBACK")}

        Dim txt As String = Trim(oNryKartPaybackKli)
        Dim NrKarty As String = ""
        Dim poz As Integer = 0
        poz = InStr(txt, ";")
        Do While poz > 0
            NrKarty = Mid(txt, 1, poz - 1)
            txt = Mid(txt, poz + 1)
            If Len(Trim(NrKarty)) > 0 Then
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = NrKarty
                foundRow = DataSetKartyPayBack.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf & _
                    '    "Ident.u¿ytkownika = " & oIdentUzytkownika, _
                    '    "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                Else
                    'przychod - dopisujemy 
                    newRow = DataSetKartyPayBack.Tables(0).NewRow()
                    newRow("NRKARTYPAYBACK") = NrKarty
                    DataSetKartyPayBack.Tables(0).Rows.Add(newRow)
                End If
            End If
            poz = InStr(txt, ";")
        Loop
        'ostatni nr karty
        NrKarty = txt
        txt = ""
        If Len(Trim(NrKarty)) > 0 Then
            ' Create an array for the key values to find.
            Dim findTheseVals(0) As Object
            findTheseVals(0) = NrKarty
            foundRow = DataSetKartyPayBack.Tables(0).Rows.Find(findTheseVals)
            ' sprawdz czy jest...
            If Not (foundRow Is Nothing) Then
                'MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf & _
                '    "Ident.u¿ytkownika = " & oIdentUzytkownika, _
                '    "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Else
                'przychod - dopisujemy 
                newRow = DataSetKartyPayBack.Tables(0).NewRow()
                newRow("NRKARTYPAYBACK") = NrKarty
                DataSetKartyPayBack.Tables(0).Rows.Add(newRow)
            End If
        End If

        '--------------------------------
        Try        'parametry DataGrid
            '-------------------------------------------------
            'parametry DataGridView
            '-------------------------------------------------
            ' Initialize basic DataGridView properties.
            dagKartyPayBack.BackgroundColor = System.Drawing.SystemColors.Control
            dagKartyPayBack.GridColor = System.Drawing.SystemColors.ControlDark
            dagKartyPayBack.ForeColor = System.Drawing.Color.Purple
            dagKartyPayBack.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
            dagKartyPayBack.BorderStyle = BorderStyle.Fixed3D
            'dagKartyPayBack.Dock = DockStyle.Fill
            dagKartyPayBack.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            ' Set property values appropriate for read-only display and limited interactivity. 
            dagKartyPayBack.AllowUserToAddRows = False
            dagKartyPayBack.AllowUserToDeleteRows = False
            dagKartyPayBack.AllowUserToOrderColumns = True
            dagKartyPayBack.AllowUserToResizeColumns = True
            dagKartyPayBack.AllowUserToResizeRows = True
            dagKartyPayBack.ReadOnly = True
            dagKartyPayBack.SelectionMode = DataGridViewSelectionMode.CellSelect
            dagKartyPayBack.MultiSelect = False

            'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
            'dagKartyPayBack.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            dagKartyPayBack.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
            dagKartyPayBack.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            dagKartyPayBack.ColumnHeadersVisible = True
            dagKartyPayBack.ColumnHeadersHeight = 40
            dagKartyPayBack.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
            dagKartyPayBack.RowHeadersVisible = True
            dagKartyPayBack.RowHeadersWidth = 50
            dagKartyPayBack.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

            ' Set the selection background color for all the cells.
            dagKartyPayBack.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
            dagKartyPayBack.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
            dagKartyPayBack.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
            dagKartyPayBack.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
            dagKartyPayBack.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
            dagKartyPayBack.DefaultCellStyle.NullValue = ""
            dagKartyPayBack.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dagKartyPayBack.DefaultCellStyle.WrapMode = DataGridViewTriState.True


            ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
            ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
            dagKartyPayBack.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
            dagKartyPayBack.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

            ' Set the background color for all rows and for alternating rows.  
            ' The value for alternating rows overrides the value for all rows. 
            dagKartyPayBack.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
            dagKartyPayBack.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

            ' Set the row and column header styles.
            dagKartyPayBack.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
            dagKartyPayBack.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
            dagKartyPayBack.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
            dagKartyPayBack.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
            '-----------------------------------
            'wypelnienie DataGridView
            dagKartyPayBack.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
            dagKartyPayBack.DataMember = DataSetKartyPayBack.Tables(0).TableName
            'dagKartyPayBack.DataSource = DataViewKartyPayBack
            '-----------------------------------
            TxtColumnView(dagKartyPayBack, DataSetKartyPayBack.Tables(0).Columns(0).ColumnName, "Nr Karty PayBack", 200, HorizontalAlignment.Left, True)

            Me.Cursor = Cursors.WaitCursor
            dagKartyPayBack.DataSource = DataViewKartyPayBack
            Me.Cursor = Cursors.Default

            'ustaw sie na pierwszym zapisie
            If DataSetKartyPayBack.Tables(0).Rows.Count > 0 Then
                dagKartyPayBack.CurrentCell = dagKartyPayBack(0, 0)
                dagKartyPayBack.CurrentCell.Selected = True
            End If


        Catch Ex As System.Exception
            ViewInLog(Ex, Me.Name(), Nothing)
            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    MessageBoxIcon.Information, _
            '    MessageBoxDefaultButton.Button1)
        End Try

        '------------------------------------
        'dodaj StatusBar na koncu
        stbKartyPayBack.BackColor = System.Drawing.SystemColors.Control
        stbKartyPayBack.ForeColor = System.Drawing.Color.DarkGreen
        stbKartyPayBack.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        stbKartyPayBack.Panels.Add("stbPanelIleRec")
        stbKartyPayBack.Panels(0).AutoSize = StatusBarPanelAutoSize.None
        stbKartyPayBack.Panels(0).Width = 80
        stbKartyPayBack.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbKartyPayBack.Panels(0).Text = "Iloœæ rec.: " & DataSetKartyPayBack.Tables(0).Rows.Count.ToString

        stbKartyPayBack.Panels.Add("stbPanelWiersz")
        stbKartyPayBack.Panels(1).AutoSize = StatusBarPanelAutoSize.None
        stbKartyPayBack.Panels(1).Width = 80
        stbKartyPayBack.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
        If dagKartyPayBack.CurrentCell Is Nothing Then
            stbKartyPayBack.Panels(1).Text = "Wiersz : "
        Else
            stbKartyPayBack.Panels(1).Text = "Wiersz : " & dagKartyPayBack.CurrentCell.RowIndex.ToString
        End If

        stbKartyPayBack.Panels.Add("stbPanelKolumna")
        stbKartyPayBack.Panels(2).AutoSize = StatusBarPanelAutoSize.None
        stbKartyPayBack.Panels(2).Width = 80
        stbKartyPayBack.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
        If dagKartyPayBack.CurrentCell Is Nothing Then
            stbKartyPayBack.Panels(2).Text = "Kolumna : "
        Else
            stbKartyPayBack.Panels(2).Text = "Kolumna : " & dagKartyPayBack.CurrentCell.ColumnIndex.ToString
        End If

        stbKartyPayBack.Panels.Add("stbPanelSort")
        stbKartyPayBack.Panels(3).AutoSize = StatusBarPanelAutoSize.None
        stbKartyPayBack.Panels(3).Width = 150
        stbKartyPayBack.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbKartyPayBack.Panels(3).Text = "Sort : "

        stbKartyPayBack.Panels.Add("stbPanelFiltr")
        stbKartyPayBack.Panels(4).AutoSize = StatusBarPanelAutoSize.None
        stbKartyPayBack.Panels(4).Width = 150
        stbKartyPayBack.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbKartyPayBack.Panels(4).Text = "Filtr : "

        stbKartyPayBack.Panels.Add("stbPanelReszta")
        stbKartyPayBack.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
        'stbKartyPayBack.Panels(5).Width = 150
        stbKartyPayBack.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbKartyPayBack.Panels(5).Text = ""

        stbKartyPayBack.ShowPanels = True
        '---------------------------------
        'dodaj StatusBar na koncu
        stbFiltry.BackColor = System.Drawing.SystemColors.Control
        stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
        stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

        stbFiltry.Panels.Add("stbFiltryRowHead")
        stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
        stbFiltry.Panels(0).Width = 50  'dagKartyPayBack.TableStyles(0).RowHeaderWidth
        stbFiltry.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbFiltry.Panels(0).Text = "Filtry"

        stbFiltry.Panels.Add("stbFiltryReszta")
        stbFiltry.Panels(1).AutoSize = StatusBarPanelAutoSize.Spring
        'stbFiltry.Panels(1).Width = 150
        stbFiltry.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
        stbFiltry.Panels(1).Text = ""

        stbFiltry.ShowPanels = True

        'dodaj ctrl do pobierania szablonu
        'cmdWszystko.Size = New Size(20, 20)
        'cmdWszystko.Location = New system.drawing.Point(stbFiltry.Location.X + _
        '                                 stbFiltry.Size.Width - _
        '                                 cmdWszystko.Size.Width, _
        '                                 stbFiltry.Location.Y + 2)
        'cmdWszystko.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or _
        '                            System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

        cmdWszystko.Size = New Size(20, 22)
        cmdWszystko.Location = New System.Drawing.Point(stbFiltry.Location.X + 50 - 20,
                                             stbFiltry.Location.Y)
        cmdWszystko.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)

        '---------------------------------
        InitInfoUzytkownika() 'inicjuj zmienne
        '-----------------------------------------------------------------
        'set the header width to something apporpriate
        dagKartyPayBack.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagKartyPayBack.RowHeaderWidth = 80 'use if no tablestyle
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetKartyPayBack.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagKartyPayBack.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagKartyPayBack.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagKartyPayBack.FirstDisplayedScrollingColumnIndex +
                        dagKartyPayBack.GetCellDisplayRectangle(dagKartyPayBack.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagKartyPayBack.Left + 4), (dagKartyPayBack.Top + 4))
            ScrollXOffset = 0
        End If
        '--------------------------------------------------
        'inicjowanie zegara dla scrollingu poziomego
        HorizScrollLastTime = ""
        HorizScrollTimer.Interval = 2 * 1000
        HorizScrollTimer.Enabled = True
        '--------------------------------------------------
        'Stworz zmienne do filtrowania
        Dim ii As Integer = 0
        For ii = 0 To DataViewKartyPayBack.Table.Columns.Count - 1
            'stworz zmienn¹ TXTBOX
            Dim CtrlT As New System.Windows.Forms.TextBox
            CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
            CtrlT.Visible = False
            Me.Controls.Add(CtrlT)
            AddHandler CtrlT.TextChanged, New EventHandler(AddressOf txtFiXX_TextChanged)
            AddHandler CtrlT.GotFocus, New EventHandler(AddressOf txtFiXX_GotFocus)
            AddHandler CtrlT.LostFocus, New EventHandler(AddressOf txtFiXX_LostFocus)
            AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

            'stworz zmienn¹ BUTTON
            Dim CtrlB As New System.Windows.Forms.Button
            CtrlB.Name = "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
            CtrlB.Image = cmdFi.Image
            CtrlB.ImageAlign = ContentAlignment.MiddleCenter
            CtrlB.Visible = False
            Me.Controls.Add(CtrlB)
            AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXX_Click)
            AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXX_GotFocus)
            AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXX_LostFocus)
        Next
        Me.Refresh()
        '--------------------------------------------------
        StartFormy = False
        dagKartyPayBack_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me, dagKartyPayBack, DataViewKartyPayBack.Table.Columns.Count, stbFiltry, StartFormy)


        InitKartyPayBack() 'inicjuj zmienne
        '-----------------------------------------------------------------
        dagKartyPayBack.ContextMenuStrip = Me.menuKartyPayBackEdytuj
        '--------------------------------------------------------------------
    End Sub


    Private Sub KatalogKartPayBack_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub







    Private Sub dagKartyPayBack_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKartyPayBack.CurrentCellChanged
        If Not StartFormy Then
            If dagKartyPayBack.CurrentCell Is Nothing Then
                stbKartyPayBack.Panels(1).Text = "Wi: "
                stbKartyPayBack.Panels(2).Text = "Ko: "
            Else
                stbKartyPayBack.Panels(1).Text = "Wi: " & dagKartyPayBack.CurrentCell.RowIndex.ToString
                stbKartyPayBack.Panels(2).Text = "Ko: " & dagKartyPayBack.CurrentCell.ColumnIndex.ToString
            End If
        End If
    End Sub





    '-=------------------------------------------
    'obsluga zegara scrollingu poziomego
    '-=------------------------------------------
    Private Sub HorizScrollTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HorizScrollTimer.Tick

    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagKartyPayBack.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagKartyPayBack.Rows(nextrow).Selected = True
            End If
        End If
    End Sub



    '*********************************************************
    '***** Filtrowanie wg poszczególnych pól
    '*********************************************************

    'AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

    Private Sub txtFiXX_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Not StartFormy Then
            Select Case e.KeyCode
                'Case Keys.Enter
                'Case Keys.Insert, Keys.Add
                'Case Keys.Delete, Keys.Subtract

                Case Keys.Down
                    PrzejdzDoDGV(dagKartyPayBack, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me, dagKartyPayBack, DataViewKartyPayBack.Table.Columns.Count, DataViewKartyPayBack, stbKartyPayBack)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me, dagKartyPayBack, Mid(sender.name, 6, 3), "KartyPayBack")
    End Sub
    Private Sub cmdFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiFiltraCmd(Me, sender)
    End Sub
    Private Sub cmdFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraCmd(Me, sender)
    End Sub


    '-------------------------------------------------
    ' filtrowanie 
    '-------------------------------------------------

    Private Sub cmdWszystko_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWszystko.Click
        Me.Cursor = Cursors.WaitCursor
        oLokalnyFiltr = ""
        Try
            StartFormy = True
            CzyscFiltryDGV(Me, dagKartyPayBack, DataViewKartyPayBack.Table.Columns.Count, DataViewKartyPayBack, stbKartyPayBack)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbKartyPayBack.Panels(0).Text = "Il.zap.: " & DataViewKartyPayBack.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagKartyPayBack_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagKartyPayBack.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagKartyPayBack.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagKartyPayBack.DefaultCellStyle.Font,
                      b,
                      e.RowBounds.Location.X + 15,
                      e.RowBounds.Location.Y + 4)
        End Using
    End Sub



    '*************************************************
    '*** Odœwiezanie filtrow kolumnowych 
    '*************************************************

    Private Sub stbFiltry_PanelClick(sender As Object, e As StatusBarPanelClickEventArgs) Handles stbFiltry.PanelClick
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKartyPayBack, DataViewKartyPayBack.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagKartyPayBack_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagKartyPayBack.CellFormatting
    '    If Not StartFormy Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagKartyPayBack, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagKartyPayBack, e.RowIndex, 4)

    '        Select Case UserName = Program_IdUzytkownika
    '            Case True
    '                e.CellStyle.ForeColor = Color.Red
    '            Case Else
    '                Select Case Mid(Upr, _Uprawnnienia_NrZnaku, 1)
    '                    Case "A"
    '                        e.CellStyle.ForeColor = Color.Green
    '                    Case "S"
    '                        e.CellStyle.ForeColor = Color.Purple
    '                    Case "U"
    '                        e.CellStyle.ForeColor = Color.Navy
    '                    Case Else
    '                        e.CellStyle.ForeColor = Color.Black
    '                End Select
    '        End Select
    '        '-----------------------
    '        ''zmieniamy BackColor
    '        'Dim Wal As String = GetTxtField(dagKursyWalut, e.RowIndex, 0)
    '        'Select Case Wal
    '        '    Case "EUR"
    '        '        e.CellStyle.BackColor = Color.White
    '        '    Case Else
    '        'End Select
    '        '-----------------------
    '    End If
    'End Sub




    '*********************************************************
    '***** obsluga dag
    '*********************************************************

    Private Sub dagKartyPayBack_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagKartyPayBack.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKartyPayBack, DataViewKartyPayBack.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKartyPayBack_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKartyPayBack.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKartyPayBack, DataViewKartyPayBack.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKartyPayBack_Scroll(sender As Object, e As ScrollEventArgs) Handles dagKartyPayBack.Scroll
        'If (Not StartFormy) And (DataViewKartyPayBack.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me, dagKartyPayBack, DataViewKartyPayBack.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me, dagKartyPayBack, DataViewKartyPayBack.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewKartyPayBack.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagKartyPayBack_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKartyPayBack.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me, dagKartyPayBack, DataViewKartyPayBack.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKartyPayBack_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagKartyPayBack.MouseMove
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None     '"the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                'If Not StartFormy Then
                '    PokazFiltryColumnDGV(Me, dagKartyPayBack, DataViewKartyPayBack.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me, dagKartyPayBack, DataViewKartyPayBack.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagKartyPayBack_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagKartyPayBack.MouseUp
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Dim message As String = "You clicked "

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None
                message &= "the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                message &= "cell at row " & hti.RowIndex & ", col " & hti.ColumnIndex
                stbKartyPayBack.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbKartyPayBack.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagKartyPayBack(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbKartyPayBack.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbKartyPayBack.Panels(3).Text = "Sort: " & DataSetKartyPayBack.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbKartyPayBack.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagKartyPayBack(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbKartyPayBack.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbKartyPayBack.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagKartyPayBack_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
        cmdEdytuj_Click(sender, e)
    End Sub

    Private Sub dagKartyPayBack_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case e.KeyCode
            Case Keys.Enter
                cmdEdytuj_Click(sender, e)
            Case Keys.Insert
                cmdDodaj_Click(sender, e)
            Case Keys.Delete
                cmdUsuñ_Click(sender, e)
            Case Else
        End Select
    End Sub



































    '*************************************************
    '*** przenoszenie danych miêdzy rekordem i zmiennymi
    '*************************************************

    Private Sub InitKartyPayBack()
        oNumerKartyPayBack = ""        'IDENT       Text, 10
    End Sub

    Private Sub PobierzKartyPayBack()
        'pobierz wartosci przed aktualizacja
        oNumerKartyPayBack = GetTxtField(dagKartyPayBack, 0)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Dim drv As DataRowView = Nothing
        If DataSetKartyPayBack.Tables(0).Rows.Count > 0 Then
            oNryKartPaybackKli = ""
            For Each drv In DataViewKartyPayBack
                oNumerKartyPayBack = GetTxtDRVField(drv, "NRKARTYPAYBACK")
                oNryKartPaybackKli &= oNumerKartyPayBack & ";"
            Next
            oNumer = GetTxtDRVField(DataViewKartyPayBack.Item(0), "NRKARTYPAYBACK")
        Else
            oNryKartPaybackKli = ""
            oNumer = ""
        End If
        Me.Close()
    End Sub

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If DataSetKartyPayBack.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzKartyPayBack()
            Dim oNumerKartyPrzedEdycja As String = oNumerKartyPayBack
            Dim Form As New EdytujKatalogKartPayBack
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                If oNumerKartyPrzedEdycja <> oNumerKartyPayBack Then
                    Dim foundRow As DataRow
                    ' Create an array for the key values to find.
                    Dim findTheseVals(0) As Object
                    findTheseVals(0) = oNumerKartyPayBack
                    foundRow = DataSetKartyPayBack.Tables(0).Rows.Find(findTheseVals)
                    ' sprawdz czy jest...
                    If Not (foundRow Is Nothing) Then
                        MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                            "Numer Karty PayBack = " & oNumerKartyPayBack & vbCrLf &
                            "Nie mogê wpisaæ kolejny raz tego Numeru Karty PayBack.",
                            "Uwaga :",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1)
                    Else
                        'dodaj nowy zapis
                        newRow = DataSetKartyPayBack.Tables(0).NewRow()
                        newRow("NRKARTYPAYBACK") = oNumerKartyPayBack
                        DataSetKartyPayBack.Tables(0).Rows.Add(newRow)

                        'Usun stary zapis
                        findTheseVals(0) = oNumerKartyPrzedEdycja
                        foundRow = DataSetKartyPayBack.Tables(0).Rows.Find(findTheseVals)
                        ' sprawdz czy jest...
                        If Not (foundRow Is Nothing) Then
                            foundRow.Delete()
                        Else
                        End If

                        'aktualizuj DataGrid
                        dagKartyPayBack.Update()
                        'wyswietl ilosc rekordow
                        stbKartyPayBack.Panels(0).Text = "Iloœæ rec.: " & DataSetKartyPayBack.Tables(0).Rows.Count.ToString
                    End If
                End If
            End If
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        If DataSetKartyPayBack.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie usun¹æ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            If MessageBox.Show("Czy na pewno mam usun¹æ ten zapis ?",
                        "Proszê o potwierdzenie :",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then
                oOperacja = "USUN"
                PobierzKartyPayBack()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = oNumerKartyPayBack
                foundRow = DataSetKartyPayBack.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    stbKartyPayBack.Panels(0).Text = "Iloœæ rec.: " & DataSetKartyPayBack.Tables(0).Rows.Count.ToString
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitKartyPayBack()
        Dim Form As New EdytujKatalogKartPayBack
        Do While True
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
                Exit Do
            Else
                Dim foundRow As DataRow
                Dim NewRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = oNumerKartyPayBack
                foundRow = DataSetKartyPayBack.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "Numer Karty PayBack = " & oNumerKartyPayBack,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetKartyPayBack.Tables(0).NewRow()
                    NewRow("NRKARTYPAYBACK") = oNumerKartyPayBack
                    DataSetKartyPayBack.Tables(0).Rows.Add(NewRow)
                    'aktualizuj DataGrid
                    dagKartyPayBack.Update()
                    'wyswietl ilosc rekordow
                    stbKartyPayBack.Panels(0).Text = "Iloœæ rec.: " & DataSetKartyPayBack.Tables(0).Rows.Count.ToString
                    Exit Do
                End If
            End If
        Loop
    End Sub



    '****************************************
    '** obsluga Menu Kontekstowego
    '****************************************

    Private Sub MenuUEDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuUEDodaj.Click
        cmdDodaj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub MenuUEEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuUEEdytuj.Click
        cmdEdytuj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub MenuUEUsun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuUEUsun.Click
        cmdUsuñ_Click(Me, System.EventArgs.Empty)
    End Sub


End Class
