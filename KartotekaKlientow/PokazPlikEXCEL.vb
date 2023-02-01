'Imports System.Data.Odbc
Imports System.Data.Oledb

Public Class PokazPlikEXCEL
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
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents dagEXCEL As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PokazPlikEXCEL))
        Me.dagEXCEL = New System.Windows.Forms.DataGrid()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cmdStop = New System.Windows.Forms.Button()
        CType(Me.dagEXCEL, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.dagEXCEL.Size = New System.Drawing.Size(512, 271)
        Me.dagEXCEL.TabIndex = 0
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 287)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 25)
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
        Me.cmdStop.Location = New System.Drawing.Point(424, 287)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(87, 24)
        Me.cmdStop.TabIndex = 41
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PokazPlikEXCEL
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(528, 318)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.dagEXCEL)
        Me.Name = "PokazPlikEXCEL"
        Me.ShowInTaskbar = False
        Me.Text = "Pokaz plik EXCEL"
        CType(Me.dagEXCEL, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    '-------------odbc
    'Private EXCELConnection As String = "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & oPlikNaDysku & ";"
    Private EXCELConnectionODBC As String = _
                "Driver={Microsoft Excel Driver (*.xls)};" & _
                "DSN=Pliki programu Excel;" & _
                "DBQ=" & oPlikNaDysku & ";" & _
                "DriverId=790;" & _
                "MaxBufferSize=2048;" & _
                "PageTimeout=5;"

    '-------------oleDB
    Private EXCELConnectionOLeDB As String = ""
    '=========================================================
    'Connection String: 
    'The connection string should be set to the OleDbConnection object.
    'This is very critical as Jet Engine might not give a proper error message
    'if the appropriate details are not given.
    '
    'Syntax: Provider=Microsoft.Jet.OLEDB.4.0;
    '                 Data Source=<Full Path of Excel File>;
    '                 Extended Properties="Excel 8.0; HDR=No; IMEX=1".
    '
    'Definition of Extended Properties: 
    '   Excel = <No> 
    '   One should specify the version of Excel Sheet here.
    '   For Excel 2000 and above, it is set it to Excel 8.0
    '   and for all others, it is Excel 5.0.
    '
    '   HDR= <Yes/No> 
    '   This property will be used to specify the definition of header
    '   for each column. If the value is ‘Yes’, the first row will be
    '   treated as heading. Otherwise, the heading will be generated
    '   by the system like F1, F2 and so on.
    '
    '   IMEX= <0/1/2> 
    '   IMEX refers to IMport EXport mode. This can take three possible values.
    '   IMEX=0 and IMEX=2 will result in ImportMixedTypes being ignored and the default value of  
    '       'Majority Types’ is used. In this case, it will take the first 8 rows
    '       and then the data type for each column will be decided. 
    '   IMEX=1 is the only way to set the value of ImportMixedTypes
    '       as Text. Here, everything will be treated as text. 
    '=========================================================

    Private Sub PokazPlikEXCEL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        If OSArchitectureIs64bit() Then
            ''ACE
            'If you are an application developer using OLEDB, set the Provider argument 
            'of the ConnectionString property to “Microsoft.ACE.OLEDB.12.0”. 
            EXCELConnectionOLeDB = "Provider=""Microsoft.ACE.OLEDB.12.0"";" & _
                                    "Data Source=" & oPlikNaDysku & "; " & _
                                    "Extended Properties='Excel 8.0; HDR=YES; IMEX=1';"

            'If you are an application developer using ODBC to connect to Microsoft Office Access data, 
            'set the Connection String to “Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=path to mdb/accdb file”
            'If you are an application developer using ODBC to connect to Microsoft Office Excel data, 
            'set the Connection String to “Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=path to xls/xlsx/xlsm/xlsb file”
            'Return ("Driver={Microsoft Access Driver (*.mdb, *.accdb)};" & _
            '        "DBQ=""" & ParL_KatZbiorow & "\" & oPlikKartoteki & """ ")
        Else
            'Jet 4,0
            EXCELConnectionOLeDB = "Provider=Microsoft.Jet.OLEDB.4.0; " & _
                                "Data Source=" & oPlikNaDysku & "; " & _
                                "Extended Properties='Excel 8.0; HDR=YES; IMEX=1';"
        End If

        '-------------odbc
        'Dim EXCELCon As New OdbcConnection(EXCELConnectionODBC)
        'Dim EXCELAdapter As New OdbcDataAdapter("SELECT * FROM [Arkusz1$]", EXCELCon)

        '-------------oleDB
        Dim EXCELCon As New OleDbConnection(EXCELConnectionOLeDB)
        Dim EXCELAdapter As New OleDbDataAdapter("SELECT * FROM [Arkusz1$]", EXCELCon)

        Dim EXCELds As New DataSet
        Dim EXCELdv As New DataView

        'wypelnij DATASET
        Try
            EXCELCon.Open()
        Catch Ex As System.Exception
            ViewInLog(Ex, Me.Name(), Nothing)
            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    MessageBoxIcon.Information, _
            '    MessageBoxDefaultButton.Button1)
        End Try

        If EXCELCon.State = ConnectionState.Open Then
            Try
                EXCELAdapter.Fill(EXCELds, "Arkusz1")
            Catch Ex As System.Exception
                MessageBox.Show(Err.ToString())
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                EXCELCon.Close()
            End Try

            'parametry DataGrid
            dagEXCEL.CaptionVisible = False
            dagEXCEL.CaptionText = EXCELDs.Tables(0).TableName
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
            If EXCELds.Tables(0).Rows.Count > 0 Then dagEXCEL.Select(0)
        End If
    End Sub

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub

    Private Sub PokazPlikEXCEL_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub
End Class
