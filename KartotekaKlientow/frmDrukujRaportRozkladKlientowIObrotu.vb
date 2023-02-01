Imports System.Drawing.Printing

Public Class frmDrukujRaportRokladKlientowIObrotu
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    'wywolanie z menu
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
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents cmdDrukuj As System.Windows.Forms.Button
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblPopRok As System.Windows.Forms.Label
    Friend WithEvents cbxBiezRok As System.Windows.Forms.ComboBox
    Friend WithEvents lblRaport As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDrukujRaportRokladKlientowIObrotu))
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.cmdDrukuj = New System.Windows.Forms.Button()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblRaport = New System.Windows.Forms.Label()
        Me.lblPopRok = New System.Windows.Forms.Label()
        Me.cbxBiezRok = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.Location = New System.Drawing.Point(8, 125)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(104, 23)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 84
        Me.picLogo.TabStop = False
        '
        'cmdDrukuj
        '
        Me.cmdDrukuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDrukuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDrukuj.Image = CType(resources.GetObject("cmdDrukuj.Image"), System.Drawing.Image)
        Me.cmdDrukuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDrukuj.Location = New System.Drawing.Point(299, 126)
        Me.cmdDrukuj.Name = "cmdDrukuj"
        Me.cmdDrukuj.Size = New System.Drawing.Size(70, 22)
        Me.cmdDrukuj.TabIndex = 83
        Me.cmdDrukuj.Text = "Raport"
        Me.cmdDrukuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(376, 126)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(80, 22)
        Me.cmdStop.TabIndex = 82
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lblRaport)
        Me.Panel1.Controls.Add(Me.lblPopRok)
        Me.Panel1.Controls.Add(Me.cbxBiezRok)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(452, 112)
        Me.Panel1.TabIndex = 81
        '
        'lblRaport
        '
        Me.lblRaport.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRaport.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblRaport.ForeColor = System.Drawing.Color.Navy
        Me.lblRaport.Location = New System.Drawing.Point(12, 67)
        Me.lblRaport.Name = "lblRaport"
        Me.lblRaport.Size = New System.Drawing.Size(352, 16)
        Me.lblRaport.TabIndex = 367
        Me.lblRaport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblPopRok
        '
        Me.lblPopRok.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPopRok.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblPopRok.ForeColor = System.Drawing.Color.Navy
        Me.lblPopRok.Location = New System.Drawing.Point(196, 37)
        Me.lblPopRok.Name = "lblPopRok"
        Me.lblPopRok.Size = New System.Drawing.Size(121, 21)
        Me.lblPopRok.TabIndex = 366
        '
        'cbxBiezRok
        '
        Me.cbxBiezRok.FormattingEnabled = True
        Me.cbxBiezRok.Location = New System.Drawing.Point(196, 13)
        Me.cbxBiezRok.Name = "cbxBiezRok"
        Me.cbxBiezRok.Size = New System.Drawing.Size(121, 21)
        Me.cbxBiezRok.TabIndex = 365
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(12, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(267, 16)
        Me.Label2.TabIndex = 364
        Me.Label2.Text = "Poprzedni Rok Obliczeniowy . . . . . . . . . . . . . . . . . ."
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(12, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(260, 16)
        Me.Label1.TabIndex = 336
        Me.Label1.Text = "Bie¿¹cy Rok Obliczeniowy . . . . . . . . . . . . . . . . . . . ."
        '
        'frmDrukujRaportRokladKlientowIObrotu
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(468, 153)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.cmdDrukuj)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "frmDrukujRaportRokladKlientowIObrotu"
        Me.Text = "Raport Rozk³ad Klientów i Obrotu"
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Dim KlienciOK As Boolean = False
    Dim KontaktyOK As Boolean = False
    Dim ObrotyOK As Boolean = False
    Dim ObrotyMiesOK As Boolean = False
    Dim RoboczyOK As Boolean = False

    Dim Tytul As String
    Dim LicznikRec As Long = 0
    '------------------------------------------------------------------
    Dim dbSelectSQLKlienci As String = ""       'sSelectSQLKlienci

    Dim dbConnectionKlienci As OleDb.OleDbConnection
    Dim dbCommandSelectKlienci As OleDb.OleDbCommand
    Dim dbCommandDeleteKlienci As OleDb.OleDbCommand
    Dim dbCommandUpdateKlienci As OleDb.OleDbCommand
    Dim dbCommandInsertKlienci As OleDb.OleDbCommand
    Dim dbDataAdapterKlienci As OleDb.OleDbDataAdapter

    Dim sqlConnectionKlienci As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienci As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienci As SqlClient.SqlCommand
    Dim sqlCommandUpdateKlienci As SqlClient.SqlCommand
    Dim sqlCommandInsertKlienci As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienci As SqlClient.SqlDataAdapter

    Dim DataSetKlienci As DataSet
    Dim DataViewKlienci As DataView
    '------------------------------------------------------------------
    Dim dbSelectSQLKontakty As String = sSelectSQLKontakty

    Dim dbConnectionKontakty As OleDb.OleDbConnection
    Dim dbCommandSelectKontakty As OleDb.OleDbCommand
    Dim dbCommandDeleteKontakty As OleDb.OleDbCommand
    Dim dbCommandUpdateKontakty As OleDb.OleDbCommand
    Dim dbCommandInsertKontakty As OleDb.OleDbCommand
    Dim dbDataAdapterKontakty As OleDb.OleDbDataAdapter

    Dim sqlConnectionKontakty As SqlClient.SqlConnection
    Dim sqlCommandSelectKontakty As SqlClient.SqlCommand
    Dim sqlCommandDeleteKontakty As SqlClient.SqlCommand
    Dim sqlCommandUpdateKontakty As SqlClient.SqlCommand
    Dim sqlCommandInsertKontakty As SqlClient.SqlCommand
    Dim sqlDataAdapterKontakty As SqlClient.SqlDataAdapter

    Dim DataSetKontakty As DataSet
    Dim DataViewKontakty As DataView
    '------------------------------------------------------------------
    Dim dbSelectSQLObroty As String = sSelectSQLObroty

    Dim dbConnectionObroty As OleDb.OleDbConnection
    Dim dbCommandSelectObroty As OleDb.OleDbCommand
    Dim dbCommandDeleteObroty As OleDb.OleDbCommand
    Dim dbCommandUpdateObroty As OleDb.OleDbCommand
    Dim dbCommandInsertObroty As OleDb.OleDbCommand
    Dim dbDataAdapterObroty As OleDb.OleDbDataAdapter

    Dim sqlConnectionObroty As SqlClient.SqlConnection
    Dim sqlCommandSelectObroty As SqlClient.SqlCommand
    Dim sqlCommandDeleteObroty As SqlClient.SqlCommand
    Dim sqlCommandUpdateObroty As SqlClient.SqlCommand
    Dim sqlCommandInsertObroty As SqlClient.SqlCommand
    Dim sqlDataAdapterObroty As SqlClient.SqlDataAdapter

    Dim DataSetObroty As DataSet
    Dim DataViewObroty As DataView
    '------------------------------------------------------------------
    Dim dbSelectSQLObrotyMies As String = sSelectSQLObrotyMies

    Dim dbConnectionObrotyMies As OleDb.OleDbConnection
    Dim dbCommandSelectObrotyMies As OleDb.OleDbCommand
    Dim dbCommandDeleteObrotyMies As OleDb.OleDbCommand
    Dim dbCommandUpdateObrotyMies As OleDb.OleDbCommand
    Dim dbCommandInsertObrotyMies As OleDb.OleDbCommand
    Dim dbDataAdapterObrotyMies As OleDb.OleDbDataAdapter

    Dim sqlConnectionObrotyMies As SqlClient.SqlConnection
    Dim sqlCommandSelectObrotyMies As SqlClient.SqlCommand
    Dim sqlCommandDeleteObrotyMies As SqlClient.SqlCommand
    Dim sqlCommandUpdateObrotyMies As SqlClient.SqlCommand
    Dim sqlCommandInsertObrotyMies As SqlClient.SqlCommand
    Dim sqlDataAdapterObrotyMies As SqlClient.SqlDataAdapter

    Dim DataSetObrotyMies As DataSet
    Dim DataViewObrotyMies As DataView
    '------------------------------------------------------------------
    Dim ConnectionState As System.Data.ConnectionState



    'zakladamy tablice
    Dim TableRoboczy As System.Data.DataTable = Nothing
    Dim DataSetRoboczy As DataSet = Nothing
    Dim DataViewRoboczy As DataView = Nothing
    Dim findRobo(1) As Object









    Private Sub DrukujZestawienie_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.picLogo.Image = My.Resources.logomini
        Me.Icon = My.Resources.MrJones
        '------------------------------------
        InicjujListeLat(cbxBiezRok)
        cbxBiezRok.Text = Mid(SysData, 1, 4)
        lblPopRok.Text = Microsoft.VisualBasic.Right("0000" & Str(Val(Trim(cbxBiezRok.Text)) - 1), 4)
        lblRaport.Text = ""
    End Sub

    Private Sub frmDrukujRaportRokladKlientowIObrotu_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub

    Private Sub cbxBiezRok_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxBiezRok.SelectedIndexChanged
        lblPopRok.Text = Microsoft.VisualBasic.Right("0000" & Str(Val(Trim(cbxBiezRok.Text)) - 1), 4)
    End Sub



    '-----------------------------------
    ' wylicza wartosc graniczna poprzedniego roku
    '-----------------------------------

    Private Function WyliczWartoscGraniczna(ByVal pRok As String) As Double
        Dim CalkSprzedaz As Double = 0
        Dim GranSprzedaz As Double = 0
        Dim WartGraniczna As Double = 0

        'klienci o sprzeda¿y poni¿ej 1000 z³ w roku
        dbSelectSQLKlienci = "SELECT * FROM " & _
                             "(SELECT IDENTKLIENTA," & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pRok & "')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pRok & "')),0)) AS SPRZEDAZ " & _
                            " FROM Klienci) AS KL " & _
                            " WHERE (KL.SPRZEDAZ>0) AND (KL.SPRZEDAZ<=1000) ORDER BY KL.SPRZEDAZ ASC"
        DataSetKlienci = New DataSet
        DataViewKlienci = New DataView

        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionKlienci = New OleDb.OleDbConnection(sConnectionKlienci)
            'dbCommandSelectKlienci = New OleDb.OleDbCommand(dbSelectSQLKlienci, dbConnectionKlienci)
            ''dbCommandDeleteKlienci = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnectionKlienci)
            ''dbCommandUpdateKlienci = New OleDb.OleDbCommand(sUpdateSQLKlienci, dbConnectionKlienci)
            ''dbCommandInsertKlienci = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnectionKlienci)
            'dbDataAdapterKlienci = New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)

            'Dim DBMapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            'DBMapowanieTabeliKlienci = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionKlienci.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnectionKlienci.State
            'End Try
        Else
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(dbSelectSQLKlienci, sqlConnectionKlienci)
            'sqlCommandDeleteKlienci = New sqlclient.sqlCommand(sDeleteSQLKlienci, sqlconnectionKlienci)
            'sqlCommandUpdateKlienci = New sqlclient.sqlCommand(sUpdateSQLKlienci, sqlconnectionKlienci)
            'sqlCommandInsertKlienci = New sqlclient.sqlCommand(sInsertSQLKlienci, sqlconnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

            Dim sqlMapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliKlienci = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionKlienci.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    dbDataAdapterKlienci.Fill(DataSetKlienci)
                    dbConnectionKlienci.Close()
                Else
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                End If
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}
                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))

                Dim i As Integer = 0
                If DataViewKlienci.Count > 0 Then
                    'liczymy ca³kowit¹ sprzedaz w tej grupie
                    CalkSprzedaz = 0
                    For i = 0 To DataViewKlienci.Count - 1
                        CalkSprzedaz += DataViewKlienci.Item(i).Item("SPRZEDAZ")
                    Next

                    'liczymy granice wartosci 80 % sprzedazy w tej grupie
                    GranSprzedaz = 0.8 * CalkSprzedaz

                    'liczymy wartosc graniczn¹ 
                    CalkSprzedaz = 0
                    For i = 0 To DataViewKlienci.Count - 1
                        CalkSprzedaz += DataViewKlienci.Item(i).Item("SPRZEDAZ")
                        If CalkSprzedaz > GranSprzedaz Then
                            Exit For
                        Else
                            WartGraniczna = DataViewKlienci.Item(i).Item("SPRZEDAZ")
                        End If
                    Next

                End If

            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)

            End Try
        End If

        Return WartGraniczna
    End Function








    '-----------------------------------
    ' pobierz dane do zestawienia
    '-----------------------------------

    Dim rPozycja As Integer = 0
    Dim rIdent As String = ""
    Dim rWykonanie As Double = 0
    Dim rWykonanie01 As Double = 0
    Dim rWykonanie02 As Double = 0
    Dim rWykonanie03 As Double = 0
    Dim rWykonanie04 As Double = 0
    Dim rWykonanie05 As Double = 0
    Dim rWykonanie06 As Double = 0
    Dim rWykonanie07 As Double = 0
    Dim rWykonanie08 As Double = 0
    Dim rWykonanie09 As Double = 0
    Dim rWykonanie10 As Double = 0
    Dim rWykonanie11 As Double = 0
    Dim rWykonanie12 As Double = 0
    Dim rPlan As Double = 0
    Dim rPlan01 As Double = 0
    Dim rPlan02 As Double = 0
    Dim rPlan03 As Double = 0
    Dim rPlan04 As Double = 0
    Dim rPlan05 As Double = 0
    Dim rPlan06 As Double = 0
    Dim rPlan07 As Double = 0
    Dim rPlan08 As Double = 0
    Dim rPlan09 As Double = 0
    Dim rPlan10 As Double = 0
    Dim rPlan11 As Double = 0
    Dim rPlan12 As Double = 0
    Dim rPopRok As Double = 0
    Dim rPopRok01 As Double = 0
    Dim rPopRok02 As Double = 0
    Dim rPopRok03 As Double = 0
    Dim rPopRok04 As Double = 0
    Dim rPopRok05 As Double = 0
    Dim rPopRok06 As Double = 0
    Dim rPopRok07 As Double = 0
    Dim rPopRok08 As Double = 0
    Dim rPopRok09 As Double = 0
    Dim rPopRok10 As Double = 0
    Dim rPopRok11 As Double = 0
    Dim rPopRok12 As Double = 0

    Dim WartGraniczna As Double = 0

    Private Sub InicjujRoboczy(ByVal pPozycja As Integer, _
                       ByVal pIdent As String)
        rPozycja = pPozycja
        rIdent = pIdent
        rWykonanie = 0
        rWykonanie01 = 0
        rWykonanie02 = 0
        rWykonanie03 = 0
        rWykonanie04 = 0
        rWykonanie05 = 0
        rWykonanie06 = 0
        rWykonanie07 = 0
        rWykonanie08 = 0
        rWykonanie09 = 0
        rWykonanie10 = 0
        rWykonanie11 = 0
        rWykonanie12 = 0
        rPlan = 0
        rPlan01 = 0
        rPlan02 = 0
        rPlan03 = 0
        rPlan04 = 0
        rPlan05 = 0
        rPlan06 = 0
        rPlan07 = 0
        rPlan08 = 0
        rPlan09 = 0
        rPlan10 = 0
        rPlan11 = 0
        rPlan12 = 0
        rPopRok = 0
        rPopRok01 = 0
        rPopRok02 = 0
        rPopRok03 = 0
        rPopRok04 = 0
        rPopRok05 = 0
        rPopRok06 = 0
        rPopRok07 = 0
        rPopRok08 = 0
        rPopRok09 = 0
        rPopRok10 = 0
        rPopRok11 = 0
        rPopRok12 = 0
    End Sub

    Private Sub ZapiszRoboczy()
        Dim dr As DataRow = Nothing

        dr = DataSetRoboczy.Tables(0).NewRow()

        dr("POZYCJA") = rPozycja
        dr("IDENT") = rIdent
        dr("WYKONANIE") = rWykonanie
        dr("WYKONANIE01") = rWykonanie01
        dr("WYKONANIE02") = rWykonanie02
        dr("WYKONANIE03") = rWykonanie03
        dr("WYKONANIE04") = rWykonanie04
        dr("WYKONANIE05") = rWykonanie05
        dr("WYKONANIE06") = rWykonanie06
        dr("WYKONANIE07") = rWykonanie07
        dr("WYKONANIE08") = rWykonanie08
        dr("WYKONANIE09") = rWykonanie09
        dr("WYKONANIE10") = rWykonanie10
        dr("WYKONANIE11") = rWykonanie11
        dr("WYKONANIE12") = rWykonanie12
        dr("PLAN") = rPlan
        dr("PLAN01") = rPlan01
        dr("PLAN02") = rPlan02
        dr("PLAN03") = rPlan03
        dr("PLAN04") = rPlan04
        dr("PLAN05") = rPlan05
        dr("PLAN06") = rPlan06
        dr("PLAN07") = rPlan07
        dr("PLAN08") = rPlan08
        dr("PLAN09") = rPlan09
        dr("PLAN10") = rPlan10
        dr("PLAN11") = rPlan11
        dr("PLAN12") = rPlan12
        dr("POPROK") = rPopRok
        dr("POPROK01") = rPopRok01
        dr("POPROK02") = rPopRok02
        dr("POPROK03") = rPopRok03
        dr("POPROK04") = rPopRok04
        dr("POPROK05") = rPopRok05
        dr("POPROK06") = rPopRok06
        dr("POPROK07") = rPopRok07
        dr("POPROK08") = rPopRok08
        dr("POPROK09") = rPopRok09
        dr("POPROK10") = rPopRok10
        dr("POPROK11") = rPopRok11
        dr("POPROK12") = rPopRok12

        DataSetRoboczy.Tables(0).Rows.Add(dr)
    End Sub



    Private Function PrzygotujDane(ByVal pBiezRok As String, _
                                   ByVal pPopRok As String) As Boolean
        Dim RetValue As Boolean = False
        Dim i As Integer = 0
        Dim drv As DataRowView = Nothing
        Dim dr As DataRow = Nothing

        lblRaport.Text = "1/4 Wyliczam wartoœæ graniczn¹"
        System.Windows.Forms.Application.DoEvents()
        WartGraniczna = WyliczWartoscGraniczna(pPopRok)

        lblRaport.Text = "2/4 Analizujê Obroty"
        System.Windows.Forms.Application.DoEvents()
        TableRoboczy = New DataTable()
        DataSetRoboczy = New DataSet()
        DataViewRoboczy = Nothing

        Dim robCol1 As DataColumn = TableRoboczy.Columns.Add("POZYCJA", GetType(System.Int32))
        Dim robCol2 As DataColumn = TableRoboczy.Columns.Add("IDENT", GetType(System.String))
        Dim robCol5 As DataColumn = TableRoboczy.Columns.Add("WYKONANIE", GetType(System.Double))
        Dim robCol6 As DataColumn = TableRoboczy.Columns.Add("WYKONANIE01", GetType(System.Double))
        Dim robCol7 As DataColumn = TableRoboczy.Columns.Add("WYKONANIE02", GetType(System.Double))
        Dim robCol8 As DataColumn = TableRoboczy.Columns.Add("WYKONANIE03", GetType(System.Double))
        Dim robCol9 As DataColumn = TableRoboczy.Columns.Add("WYKONANIE04", GetType(System.Double))
        Dim robColA As DataColumn = TableRoboczy.Columns.Add("WYKONANIE05", GetType(System.Double))
        Dim robColB As DataColumn = TableRoboczy.Columns.Add("WYKONANIE06", GetType(System.Double))
        Dim robColC As DataColumn = TableRoboczy.Columns.Add("WYKONANIE07", GetType(System.Double))
        Dim robColD As DataColumn = TableRoboczy.Columns.Add("WYKONANIE08", GetType(System.Double))
        Dim robColE As DataColumn = TableRoboczy.Columns.Add("WYKONANIE09", GetType(System.Double))
        Dim robColF As DataColumn = TableRoboczy.Columns.Add("WYKONANIE10", GetType(System.Double))
        Dim robColG As DataColumn = TableRoboczy.Columns.Add("WYKONANIE11", GetType(System.Double))
        Dim robColH As DataColumn = TableRoboczy.Columns.Add("WYKONANIE12", GetType(System.Double))

        Dim robCol5p As DataColumn = TableRoboczy.Columns.Add("PLAN", GetType(System.Double))
        Dim robCol6p As DataColumn = TableRoboczy.Columns.Add("PLAN01", GetType(System.Double))
        Dim robCol7p As DataColumn = TableRoboczy.Columns.Add("PLAN02", GetType(System.Double))
        Dim robCol8p As DataColumn = TableRoboczy.Columns.Add("PLAN03", GetType(System.Double))
        Dim robCol9p As DataColumn = TableRoboczy.Columns.Add("PLAN04", GetType(System.Double))
        Dim robColAp As DataColumn = TableRoboczy.Columns.Add("PLAN05", GetType(System.Double))
        Dim robColBp As DataColumn = TableRoboczy.Columns.Add("PLAN06", GetType(System.Double))
        Dim robColCp As DataColumn = TableRoboczy.Columns.Add("PLAN07", GetType(System.Double))
        Dim robColDp As DataColumn = TableRoboczy.Columns.Add("PLAN08", GetType(System.Double))
        Dim robColEp As DataColumn = TableRoboczy.Columns.Add("PLAN09", GetType(System.Double))
        Dim robColFp As DataColumn = TableRoboczy.Columns.Add("PLAN10", GetType(System.Double))
        Dim robColGp As DataColumn = TableRoboczy.Columns.Add("PLAN11", GetType(System.Double))
        Dim robColHp As DataColumn = TableRoboczy.Columns.Add("PLAN12", GetType(System.Double))

        Dim robCol5r As DataColumn = TableRoboczy.Columns.Add("POPROK", GetType(System.Double))
        Dim robCol6r As DataColumn = TableRoboczy.Columns.Add("POPROK01", GetType(System.Double))
        Dim robCol7r As DataColumn = TableRoboczy.Columns.Add("POPROK02", GetType(System.Double))
        Dim robCol8r As DataColumn = TableRoboczy.Columns.Add("POPROK03", GetType(System.Double))
        Dim robCol9r As DataColumn = TableRoboczy.Columns.Add("POPROK04", GetType(System.Double))
        Dim robColAr As DataColumn = TableRoboczy.Columns.Add("POPROK05", GetType(System.Double))
        Dim robColBr As DataColumn = TableRoboczy.Columns.Add("POPROK06", GetType(System.Double))
        Dim robColCr As DataColumn = TableRoboczy.Columns.Add("POPROK07", GetType(System.Double))
        Dim robColDr As DataColumn = TableRoboczy.Columns.Add("POPROK08", GetType(System.Double))
        Dim robColEr As DataColumn = TableRoboczy.Columns.Add("POPROK09", GetType(System.Double))
        Dim robColFr As DataColumn = TableRoboczy.Columns.Add("POPROK10", GetType(System.Double))
        Dim robColGr As DataColumn = TableRoboczy.Columns.Add("POPROK11", GetType(System.Double))
        Dim robColHr As DataColumn = TableRoboczy.Columns.Add("POPROK12", GetType(System.Double))


        DataSetRoboczy.Tables.Add(TableRoboczy)
        'definiuj DataView
        DataViewRoboczy = New DataView(DataSetRoboczy.Tables(0))
        DataViewRoboczy.Sort = "POZYCJA"
        '-------------------------------------------
        dbSelectSQLKlienci = "SELECT * FROM " & _
                             "(SELECT IDENTKLIENTA, ROKOWANIABR, " & _
                                 "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "')),0) + " & _
                                 " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "')),0)) AS SPRZEDAZPOPROK, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='01')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='01')),0)) AS SPRZEDAZPOPROK01, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='02')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='02')),0)) AS SPRZEDAZPOPROK02, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='03')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='03')),0)) AS SPRZEDAZPOPROK03, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='04')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='04')),0)) AS SPRZEDAZPOPROK04, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='05')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='05')),0)) AS SPRZEDAZPOPROK05, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='06')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='06')),0)) AS SPRZEDAZPOPROK06, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='07')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='07')),0)) AS SPRZEDAZPOPROK07, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='08')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='08')),0)) AS SPRZEDAZPOPROK08, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='09')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='09')),0)) AS SPRZEDAZPOPROK09, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='10')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='10')),0)) AS SPRZEDAZPOPROK10, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='11')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='11')),0)) AS SPRZEDAZPOPROK11, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='12')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='12')),0)) AS SPRZEDAZPOPROK12, " & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "')),0)) AS SPRZEDAZ, " & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='01')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='01')),0)) AS SPRZEDAZ01, " & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='02')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='02')),0)) AS SPRZEDAZ02, " & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='03')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='03')),0)) AS SPRZEDAZ03, " & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='04')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='04')),0)) AS SPRZEDAZ04, " & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='05')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='05')),0)) AS SPRZEDAZ05, " & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='06')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='06')),0)) AS SPRZEDAZ06, " & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='07')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='07')),0)) AS SPRZEDAZ07, " & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='08')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='08')),0)) AS SPRZEDAZ08, " & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='09')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='09')),0)) AS SPRZEDAZ09, " & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='10')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='10')),0)) AS SPRZEDAZ10, " & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='11')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='11')),0)) AS SPRZEDAZ11, " & _
                                    "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='12')),0) + " & _
                                    " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='12')),0)) AS SPRZEDAZ12, " & _
                                "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "')),0) + " & _
                                " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "')),0)) AS SPRZEDAZLPOPROK, " & _
                                "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='01')),0) + " & _
                                " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='01')),0)) AS SPRZEDAZLPOPROK01, " & _
                                "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='02')),0) + " & _
                                " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='02')),0)) AS SPRZEDAZLPOPROK02, " & _
                                "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='03')),0) + " & _
                                " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='03')),0)) AS SPRZEDAZLPOPROK03, " & _
                                "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='04')),0) + " & _
                                " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='04')),0)) AS SPRZEDAZLPOPROK04, " & _
                                "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='05')),0) + " & _
                                " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='05')),0)) AS SPRZEDAZLPOPROK05, " & _
                                "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='06')),0) + " & _
                                " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='06')),0)) AS SPRZEDAZLPOPROK06, " & _
                                "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='07')),0) + " & _
                                " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='07')),0)) AS SPRZEDAZLPOPROK07, " & _
                                "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='08')),0) + " & _
                                " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='08')),0)) AS SPRZEDAZLPOPROK08, " & _
                                "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='09')),0) + " & _
                                " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='09')),0)) AS SPRZEDAZLPOPROK09, " & _
                                "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='10')),0) + " & _
                                " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='10')),0)) AS SPRZEDAZLPOPROK10, " & _
                                "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='11')),0) + " & _
                                " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='11')),0)) AS SPRZEDAZLPOPROK11, " & _
                                "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='12')),0) + " & _
                                " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='12')),0)) AS SPRZEDAZLPOPROK12, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "')),0)) AS SPRZEDAZL, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='01')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='01')),0)) AS SPRZEDAZL01, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='02')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='02')),0)) AS SPRZEDAZL02, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='03')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='03')),0)) AS SPRZEDAZL03, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='04')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='04')),0)) AS SPRZEDAZL04, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='05')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='05')),0)) AS SPRZEDAZL05, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='06')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='06')),0)) AS SPRZEDAZL06, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='07')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='07')),0)) AS SPRZEDAZL07, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='08')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='08')),0)) AS SPRZEDAZL08, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='09')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='09')),0)) AS SPRZEDAZL09, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='10')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='10')),0)) AS SPRZEDAZL10, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='11')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='11')),0)) AS SPRZEDAZL11, " & _
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='12')),0) + " & _
                                  " ISNULL((SELECT SUM(LWARTSPRZED+LOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='12')),0)) AS SPRZEDAZL12, " & _
                              "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "')),0) + " & _
                              " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "')),0)) AS SPRZEDAZAPOPROK, " & _
                              "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='01')),0) + " & _
                              " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='01')),0)) AS SPRZEDAZAPOPROK01, " & _
                              "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='02')),0) + " & _
                              " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='02')),0)) AS SPRZEDAZAPOPROK02, " & _
                              "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='03')),0) + " & _
                              " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='03')),0)) AS SPRZEDAZAPOPROK03, " & _
                              "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='04')),0) + " & _
                              " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='04')),0)) AS SPRZEDAZAPOPROK04, " & _
                              "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='05')),0) + " & _
                              " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='05')),0)) AS SPRZEDAZAPOPROK05, " & _
                              "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='06')),0) + " & _
                              " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='06')),0)) AS SPRZEDAZAPOPROK06, " & _
                              "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='07')),0) + " & _
                              " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='07')),0)) AS SPRZEDAZAPOPROK07, " & _
                              "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='08')),0) + " & _
                              " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='08')),0)) AS SPRZEDAZAPOPROK08, " & _
                              "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='09')),0) + " & _
                              " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='09')),0)) AS SPRZEDAZAPOPROK09, " & _
                              "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='10')),0) + " & _
                              " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='10')),0)) AS SPRZEDAZAPOPROK10, " & _
                              "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='11')),0) + " & _
                              " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='11')),0)) AS SPRZEDAZAPOPROK11, " & _
                              "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='12')),0) + " & _
                              " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='12')),0)) AS SPRZEDAZAPOPROK12, " & _
                                "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "')),0) + " & _
                                " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "')),0)) AS SPRZEDAZA, " & _
                                "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='01')),0) + " & _
                                " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='01')),0)) AS SPRZEDAZA01, " & _
                                "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='02')),0) + " & _
                                " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='02')),0)) AS SPRZEDAZA02, " & _
                                "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='03')),0) + " & _
                                " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='03')),0)) AS SPRZEDAZA03, " & _
                                "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='04')),0) + " & _
                                " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='04')),0)) AS SPRZEDAZA04, " & _
                                "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='05')),0) + " & _
                                " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='05')),0)) AS SPRZEDAZA05, " & _
                                "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='06')),0) + " & _
                                " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='06')),0)) AS SPRZEDAZA06, " & _
                                "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='07')),0) + " & _
                                " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='07')),0)) AS SPRZEDAZA07, " & _
                                "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='08')),0) + " & _
                                " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='08')),0)) AS SPRZEDAZA08, " & _
                                "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='09')),0) + " & _
                                " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='09')),0)) AS SPRZEDAZA09, " & _
                                "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='10')),0) + " & _
                                " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='10')),0)) AS SPRZEDAZA10, " & _
                                "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='11')),0) + " & _
                                " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='11')),0)) AS SPRZEDAZA11, " & _
                                "(ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='12')),0) + " & _
                                " ISNULL((SELECT SUM(AWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='12')),0)) AS SPRZEDAZA12, " & _
                            "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "')),0) + " & _
                            " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "')),0)) AS MARZAPOPROK, " & _
                            "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='01')),0) + " & _
                            " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='01')),0)) AS MARZAPOPROK01, " & _
                            "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='02')),0) + " & _
                            " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='02')),0)) AS MARZAPOPROK02, " & _
                            "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='03')),0) + " & _
                            " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='03')),0)) AS MARZAPOPROK03, " & _
                            "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='04')),0) + " & _
                            " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='04')),0)) AS MARZAPOPROK04, " & _
                            "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='05')),0) + " & _
                            " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='05')),0)) AS MARZAPOPROK05, " & _
                            "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='06')),0) + " & _
                            " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='06')),0)) AS MARZAPOPROK06, " & _
                            "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='07')),0) + " & _
                            " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='07')),0)) AS MARZAPOPROK07, " & _
                            "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='08')),0) + " & _
                            " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='08')),0)) AS MARZAPOPROK08, " & _
                            "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='09')),0) + " & _
                            " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='09')),0)) AS MARZAPOPROK09, " & _
                            "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='10')),0) + " & _
                            " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='10')),0)) AS MARZAPOPROK10, " & _
                            "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='11')),0) + " & _
                            " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='11')),0)) AS MARZAPOPROK11, " & _
                            "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pPopRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='12')),0) + " & _
                            " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pPopRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='12')),0)) AS MARZAPOPROK12, " & _
                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARZAO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "')),0) + " & _
                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARZAM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "')),0)) AS MARZA, " & _
                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='01')),0) + " & _
                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='01')),0)) AS MARZA01, " & _
                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='02')),0) + " & _
                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='02')),0)) AS MARZA02, " & _
                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='03')),0) + " & _
                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='03')),0)) AS MARZA03, " & _
                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='04')),0) + " & _
                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='04')),0)) AS MARZA04, " & _
                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='05')),0) + " & _
                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='05')),0)) AS MARZA05, " & _
                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='06')),0) + " & _
                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='06')),0)) AS MARZA06, " & _
                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='07')),0) + " & _
                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='07')),0)) AS MARZA07, " & _
                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='08')),0) + " & _
                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='08')),0)) AS MARZA08, " & _
                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='09')),0) + " & _
                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='09')),0)) AS MARZA09, " & _
                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='10')),0) + " & _
                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='10')),0)) AS MARZA10, " & _
                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='11')),0) + " & _
                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='11')),0)) AS MARZA11, " & _
                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,4)='" & pBiezRok & "') AND (SUBSTRING(Obroty.DATA,6,2)<='12')),0) + " & _
                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS MARM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,4)='" & pBiezRok & "') AND (SUBSTRING(ObrotyMies.MIESIAC,6,2)<='12')),0)) AS MARZA12 " & _
                        " FROM Klienci) AS KL " & _
                        " WHERE (KL.SPRZEDAZPOPROK>" & WartGraniczna.ToString("F0") & ") OR ((KL.SPRZEDAZPOPROK=0) AND (KL.ROKOWANIABR<>'')) ORDER BY KL.SPRZEDAZ ASC"

        DataSetKlienci = New DataSet
        DataViewKlienci = New DataView

        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnectionKlienci = New OleDb.OleDbConnection(sConnectionKlienci)
            'dbCommandSelectKlienci = New OleDb.OleDbCommand(dbSelectSQLKlienci, dbConnectionKlienci)
            ''dbCommandDeleteKlienci = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnectionKlienci)
            ''dbCommandUpdateKlienci = New OleDb.OleDbCommand(sUpdateSQLKlienci, dbConnectionKlienci)
            ''dbCommandInsertKlienci = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnectionKlienci)
            'dbDataAdapterKlienci = New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)

            'Dim DBMapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            'DBMapowanieTabeliKlienci = dbDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionKlienci.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnectionKlienci.State
            'End Try
        Else
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(dbSelectSQLKlienci, sqlConnectionKlienci)
            'sqlCommandDeleteKlienci = New sqlclient.sqlCommand(sDeleteSQLKlienci, sqlconnectionKlienci)
            'sqlCommandUpdateKlienci = New sqlclient.sqlCommand(sUpdateSQLKlienci, sqlconnectionKlienci)
            'sqlCommandInsertKlienci = New sqlclient.sqlCommand(sInsertSQLKlienci, sqlconnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

            Dim sqlMapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliKlienci = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionKlienci.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    dbDataAdapterKlienci.Fill(DataSetKlienci)
                    dbConnectionKlienci.Close()
                Else
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                End If
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}

                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                DataViewKlienci.AllowDelete = False
                DataViewKlienci.AllowNew = False

                If DataViewKlienci.Count > 0 Then
                    InicjujRoboczy(1, "OBRÓT w grupie klientów >WGPR")
                    For i = 0 To DataViewKlienci.Count - 1
                        drv = DataViewKlienci.Item(i)
                        If drv("SPRZEDAZPOPROK") > 0 Then
                            rWykonanie += drv("SPRZEDAZ")
                            rWykonanie01 += drv("SPRZEDAZ01")
                            rWykonanie02 += drv("SPRZEDAZ02")
                            rWykonanie03 += drv("SPRZEDAZ03")
                            rWykonanie04 += drv("SPRZEDAZ04")
                            rWykonanie05 += drv("SPRZEDAZ05")
                            rWykonanie06 += drv("SPRZEDAZ06")
                            rWykonanie07 += drv("SPRZEDAZ07")
                            rWykonanie08 += drv("SPRZEDAZ08")
                            rWykonanie09 += drv("SPRZEDAZ09")
                            rWykonanie10 += drv("SPRZEDAZ10")
                            rWykonanie11 += drv("SPRZEDAZ11")
                            rWykonanie12 += drv("SPRZEDAZ12")

                            rPlan += 0
                            rPlan01 += 0
                            rPlan02 += 0
                            rPlan03 += 0
                            rPlan04 += 0
                            rPlan05 += 0
                            rPlan06 += 0
                            rPlan07 += 0
                            rPlan08 += 0
                            rPlan09 += 0
                            rPlan10 += 0
                            rPlan11 += 0
                            rPlan12 += 0

                            rPopRok += drv("SPRZEDAZPOPROK")
                            rPopRok01 += drv("SPRZEDAZPOPROK01")
                            rPopRok02 += drv("SPRZEDAZPOPROK02")
                            rPopRok03 += drv("SPRZEDAZPOPROK03")
                            rPopRok04 += drv("SPRZEDAZPOPROK04")
                            rPopRok05 += drv("SPRZEDAZPOPROK05")
                            rPopRok06 += drv("SPRZEDAZPOPROK06")
                            rPopRok07 += drv("SPRZEDAZPOPROK07")
                            rPopRok08 += drv("SPRZEDAZPOPROK08")
                            rPopRok09 += drv("SPRZEDAZPOPROK09")
                            rPopRok10 += drv("SPRZEDAZPOPROK10")
                            rPopRok11 += drv("SPRZEDAZPOPROK11")
                            rPopRok12 += drv("SPRZEDAZPOPROK12")
                        End If
                    Next
                    ZapiszRoboczy()
                    '----------------------------------
                    InicjujRoboczy(2, "OBRÓT nowych klientów rokuj¹cych")
                    For i = 0 To DataViewKlienci.Count - 1
                        drv = DataViewKlienci.Item(i)
                        If drv("SPRZEDAZPOPROK") = 0 Then
                            rWykonanie += drv("SPRZEDAZ")
                            rWykonanie01 += drv("SPRZEDAZ01")
                            rWykonanie02 += drv("SPRZEDAZ02")
                            rWykonanie03 += drv("SPRZEDAZ03")
                            rWykonanie04 += drv("SPRZEDAZ04")
                            rWykonanie05 += drv("SPRZEDAZ05")
                            rWykonanie06 += drv("SPRZEDAZ06")
                            rWykonanie07 += drv("SPRZEDAZ07")
                            rWykonanie08 += drv("SPRZEDAZ08")
                            rWykonanie09 += drv("SPRZEDAZ09")
                            rWykonanie10 += drv("SPRZEDAZ10")
                            rWykonanie11 += drv("SPRZEDAZ11")
                            rWykonanie12 += drv("SPRZEDAZ12")

                            rPlan += 0
                            rPlan01 += 0
                            rPlan02 += 0
                            rPlan03 += 0
                            rPlan04 += 0
                            rPlan05 += 0
                            rPlan06 += 0
                            rPlan07 += 0
                            rPlan08 += 0
                            rPlan09 += 0
                            rPlan10 += 0
                            rPlan11 += 0
                            rPlan12 += 0

                            rPopRok += drv("SPRZEDAZPOPROK")
                            rPopRok01 += drv("SPRZEDAZPOPROK01")
                            rPopRok02 += drv("SPRZEDAZPOPROK02")
                            rPopRok03 += drv("SPRZEDAZPOPROK03")
                            rPopRok04 += drv("SPRZEDAZPOPROK04")
                            rPopRok05 += drv("SPRZEDAZPOPROK05")
                            rPopRok06 += drv("SPRZEDAZPOPROK06")
                            rPopRok07 += drv("SPRZEDAZPOPROK07")
                            rPopRok08 += drv("SPRZEDAZPOPROK08")
                            rPopRok09 += drv("SPRZEDAZPOPROK09")
                            rPopRok10 += drv("SPRZEDAZPOPROK10")
                            rPopRok11 += drv("SPRZEDAZPOPROK11")
                            rPopRok12 += drv("SPRZEDAZPOPROK12")
                        End If
                    Next
                    ZapiszRoboczy()
                    InicjujRoboczy(3, "MAR¯A w grupie klientów >WGPR")
                    For i = 0 To DataViewKlienci.Count - 1
                        drv = DataViewKlienci.Item(i)
                        If drv("SPRZEDAZPOPROK") > 0 Then
                            rWykonanie += drv("MARZA")
                            rWykonanie01 += drv("MARZA01")
                            rWykonanie02 += drv("MARZA02")
                            rWykonanie03 += drv("MARZA03")
                            rWykonanie04 += drv("MARZA04")
                            rWykonanie05 += drv("MARZA05")
                            rWykonanie06 += drv("MARZA06")
                            rWykonanie07 += drv("MARZA07")
                            rWykonanie08 += drv("MARZA08")
                            rWykonanie09 += drv("MARZA09")
                            rWykonanie10 += drv("MARZA10")
                            rWykonanie11 += drv("MARZA11")
                            rWykonanie12 += drv("MARZA12")

                            rPlan += 0
                            rPlan01 += 0
                            rPlan02 += 0
                            rPlan03 += 0
                            rPlan04 += 0
                            rPlan05 += 0
                            rPlan06 += 0
                            rPlan07 += 0
                            rPlan08 += 0
                            rPlan09 += 0
                            rPlan10 += 0
                            rPlan11 += 0
                            rPlan12 += 0

                            rPopRok += drv("MARZAPOPROK")
                            rPopRok01 += drv("MARZAPOPROK01")
                            rPopRok02 += drv("MARZAPOPROK02")
                            rPopRok03 += drv("MARZAPOPROK03")
                            rPopRok04 += drv("MARZAPOPROK04")
                            rPopRok05 += drv("MARZAPOPROK05")
                            rPopRok06 += drv("MARZAPOPROK06")
                            rPopRok07 += drv("MARZAPOPROK07")
                            rPopRok08 += drv("MARZAPOPROK08")
                            rPopRok09 += drv("MARZAPOPROK09")
                            rPopRok10 += drv("MARZAPOPROK10")
                            rPopRok11 += drv("MARZAPOPROK11")
                            rPopRok12 += drv("MARZAPOPROK12")
                        End If
                    Next
                    ZapiszRoboczy()
                    '----------------------------------
                    InicjujRoboczy(11, "LICZBA klientów >WGPR")
                    For i = 0 To DataViewKlienci.Count - 1
                        drv = DataViewKlienci.Item(i)
                        If drv("SPRZEDAZPOPROK") > 0 Then
                            rWykonanie += IIf(drv("SPRZEDAZ") > 0, 1, 0)
                            rWykonanie01 += IIf(drv("SPRZEDAZ01") > 0, 1, 0)
                            rWykonanie02 += IIf(drv("SPRZEDAZ02") > 0, 1, 0)
                            rWykonanie03 += IIf(drv("SPRZEDAZ03") > 0, 1, 0)
                            rWykonanie04 += IIf(drv("SPRZEDAZ04") > 0, 1, 0)
                            rWykonanie05 += IIf(drv("SPRZEDAZ05") > 0, 1, 0)
                            rWykonanie06 += IIf(drv("SPRZEDAZ06") > 0, 1, 0)
                            rWykonanie07 += IIf(drv("SPRZEDAZ07") > 0, 1, 0)
                            rWykonanie08 += IIf(drv("SPRZEDAZ08") > 0, 1, 0)
                            rWykonanie09 += IIf(drv("SPRZEDAZ09") > 0, 1, 0)
                            rWykonanie10 += IIf(drv("SPRZEDAZ10") > 0, 1, 0)
                            rWykonanie11 += IIf(drv("SPRZEDAZ11") > 0, 1, 0)
                            rWykonanie12 += IIf(drv("SPRZEDAZ12") > 0, 1, 0)

                            rPlan += 0
                            rPlan01 += 0
                            rPlan02 += 0
                            rPlan03 += 0
                            rPlan04 += 0
                            rPlan05 += 0
                            rPlan06 += 0
                            rPlan07 += 0
                            rPlan08 += 0
                            rPlan09 += 0
                            rPlan10 += 0
                            rPlan11 += 0
                            rPlan12 += 0

                            rPopRok += IIf(drv("SPRZEDAZPOPROK") > 0, 1, 0)
                            rPopRok01 += IIf(drv("SPRZEDAZPOPROK01") > 0, 1, 0)
                            rPopRok02 += IIf(drv("SPRZEDAZPOPROK02") > 0, 1, 0)
                            rPopRok03 += IIf(drv("SPRZEDAZPOPROK03") > 0, 1, 0)
                            rPopRok04 += IIf(drv("SPRZEDAZPOPROK04") > 0, 1, 0)
                            rPopRok05 += IIf(drv("SPRZEDAZPOPROK05") > 0, 1, 0)
                            rPopRok06 += IIf(drv("SPRZEDAZPOPROK06") > 0, 1, 0)
                            rPopRok07 += IIf(drv("SPRZEDAZPOPROK07") > 0, 1, 0)
                            rPopRok08 += IIf(drv("SPRZEDAZPOPROK08") > 0, 1, 0)
                            rPopRok09 += IIf(drv("SPRZEDAZPOPROK09") > 0, 1, 0)
                            rPopRok10 += IIf(drv("SPRZEDAZPOPROK10") > 0, 1, 0)
                            rPopRok11 += IIf(drv("SPRZEDAZPOPROK11") > 0, 1, 0)
                            rPopRok12 += IIf(drv("SPRZEDAZPOPROK12") > 0, 1, 0)

                        End If
                    Next
                    ZapiszRoboczy()
                    '----------------------------------
                    InicjujRoboczy(12, "LICZBA nowych klientów rokuj¹cych")
                    For i = 0 To DataViewKlienci.Count - 1
                        drv = DataViewKlienci.Item(i)
                        If drv("SPRZEDAZPOPROK") = 0 Then
                            rWykonanie += IIf(drv("SPRZEDAZ") > 0, 1, 0)
                            rWykonanie01 += IIf(drv("SPRZEDAZ01") > 0, 1, 0)
                            rWykonanie02 += IIf(drv("SPRZEDAZ02") > 0, 1, 0)
                            rWykonanie03 += IIf(drv("SPRZEDAZ03") > 0, 1, 0)
                            rWykonanie04 += IIf(drv("SPRZEDAZ04") > 0, 1, 0)
                            rWykonanie05 += IIf(drv("SPRZEDAZ05") > 0, 1, 0)
                            rWykonanie06 += IIf(drv("SPRZEDAZ06") > 0, 1, 0)
                            rWykonanie07 += IIf(drv("SPRZEDAZ07") > 0, 1, 0)
                            rWykonanie08 += IIf(drv("SPRZEDAZ08") > 0, 1, 0)
                            rWykonanie09 += IIf(drv("SPRZEDAZ09") > 0, 1, 0)
                            rWykonanie10 += IIf(drv("SPRZEDAZ10") > 0, 1, 0)
                            rWykonanie11 += IIf(drv("SPRZEDAZ11") > 0, 1, 0)
                            rWykonanie12 += IIf(drv("SPRZEDAZ12") > 0, 1, 0)

                            rPlan += 0
                            rPlan01 += 0
                            rPlan02 += 0
                            rPlan03 += 0
                            rPlan04 += 0
                            rPlan05 += 0
                            rPlan06 += 0
                            rPlan07 += 0
                            rPlan08 += 0
                            rPlan09 += 0
                            rPlan10 += 0
                            rPlan11 += 0
                            rPlan12 += 0

                            rPopRok += IIf(drv("SPRZEDAZPOPROK") > 0, 1, 0)
                            rPopRok01 += IIf(drv("SPRZEDAZPOPROK01") > 0, 1, 0)
                            rPopRok02 += IIf(drv("SPRZEDAZPOPROK02") > 0, 1, 0)
                            rPopRok03 += IIf(drv("SPRZEDAZPOPROK03") > 0, 1, 0)
                            rPopRok04 += IIf(drv("SPRZEDAZPOPROK04") > 0, 1, 0)
                            rPopRok05 += IIf(drv("SPRZEDAZPOPROK05") > 0, 1, 0)
                            rPopRok06 += IIf(drv("SPRZEDAZPOPROK06") > 0, 1, 0)
                            rPopRok07 += IIf(drv("SPRZEDAZPOPROK07") > 0, 1, 0)
                            rPopRok08 += IIf(drv("SPRZEDAZPOPROK08") > 0, 1, 0)
                            rPopRok09 += IIf(drv("SPRZEDAZPOPROK09") > 0, 1, 0)
                            rPopRok10 += IIf(drv("SPRZEDAZPOPROK10") > 0, 1, 0)
                            rPopRok11 += IIf(drv("SPRZEDAZPOPROK11") > 0, 1, 0)
                            rPopRok12 += IIf(drv("SPRZEDAZPOPROK12") > 0, 1, 0)

                        End If
                    Next
                    ZapiszRoboczy()
                    '----------------------------------
                    InicjujRoboczy(13, "LICZBA klientów kupuj¹cych produkty L")
                    For i = 0 To DataViewKlienci.Count - 1
                        drv = DataViewKlienci.Item(i)
                        If drv("SPRZEDAZPOPROK") > 0 Then
                            rWykonanie += IIf(drv("SPRZEDAZL") > 0, 1, 0)
                            rWykonanie01 += IIf(drv("SPRZEDAZL01") > 0, 1, 0)
                            rWykonanie02 += IIf(drv("SPRZEDAZL02") > 0, 1, 0)
                            rWykonanie03 += IIf(drv("SPRZEDAZL03") > 0, 1, 0)
                            rWykonanie04 += IIf(drv("SPRZEDAZL04") > 0, 1, 0)
                            rWykonanie05 += IIf(drv("SPRZEDAZL05") > 0, 1, 0)
                            rWykonanie06 += IIf(drv("SPRZEDAZL06") > 0, 1, 0)
                            rWykonanie07 += IIf(drv("SPRZEDAZL07") > 0, 1, 0)
                            rWykonanie08 += IIf(drv("SPRZEDAZL08") > 0, 1, 0)
                            rWykonanie09 += IIf(drv("SPRZEDAZL09") > 0, 1, 0)
                            rWykonanie10 += IIf(drv("SPRZEDAZL10") > 0, 1, 0)
                            rWykonanie11 += IIf(drv("SPRZEDAZL11") > 0, 1, 0)
                            rWykonanie12 += IIf(drv("SPRZEDAZL12") > 0, 1, 0)

                            rPlan += 0
                            rPlan01 += 0
                            rPlan02 += 0
                            rPlan03 += 0
                            rPlan04 += 0
                            rPlan05 += 0
                            rPlan06 += 0
                            rPlan07 += 0
                            rPlan08 += 0
                            rPlan09 += 0
                            rPlan10 += 0
                            rPlan11 += 0
                            rPlan12 += 0

                            rPopRok += IIf(drv("SPRZEDAZLPOPROK") > 0, 1, 0)
                            rPopRok01 += IIf(drv("SPRZEDAZLPOPROK01") > 0, 1, 0)
                            rPopRok02 += IIf(drv("SPRZEDAZLPOPROK02") > 0, 1, 0)
                            rPopRok03 += IIf(drv("SPRZEDAZLPOPROK03") > 0, 1, 0)
                            rPopRok04 += IIf(drv("SPRZEDAZLPOPROK04") > 0, 1, 0)
                            rPopRok05 += IIf(drv("SPRZEDAZLPOPROK05") > 0, 1, 0)
                            rPopRok06 += IIf(drv("SPRZEDAZLPOPROK06") > 0, 1, 0)
                            rPopRok07 += IIf(drv("SPRZEDAZLPOPROK07") > 0, 1, 0)
                            rPopRok08 += IIf(drv("SPRZEDAZLPOPROK08") > 0, 1, 0)
                            rPopRok09 += IIf(drv("SPRZEDAZLPOPROK09") > 0, 1, 0)
                            rPopRok10 += IIf(drv("SPRZEDAZLPOPROK10") > 0, 1, 0)
                            rPopRok11 += IIf(drv("SPRZEDAZLPOPROK11") > 0, 1, 0)
                            rPopRok12 += IIf(drv("SPRZEDAZLPOPROK12") > 0, 1, 0)
                        End If
                    Next
                    ZapiszRoboczy()
                    '----------------------------------
                    InicjujRoboczy(14, "LICZBA klientów kupuj¹cych produkty A")
                    For i = 0 To DataViewKlienci.Count - 1
                        drv = DataViewKlienci.Item(i)
                        If drv("SPRZEDAZPOPROK") > 0 Then
                            rWykonanie += IIf(drv("SPRZEDAZA") > 0, 1, 0)
                            rWykonanie01 += IIf(drv("SPRZEDAZA01") > 0, 1, 0)
                            rWykonanie02 += IIf(drv("SPRZEDAZA02") > 0, 1, 0)
                            rWykonanie03 += IIf(drv("SPRZEDAZA03") > 0, 1, 0)
                            rWykonanie04 += IIf(drv("SPRZEDAZA04") > 0, 1, 0)
                            rWykonanie05 += IIf(drv("SPRZEDAZA05") > 0, 1, 0)
                            rWykonanie06 += IIf(drv("SPRZEDAZA06") > 0, 1, 0)
                            rWykonanie07 += IIf(drv("SPRZEDAZA07") > 0, 1, 0)
                            rWykonanie08 += IIf(drv("SPRZEDAZA08") > 0, 1, 0)
                            rWykonanie09 += IIf(drv("SPRZEDAZA09") > 0, 1, 0)
                            rWykonanie10 += IIf(drv("SPRZEDAZA10") > 0, 1, 0)
                            rWykonanie11 += IIf(drv("SPRZEDAZA11") > 0, 1, 0)
                            rWykonanie12 += IIf(drv("SPRZEDAZA12") > 0, 1, 0)

                            rPlan += 0
                            rPlan01 += 0
                            rPlan02 += 0
                            rPlan03 += 0
                            rPlan04 += 0
                            rPlan05 += 0
                            rPlan06 += 0
                            rPlan07 += 0
                            rPlan08 += 0
                            rPlan09 += 0
                            rPlan10 += 0
                            rPlan11 += 0
                            rPlan12 += 0

                            rPopRok += IIf(drv("SPRZEDAZAPOPROK") > 0, 1, 0)
                            rPopRok01 += IIf(drv("SPRZEDAZAPOPROK01") > 0, 1, 0)
                            rPopRok02 += IIf(drv("SPRZEDAZAPOPROK02") > 0, 1, 0)
                            rPopRok03 += IIf(drv("SPRZEDAZAPOPROK03") > 0, 1, 0)
                            rPopRok04 += IIf(drv("SPRZEDAZAPOPROK04") > 0, 1, 0)
                            rPopRok05 += IIf(drv("SPRZEDAZAPOPROK05") > 0, 1, 0)
                            rPopRok06 += IIf(drv("SPRZEDAZAPOPROK06") > 0, 1, 0)
                            rPopRok07 += IIf(drv("SPRZEDAZAPOPROK07") > 0, 1, 0)
                            rPopRok08 += IIf(drv("SPRZEDAZAPOPROK08") > 0, 1, 0)
                            rPopRok09 += IIf(drv("SPRZEDAZAPOPROK09") > 0, 1, 0)
                            rPopRok10 += IIf(drv("SPRZEDAZAPOPROK10") > 0, 1, 0)
                            rPopRok11 += IIf(drv("SPRZEDAZAPOPROK11") > 0, 1, 0)
                            rPopRok12 += IIf(drv("SPRZEDAZAPOPROK12") > 0, 1, 0)
                        End If
                    Next
                    ZapiszRoboczy()
                    '----------------------------------


                End If

            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)

            End Try
        End If









        ' ''---------------------------------------------------------------------
        ' ''Kontakty HANDLOWE - nowe 05.09.2015
        ' ''-----------------------------------------
        ''Public oUniqIdKon As String           'UNIQID            varchar(40)
        ''Public oIdentKon As String            'IDENTKLIENTA     Text, 6
        ''Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
        ''Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
        ''Public oTematKon As String            'TEMAT            Text, 50
        ''Public oUwagiKon As String            'UWAGI            Memo
        ''Public oWersjaKon As Integer          'WERSJA           Integer

        'lblRaport.Text = "3/4 Analizujê Kontakty"
        'System.Windows.Forms.Application.DoEvents()
        ''dbSelectSQLKontakty = sSelectSQLKontakty & " WHERE (SUBSTRING(DATAKONTAKTU,1,4)='" & pBiezRok & "') OR (SUBSTRING(DATAKONTAKTU,1,4)='" & pPopRok & "') "
        'dbSelectSQLKontakty = "SELECT * FROM " & _
        '                      "(SELECT IDENTKLIENTA, SUBSTRING(SKUPPLAN,1,4) AS WERYFPOTENC, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pPopRok & "')) AS ILEKONPOPROK, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "')) AS ILEKON, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='01')) AS ILEKON01, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='02')) AS ILEKON02, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='03')) AS ILEKON03, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='04')) AS ILEKON04, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='05')) AS ILEKON05, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='06')) AS ILEKON06, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='07')) AS ILEKON07, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='08')) AS ILEKON08, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='09')) AS ILEKON09, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='10')) AS ILEKON10, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='11')) AS ILEKON11, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='12')) AS ILEKON12, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pPopRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%')) AS ILEKONPOPROKTEL, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%')) AS ILEKONTEL, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='01')) AS ILEKONTEL01, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='02')) AS ILEKONTEL02, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='03')) AS ILEKONTEL03, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='04')) AS ILEKONTEL04, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='05')) AS ILEKONTEL05, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='06')) AS ILEKONTEL06, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='07')) AS ILEKONTEL07, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='08')) AS ILEKONTEL08, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='09')) AS ILEKONTEL09, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='10')) AS ILEKONTEL10, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='11')) AS ILEKONTEL11, " & _
        '                             "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='12')) AS ILEKONTEL12, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pPopRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%')) AS ILEKONPOPROKRZU, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%')) AS ILEKONRZU, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='01')) AS ILEKONRZU01, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='02')) AS ILEKONRZU02, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='03')) AS ILEKONRZU03, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='04')) AS ILEKONRZU04, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='05')) AS ILEKONRZU05, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='06')) AS ILEKONRZU06, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='07')) AS ILEKONRZU07, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='08')) AS ILEKONRZU08, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='09')) AS ILEKONRZU09, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='10')) AS ILEKONRZU10, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='11')) AS ILEKONRZU11, " & _
        '                               "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='12')) AS ILEKONRZU12, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pPopRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%')) AS ILEKONPOPROKTEST, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%')) AS ILEKONTEST, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='01')) AS ILEKONTEST01, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='02')) AS ILEKONTEST02, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='03')) AS ILEKONTEST03, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='04')) AS ILEKONTEST04, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='05')) AS ILEKONTEST05, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='06')) AS ILEKONTEST06, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='07')) AS ILEKONTEST07, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='08')) AS ILEKONTEST08, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='09')) AS ILEKONTEST09, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='10')) AS ILEKONTEST10, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='11')) AS ILEKONTEST11, " & _
        '                                 "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,4)='" & pBiezRok & "') AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST%') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,6,2)<='12')) AS ILEKONTEST12 " & _
        '                " FROM Klienci) AS KL " & _
        '                " WHERE (KL.ILEKON>0) "


        'DataSetKontakty = New DataSet
        'DataViewKontakty = New DataView
        'DataSetKontakty.Locale = New System.Globalization.CultureInfo("pl-PL")

        'If ParL_DataBaseType = "MS ACCESS" Then
        '    dbConnectionKontakty = New OleDb.OleDbConnection(sConnectionKontakty)
        '    dbCommandSelectKontakty = New OleDb.OleDbCommand(dbSelectSQLKontakty, dbConnectionKontakty)
        '    'dbCommandDeleteKontakty = New OleDb.OleDbCommand(sDeleteSQLKontakty, dbConnectionKontakty)
        '    'dbCommandUpdateKontakty = New OleDb.OleDbCommand(sUpdateSQLKontakty, dbConnectionKontakty)
        '    'dbCommandInsertKontakty = New OleDb.OleDbCommand(sInsertSQLKontakty, dbConnectionKontakty)
        '    dbDataAdapterKontakty = New OleDb.OleDbDataAdapter(dbCommandSelectKontakty)

        '    Dim DBMapowanieTabeliKontakty As System.Data.Common.DataTableMapping
        '    DBMapowanieTabeliKontakty = dbDataAdapterKontakty.TableMappings.Add("Table", "TABELA_Kontakty")

        '    '------------------------------------------
        '    'wypelnij DATASET
        '    Try
        '        dbConnectionKontakty.Open()
        '    Catch Ex As System.Exception
        '        ViewInLog(Ex, Me.Name(), Nothing)
        '        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
        '        '    System.Windows.Forms.MessageBoxButtons.OK, _
        '        '    MessageBoxIcon.Information, _
        '        '    MessageBoxDefaultButton.Button1)
        '    Finally
        '        ConnectionState = dbConnectionKontakty.State
        '    End Try
        'Else
        '    sqlConnectionKontakty = New SqlClient.SqlConnection(sConnectionKontakty)
        '    sqlCommandSelectKontakty = New SqlClient.SqlCommand(dbSelectSQLKontakty, sqlConnectionKontakty)
        '    'sqlCommandDeleteKontakty = New sqlclient.sqlCommand(sDeleteSQLKontakty, sqlconnectionKontakty)
        '    'sqlCommandUpdateKontakty = New sqlclient.sqlCommand(sUpdateSQLKontakty, sqlconnectionKontakty)
        '    'sqlCommandInsertKontakty = New sqlclient.sqlCommand(sInsertSQLKontakty, sqlconnectionKontakty)
        '    sqlDataAdapterKontakty = New SqlClient.SqlDataAdapter(sqlCommandSelectKontakty)

        '    Dim sqlMapowanieTabeliKontakty As System.Data.Common.DataTableMapping
        '    sqlMapowanieTabeliKontakty = sqlDataAdapterKontakty.TableMappings.Add("Table", "TABELA_Kontakty")

        '    '------------------------------------------
        '    'wypelnij DATASET
        '    Try
        '        sqlConnectionKontakty.Open()
        '    Catch Ex As System.Exception
        '        ViewInLog(Ex, Me.Name(), Nothing)
        '        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
        '        '    System.Windows.Forms.MessageBoxButtons.OK, _
        '        '    MessageBoxIcon.Information, _
        '        '    MessageBoxDefaultButton.Button1)
        '    Finally
        '        ConnectionState = sqlConnectionKontakty.State
        '    End Try
        'End If

        'If ConnectionState = ConnectionState.Open Then
        '    Try
        '        If ParL_DataBaseType = "MS ACCESS" Then
        '            dbDataAdapterKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
        '            dbDataAdapterKontakty.Fill(DataSetKontakty)
        '            dbConnectionKontakty.Close()
        '        Else
        '            sqlDataAdapterKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
        '            sqlDataAdapterKontakty.Fill(DataSetKontakty)
        '            sqlConnectionKontakty.Close()
        '        End If
        '        DataSetKontakty.Tables(0).PrimaryKey = New DataColumn() {DataSetKontakty.Tables(0).Columns("UNIQID")}

        '        DataViewKontakty = New DataView(DataSetKontakty.Tables(0))
        '        DataViewKontakty.AllowDelete = False
        '        DataViewKontakty.AllowNew = False

        '        If DataViewKontakty.Count > 0 Then
        '            InicjujRoboczy(21, "LICZBA kontaktów")
        '            For i = 0 To DataViewKontakty.Count - 1
        '                drv = DataViewKontakty.Item(i)
        '                rWykonanie += drv("ILEKON")
        '                rWykonanie01 += drv("ILEKON01")
        '                rWykonanie02 += drv("ILEKON02")
        '                rWykonanie03 += drv("ILEKON03")
        '                rWykonanie04 += drv("ILEKON04")
        '                rWykonanie05 += drv("ILEKON05")
        '                rWykonanie06 += drv("ILEKON06")
        '                rWykonanie07 += drv("ILEKON07")
        '                rWykonanie08 += drv("ILEKON08")
        '                rWykonanie09 += drv("ILEKON09")
        '                rWykonanie10 += drv("ILEKON10")
        '                rWykonanie11 += drv("ILEKON11")
        '                rWykonanie12 += drv("ILEKON12")
        '            Next
        '            ZapiszRoboczy()
        '            '----------------------------------
        '            InicjujRoboczy(22, "LICZBA kontaktów telefonicznych")
        '            For i = 0 To DataViewKontakty.Count - 1
        '                drv = DataViewKontakty.Item(i)
        '                rWykonanie += drv("ILEKONTEL")
        '                rWykonanie01 += drv("ILEKONTEL01")
        '                rWykonanie02 += drv("ILEKONTEL02")
        '                rWykonanie03 += drv("ILEKONTEL03")
        '                rWykonanie04 += drv("ILEKONTEL04")
        '                rWykonanie05 += drv("ILEKONTEL05")
        '                rWykonanie06 += drv("ILEKONTEL06")
        '                rWykonanie07 += drv("ILEKONTEL07")
        '                rWykonanie08 += drv("ILEKONTEL08")
        '                rWykonanie09 += drv("ILEKONTEL09")
        '                rWykonanie10 += drv("ILEKONTEL10")
        '                rWykonanie11 += drv("ILEKONTEL11")
        '                rWykonanie12 += drv("ILEKONTEL12")
        '            Next
        '            ZapiszRoboczy()
        '            '----------------------------------
        '            InicjujRoboczy(23, "LICZBA kontaktów z prezentacja RZU")
        '            For i = 0 To DataViewKontakty.Count - 1
        '                drv = DataViewKontakty.Item(i)
        '                rWykonanie += drv("ILEKONRZU")
        '                rWykonanie01 += drv("ILEKONRZU01")
        '                rWykonanie02 += drv("ILEKONRZU02")
        '                rWykonanie03 += drv("ILEKONRZU03")
        '                rWykonanie04 += drv("ILEKONRZU04")
        '                rWykonanie05 += drv("ILEKONRZU05")
        '                rWykonanie06 += drv("ILEKONRZU06")
        '                rWykonanie07 += drv("ILEKONRZU07")
        '                rWykonanie08 += drv("ILEKONRZU08")
        '                rWykonanie09 += drv("ILEKONRZU09")
        '                rWykonanie10 += drv("ILEKONRZU10")
        '                rWykonanie11 += drv("ILEKONRZU11")
        '                rWykonanie12 += drv("ILEKONRZU12")
        '            Next
        '            ZapiszRoboczy()
        '            '----------------------------------
        '            InicjujRoboczy(24, "LICZBA wizyt z testem us³ugi")
        '            For i = 0 To DataViewKontakty.Count - 1
        '                drv = DataViewKontakty.Item(i)
        '                rWykonanie += drv("ILEKONTEST")
        '                rWykonanie01 += drv("ILEKONTEST01")
        '                rWykonanie02 += drv("ILEKONTEST02")
        '                rWykonanie03 += drv("ILEKONTEST03")
        '                rWykonanie04 += drv("ILEKONTEST04")
        '                rWykonanie05 += drv("ILEKONTEST05")
        '                rWykonanie06 += drv("ILEKONTEST06")
        '                rWykonanie07 += drv("ILEKONTEST07")
        '                rWykonanie08 += drv("ILEKONTEST08")
        '                rWykonanie09 += drv("ILEKONTEST09")
        '                rWykonanie10 += drv("ILEKONTEST10")
        '                rWykonanie11 += drv("ILEKONTEST11")
        '                rWykonanie12 += drv("ILEKONTEST12")
        '            Next
        '            ZapiszRoboczy()
        '            '----------------------------------
        '        End If
        '        RetValue = True
        '    Catch Ex As System.Exception
        '        MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
        '            System.Windows.Forms.MessageBoxButtons.OK, _
        '            MessageBoxIcon.Information, _
        '            MessageBoxDefaultButton.Button1)

        '    End Try
        'End If

        Return RetValue
    End Function











    '**************************************************************
    '** wydruk z bazy danych
    '**************************************************************

    Dim AnalPrac As String = ""

    Private Sub cmdDrukuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDrukuj.Click
        Me.Cursor = Cursors.WaitCursor
        PrzygotujDane(cbxBiezRok.Text, lblPopRok.Text)
        Me.Cursor = Cursors.Default
        'mamy dane w DataSetRoboczy
        '--------------------------------------
        lblRaport.Text = "4/4 Generujê Raport"
        System.Windows.Forms.Application.DoEvents()
        LicznikRec = 0
        AnalPrac = ""

        ' Tworzymy dokument i do³¹czmy obs³ugê zdarzenia PrintPage
        Dim MyDoc As New Softart.myPrintDocument

        MyDoc.DocumentName = "Raport Rozk³ad Klientów i Obrotu"

        MyDoc.LineNumber = 0
        MyDoc.PageNumber = 0
        MyDoc.NextRowNumber = 0
        MyDoc.DefaultPageSettings.Landscape = True
        MyDoc.DefaultPageSettings.Margins.Bottom = cm2pts(0.5)
        MyDoc.DefaultPageSettings.Margins.Top = cm2pts(0.5)
        MyDoc.DefaultPageSettings.Margins.Left = cm2pts(0.5)
        MyDoc.DefaultPageSettings.Margins.Right = cm2pts(0.5)
        AddHandler MyDoc.PrintPage, AddressOf MyDoc_RaportRokladKlientowIObrotu
        '--------------------------------------

        Try
            'wybor drukarki...
            Dim PDialog As New PrintDialog
            PDialog.Document = MyDoc
            PDialog.AllowPrintToFile = True
            PDialog.PrintToFile = False
            PDialog.AllowSelection = True
            PDialog.AllowSomePages = True
            PDialog.ShowHelp = True
            PDialog.ShowNetwork = True
            PDialog.UseEXDialog = True
            Dim Result As DialogResult = PDialog.ShowDialog()
            ' Drukuj po nacisnieciu OK
            If Result = Windows.Forms.DialogResult.OK Then
                ' This method returns immediately, before the print job starts.
                ' The PrintPage event will fire asynchronously.
                'MyDoc.Print()

                'podiroglad wydruku
                Dim dlgPreview As New PrintPreviewDialog
                dlgPreview.Document = MyDoc
                dlgPreview.WindowState = FormWindowState.Maximized
                dlgPreview.UseAntiAlias = True  'lepsza jakosc wydruku
                dlgPreview.ShowDialog()
            End If
        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Error, _
                MessageBoxDefaultButton.Button1)
        Finally
            'MyDoc.Dispose()
        End Try
        lblRaport.Text = ""
        System.Windows.Forms.Application.DoEvents()
    End Sub






    '**************************************************************
    '** wydruk z bazy danych
    '**************************************************************

    Private Sub MyDoc_RaportRokladKlientowIObrotu(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        ' Retrieve the document that sent this event.
        ' You could store the document in a class member variable,
        ' but this approach allows you to use the same event handler
        ' to handle multiple print documents at once.
        Dim Doc As Softart.myPrintDocument = CType(sender, PrintDocument)
        Dim Text As String = ""
        Dim LiRow As Integer
        Dim drv As DataRowView = Nothing

        e.PageSettings.Landscape = True

        ' Define the font and determine the line height.
        Dim FontSize As Integer = 8
        Dim FontSizeTitle As Integer = 12
        Dim MyFontTitle As New Font("Arial", FontSizeTitle, FontStyle.Bold)
        Dim LineHeightTitle As Single = MyFontTitle.GetHeight(e.Graphics)
        Dim MyFont As New Font("Arial", FontSize, FontStyle.Regular)
        Dim LineHeight As Single = MyFont.GetHeight(e.Graphics)
        Dim DrawingPen As New Pen(Color.Black, 1)

        ' Create variables to hold position on page.
        Dim x As Single = e.MarginBounds.Left
        Dim y As Single = e.MarginBounds.Top

        Dim PageWidth As Single = cm2pts(29)      'e.MarginBounds.Width
        Dim PageHeight As Single = cm2pts(19)   'e.MarginBounds.Height
        Dim PageLeft As Single = cm2pts(0.2)       ' e.MarginBounds.Left
        Dim PageTop As Single = cm2pts(1)          'e.MarginBounds.Top

        Dim Tekst As String = ""
        Dim Drukuj As Boolean = True
        Dim xBase As Single = 0
        Dim yBase As Single = 0

        Do While True
            ' Naglowek strony
            Doc.PageNumber += 1
            y = e.MarginBounds.Top

            x = PageLeft
            LeftTxt("Punkt sprzeda¿y PRYZMAT :" & Par_IdentOddzialu, x, y, cm2pts(29), MyFont, e)
            RightTxt(Trim(Par_MiejscowoscUzytkownika) & " " & SysData(), x, y, cm2pts(29), MyFont, e)
            y += LineHeight
            LeftTxt("", x, y, cm2pts(29), MyFont, e)
            RightTxt("Str. " & Trim(Str(Doc.PageNumber)), x, y, cm2pts(29), MyFont, e)
            y += LineHeight

            Text = Doc.DocumentName
            CentrTxt(Text, x, y, cm2pts(29), MyFontTitle, e)
            y += MyFontTitle.GetHeight(e.Graphics)
            CentrTxt("Wartoœæ graniczna sprzeda¿y w " & lblPopRok.Text & " roku = " & WartGraniczna.ToString("F2"), x, y, PageWidth, MyFont, e)
            y += MyFont.GetHeight(e.Graphics)
            y += LineHeight

            'nag³ówek tabeli
            x = PageLeft
            xBase = x
            yBase = y

            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(9.0), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(11.0), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(12.5), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(14.0), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(15.5), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(17.0), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(18.5), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(20.0), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(21.5), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(23.0), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(24.5), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(26.0), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            'e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(27.5), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(29.0), CType(3 * MyFont.GetHeight(e.Graphics), Single))
            y += 0.5 * MyFont.GetHeight(e.Graphics)

            CentrTxt("WskaŸniki", x, y, cm2pts(9), MyFont, e)
            CentrTxt("Suma", x + cm2pts(9), y, cm2pts(2.0), MyFont, e)
            CentrTxt("Narastaj¹co w poszczególnych miesi¹cach", x + cm2pts(11.0), y, cm2pts(18), MyFont, e)
            y += MyFont.GetHeight(e.Graphics)

            CentrTxt("", x, y, cm2pts(9), MyFont, e)
            CentrTxt("", x + cm2pts(9), y, cm2pts(2.0), MyFont, e)
            CentrTxt("01", x + cm2pts(11.0), y, cm2pts(1.5), MyFont, e)
            CentrTxt("02", x + cm2pts(12.5), y, cm2pts(1.5), MyFont, e)
            CentrTxt("03", x + cm2pts(14.0), y, cm2pts(1.5), MyFont, e)
            CentrTxt("04", x + cm2pts(15.5), y, cm2pts(1.5), MyFont, e)
            CentrTxt("05", x + cm2pts(17.0), y, cm2pts(1.5), MyFont, e)
            CentrTxt("06", x + cm2pts(18.5), y, cm2pts(1.5), MyFont, e)
            CentrTxt("07", x + cm2pts(20.0), y, cm2pts(1.5), MyFont, e)
            CentrTxt("08", x + cm2pts(21.5), y, cm2pts(1.5), MyFont, e)
            CentrTxt("09", x + cm2pts(23.0), y, cm2pts(1.5), MyFont, e)
            CentrTxt("10", x + cm2pts(24.5), y, cm2pts(1.5), MyFont, e)
            CentrTxt("11", x + cm2pts(26.0), y, cm2pts(1.5), MyFont, e)
            CentrTxt("12", x + cm2pts(27.5), y, cm2pts(1.5), MyFont, e)
            y += MyFont.GetHeight(e.Graphics)

            ' petla do konca strony
            If Doc.NextRowNumber < DataViewRoboczy.Count Then
                Do
                    'ustaw kursor na wierszu DataGrid
                    LiRow = Doc.NextRowNumber
                    drv = DataViewRoboczy.Item(LiRow)


                    rPozycja = drv("POZYCJA")
                    rIdent = drv("IDENT")
                    rWykonanie = drv("WYKONANIE")
                    rWykonanie01 = drv("WYKONANIE01")
                    rWykonanie02 = drv("WYKONANIE02")
                    rWykonanie03 = drv("WYKONANIE03")
                    rWykonanie04 = drv("WYKONANIE04")
                    rWykonanie05 = drv("WYKONANIE05")
                    rWykonanie06 = drv("WYKONANIE06")
                    rWykonanie07 = drv("WYKONANIE07")
                    rWykonanie08 = drv("WYKONANIE08")
                    rWykonanie09 = drv("WYKONANIE09")
                    rWykonanie10 = drv("WYKONANIE10")
                    rWykonanie11 = drv("WYKONANIE11")
                    rWykonanie12 = drv("WYKONANIE12")

                    rPlan = drv("PLAN")
                    rPlan01 = drv("PLAN01")
                    rPlan02 = drv("PLAN02")
                    rPlan03 = drv("PLAN03")
                    rPlan04 = drv("PLAN04")
                    rPlan05 = drv("PLAN05")
                    rPlan06 = drv("PLAN06")
                    rPlan07 = drv("PLAN07")
                    rPlan08 = drv("PLAN08")
                    rPlan09 = drv("PLAN09")
                    rPlan10 = drv("PLAN10")
                    rPlan11 = drv("PLAN11")
                    rPlan12 = drv("PLAN12")

                    rPopRok = drv("POPROK")
                    rPopRok01 = drv("POPROK01")
                    rPopRok02 = drv("POPROK02")
                    rPopRok03 = drv("POPROK03")
                    rPopRok04 = drv("POPROK04")
                    rPopRok05 = drv("POPROK05")
                    rPopRok06 = drv("POPROK06")
                    rPopRok07 = drv("POPROK07")
                    rPopRok08 = drv("POPROK08")
                    rPopRok09 = drv("POPROK09")
                    rPopRok10 = drv("POPROK10")
                    rPopRok11 = drv("POPROK11")
                    rPopRok12 = drv("POPROK12")

                    If rPozycja = 1 Then
                        y += Int(0.5 * LineHeight)
                        LeftTxt("WSKANIKI SPRZEDA¯OWO-WARTOŒCIOWE", x, y, x + PageWidth, MyFont, e)
                        y += Int(1.5 * LineHeight)
                    ElseIf rPozycja = 11 Then
                        y += Int(0.5 * LineHeight)
                        e.Graphics.DrawLine(DrawingPen, x, y, x + cm2pts(29), y)
                        y += Int(0.5 * LineHeight)
                        LeftTxt("WSKANIKI SPRZEDA¯OWO-ILOŒCIOWE", x, y, x + PageWidth, MyFont, e)
                        y += Int(1.5 * LineHeight)
                    ElseIf rPozycja = 21 Then
                        y += Int(0.5 * LineHeight)
                        e.Graphics.DrawLine(DrawingPen, x, y, x + cm2pts(29), y)
                        y += Int(0.5 * LineHeight)
                        LeftTxt("WWSKANIKI AKTYWNOŒCI", x, y, x + PageWidth, MyFont, e)
                        y += Int(1.5 * LineHeight)
                    End If
                    '--------------------------------------
                    x = PageLeft 'od marginesu...
                    Doc.LineNumber += 1
                    '------------------------------------
                    'Tekst = Trim(Str(Doc.LineNumber))
                    'RightTxt(Tekst, x, y, cm2pts(0.8), MyFont, e)


                    LeftTxt(rIdent, x, y, x + PageWidth, MyFont, e)
                    RightTxt(IIf(rPlan = 0, 0.0, Math.Round(100 * rWykonanie / rPlan, 2)).ToString & " %", x + cm2pts(9), y, cm2pts(2.0), MyFont, e)
                    RightTxt(IIf(rPlan01 = 0, 0.0, Math.Round(100 * rWykonanie01 / rPlan01, 2)).ToString & " %", x + cm2pts(11.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(IIf(rPlan02 = 0, 0.0, Math.Round(100 * rWykonanie02 / rPlan02, 2)).ToString & " %", x + cm2pts(12.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(IIf(rPlan03 = 0, 0.0, Math.Round(100 * rWykonanie03 / rPlan03, 2)).ToString & " %", x + cm2pts(14.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(IIf(rPlan04 = 0, 0.0, Math.Round(100 * rWykonanie04 / rPlan04, 2)).ToString & " %", x + cm2pts(15.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(IIf(rPlan05 = 0, 0.0, Math.Round(100 * rWykonanie05 / rPlan05, 2)).ToString & " %", x + cm2pts(17.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(IIf(rPlan06 = 0, 0.0, Math.Round(100 * rWykonanie06 / rPlan06, 2)).ToString & " %", x + cm2pts(18.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(IIf(rPlan07 = 0, 0.0, Math.Round(100 * rWykonanie07 / rPlan07, 2)).ToString & " %", x + cm2pts(20.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(IIf(rPlan08 = 0, 0.0, Math.Round(100 * rWykonanie08 / rPlan08, 2)).ToString & " %", x + cm2pts(21.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(IIf(rPlan09 = 0, 0.0, Math.Round(100 * rWykonanie09 / rPlan09, 2)).ToString & " %", x + cm2pts(23.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(IIf(rPlan10 = 0, 0.0, Math.Round(100 * rWykonanie10 / rPlan10, 2)).ToString & " %", x + cm2pts(24.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(IIf(rPlan11 = 0, 0.0, Math.Round(100 * rWykonanie11 / rPlan11, 2)).ToString & " %", x + cm2pts(26.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(IIf(rPlan12 = 0, 0.0, Math.Round(100 * rWykonanie12 / rPlan12, 2)).ToString & " %", x + cm2pts(27.5), y, cm2pts(1.5), MyFont, e)
                    y += LineHeight
                    LeftTxt("Realizacja w " & cbxBiezRok.Text & " r.", x, y, cm2pts(5 + 2), MyFont, e)
                    RightTxt(rWykonanie.ToString("F0"), x + cm2pts(9), y, cm2pts(2.0), MyFont, e)
                    RightTxt(rWykonanie01.ToString("F0"), x + cm2pts(11.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rWykonanie02.ToString("F0"), x + cm2pts(12.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rWykonanie03.ToString("F0"), x + cm2pts(14.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rWykonanie04.ToString("F0"), x + cm2pts(15.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rWykonanie05.ToString("F0"), x + cm2pts(17.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rWykonanie06.ToString("F0"), x + cm2pts(18.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rWykonanie07.ToString("F0"), x + cm2pts(20.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rWykonanie08.ToString("F0"), x + cm2pts(21.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rWykonanie09.ToString("F0"), x + cm2pts(23.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rWykonanie10.ToString("F0"), x + cm2pts(24.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rWykonanie11.ToString("F0"), x + cm2pts(26.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rWykonanie12.ToString("F0"), x + cm2pts(27.5), y, cm2pts(1.5), MyFont, e)
                    y += LineHeight
                    LeftTxt("Planowanie w " & cbxBiezRok.Text & " r.", x, y, cm2pts(5 + 2), MyFont, e)
                    RightTxt(rPlan.ToString("F0"), x + cm2pts(9), y, cm2pts(2.0), MyFont, e)
                    RightTxt(rPlan01.ToString("F0"), x + cm2pts(11.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPlan02.ToString("F0"), x + cm2pts(12.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPlan03.ToString("F0"), x + cm2pts(14.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPlan04.ToString("F0"), x + cm2pts(15.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPlan05.ToString("F0"), x + cm2pts(17.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPlan06.ToString("F0"), x + cm2pts(18.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPlan07.ToString("F0"), x + cm2pts(20.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPlan08.ToString("F0"), x + cm2pts(21.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPlan09.ToString("F0"), x + cm2pts(23.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPlan10.ToString("F0"), x + cm2pts(24.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPlan11.ToString("F0"), x + cm2pts(26.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPlan12.ToString("F0"), x + cm2pts(27.5), y, cm2pts(1.5), MyFont, e)
                    y += LineHeight
                    LeftTxt("Realizacja w " & lblPopRok.Text & " r.", x, y, cm2pts(5 + 2), MyFont, e)
                    RightTxt(rPopRok.ToString("F0"), x + cm2pts(9), y, cm2pts(2.0), MyFont, e)
                    RightTxt(rPopRok01.ToString("F0"), x + cm2pts(11.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPopRok02.ToString("F0"), x + cm2pts(12.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPopRok03.ToString("F0"), x + cm2pts(14.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPopRok04.ToString("F0"), x + cm2pts(15.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPopRok05.ToString("F0"), x + cm2pts(17.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPopRok06.ToString("F0"), x + cm2pts(18.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPopRok07.ToString("F0"), x + cm2pts(20.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPopRok08.ToString("F0"), x + cm2pts(21.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPopRok09.ToString("F0"), x + cm2pts(23.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPopRok10.ToString("F0"), x + cm2pts(24.5), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPopRok11.ToString("F0"), x + cm2pts(26.0), y, cm2pts(1.5), MyFont, e)
                    RightTxt(rPopRok12.ToString("F0"), x + cm2pts(27.5), y, cm2pts(1.5), MyFont, e)
                    y += Int(1.5 * LineHeight)

                    Doc.NextRowNumber += 1
                Loop Until ((y + 3 * LineHeight) > e.MarginBounds.Bottom) Or (Doc.NextRowNumber >= DataViewRoboczy.Count)
            End If


            If (Doc.NextRowNumber < DataViewRoboczy.Count) Then
                ' There is still at least one more page.
                ' Signal this event to fire again.
                e.HasMorePages = True

                'e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(5.0), y - yBase)
                'e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(7.0), y - yBase)
                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(9.0), y - yBase)
                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(11.0), y - yBase)
                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(29.0), y - yBase)

            Else
                ' Printing is complete.
                'y += Int(0.5 * LineHeight)
                'e.Graphics.DrawLine(DrawingPen, x, y, x + cm2pts(29), y)
                'y += Int(0.5 * LineHeight)

                y += Int(1.0 * LineHeight)

                'e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(5.0), y - yBase)
                'e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(7.0), y - yBase)
                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(9.0), y - yBase)
                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(11.0), y - yBase)
                e.Graphics.DrawRectangle(DrawingPen, xBase, yBase, cm2pts(29.0), y - yBase)

                y += Int(0.5 * LineHeight)
                x = PageLeft 'od marginesu...
                Text = "Koniec Raportu"
                e.Graphics.DrawString(Text, MyFont, Brushes.Black, x + cm2pts(0.2), y)

                e.HasMorePages = False
            End If

            '------------------------------
            'sprawdz czy strona do wydrukowania
            If Doc.PrinterSettings.ToPage = 0 Then
                'nie ma ograniczenia stron
                Exit Do
            Else
                If Doc.PageNumber = Doc.PrinterSettings.ToPage Then
                    'ostatnia drukowana strona
                    e.HasMorePages = False
                    Exit Do
                Else
                    If Doc.PageNumber < Doc.PrinterSettings.FromPage Or Doc.PageNumber > Doc.PrinterSettings.ToPage Then
                        'nie drukuj
                        e.Graphics.Clear(Color.White)
                        If e.HasMorePages = False Then
                            Exit Do
                        End If
                    Else
                        'drukuj
                        Exit Do
                    End If
                End If
            End If
        Loop

        If e.HasMorePages = False Then
            Doc.PageNumber = 0
            Doc.LineNumber = 0
            Doc.NextRowNumber = 0

            AnalPrac = ""
            '----------------------------
        End If
    End Sub






End Class
