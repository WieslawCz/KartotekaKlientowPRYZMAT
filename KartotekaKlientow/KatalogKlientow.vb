Imports System.Drawing.Printing

Imports System.Data.Oledb
Imports System
Imports System.Reflection ' For Missing.Value and BindingFlags
Imports System.Runtime.InteropServices ' For COMException
'------------------------------------------
'UWAGA :
'Do referencji trzeba DODAC
'   Microsoft Excel 11.0 Object Library
'   Microsoft Office 11.0 Object Library
'-------------------------------------------
Imports Microsoft.Office.Interop.excel
Imports Microsoft.Office.Interop.Outlook
Imports System.ComponentModel

Public Class KatalogKlientow
    Inherits System.Windows.Forms.Form
    'Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "


    '====================================================
    'zmienne do œledzenia JAK ZAMKNIETO FORME
    Private _closeClick As Boolean
    'Private components As System.ComponentModel.Container
    Public Const SC_CLOSE As Integer = 61536
    Public Const WM_SYSCOMMAND As Integer = 274
    '====================================================



    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub


    '====================================================
    ' metoda WndProc wykorzystywana do oznaczenia sposobu zamkniecia formy
    Protected Overloads Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = WM_SYSCOMMAND AndAlso m.WParam.ToInt32 = SC_CLOSE Then
            'zamkniêto naciskaj¹c krzyzyk w prawym gornym rogu...
            Me._closeClick = True
        End If
        MyBase.WndProc(m)
    End Sub
    '====================================================



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
    Friend WithEvents cmdDrukuj As System.Windows.Forms.Button
    Friend WithEvents cmdEdytuj As System.Windows.Forms.Button
    Friend WithEvents cmdUsuñ As System.Windows.Forms.Button
    Friend WithEvents cmdDodaj As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblKod As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblMiejscowosc As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents lbleMail As System.Windows.Forms.Label
    Friend WithEvents stbKlienci As System.Windows.Forms.StatusBar
    'Friend WithEvents dagKlienci As KartotekaKlientow.Softart.MyDataGrid
    Friend WithEvents lblIdent As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents lblNrDomu As System.Windows.Forms.Label
    Friend WithEvents lblTelefon As System.Windows.Forms.Label
    Friend WithEvents lblUlica As System.Windows.Forms.Label
    Friend WithEvents lblNazwaHandlowa As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdWybierzSchemat As System.Windows.Forms.Button
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents lblRodzaj As System.Windows.Forms.Label
    Friend WithEvents cbxWydruki As System.Windows.Forms.ComboBox
    Friend WithEvents cmdWysylaj As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents lblTransakcje As System.Windows.Forms.Label
    Friend WithEvents lblNastKontakt As System.Windows.Forms.Label
    Friend WithEvents lblOstKontakt As System.Windows.Forms.Label
    Friend WithEvents lblOpiekun As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents lblPotencjal As System.Windows.Forms.Label
    Friend WithEvents stbFiltry As System.Windows.Forms.StatusBar
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdFi As System.Windows.Forms.Button
    Friend WithEvents cmdAdrExport As System.Windows.Forms.Button
    Friend WithEvents txtTOFI As System.Windows.Forms.TextBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents cmdOdswiez As System.Windows.Forms.Button
    Friend WithEvents lblIdDomu As System.Windows.Forms.Label
    Friend WithEvents cmdClearColWidth As System.Windows.Forms.Button
    Friend WithEvents HorizScrollTimer As System.Windows.Forms.Timer
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents cmdHarmonogramWizyt As System.Windows.Forms.Button
    Friend WithEvents cmdCoDalej As System.Windows.Forms.Button
    Friend WithEvents lblIlSprzetu As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents cmdDoExcela As System.Windows.Forms.Button
    Friend WithEvents dagKlienci As DataGridView
    Friend WithEvents menuEKlienci As ContextMenuStrip
    Friend WithEvents menuEEdytuj As ToolStripMenuItem
    Friend WithEvents menuEDodaj As ToolStripMenuItem
    Friend WithEvents menuEUsun As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents menuESingleL As ToolStripMenuItem
    Friend WithEvents menuEMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents menuEOdswiez As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents menuENowaA As ToolStripMenuItem
    Friend WithEvents menuEWybierzA As ToolStripMenuItem
    Friend WithEvents menuEWybierzKilkaA As ToolStripMenuItem
    Friend WithEvents menuEDodajDoA As ToolStripMenuItem
    Friend WithEvents menuEUsunZA As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents menuEWybierzWgSprzed As ToolStripMenuItem
    Friend WithEvents menuEWybierzWgRZU As ToolStripMenuItem
    Friend WithEvents menuEWybierzWgWzrostu As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents menuEUstawM As ToolStripMenuItem
    Friend WithEvents menuEUsunM As ToolStripMenuItem
    Friend WithEvents menuEZmienM As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator6 As ToolStripSeparator
    Friend WithEvents menuEUstawMP As ToolStripMenuItem
    Friend WithEvents menuEUsunMP As ToolStripMenuItem
    Friend WithEvents menuEZmienMP As ToolStripMenuItem
    Friend WithEvents Panel4 As Panel
    Friend WithEvents ToolStripSeparator7 As ToolStripSeparator
    Friend WithEvents menuZmienWskAkt As ToolStripMenuItem
    Friend WithEvents menuEWybierzWgOsoby As ToolStripMenuItem
    Friend WithEvents cmdWszystko As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogKlientow))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdDoExcela = New System.Windows.Forms.Button()
        Me.cmdCoDalej = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.cmdClearColWidth = New System.Windows.Forms.Button()
        Me.cmdOdswiez = New System.Windows.Forms.Button()
        Me.cmdAdrExport = New System.Windows.Forms.Button()
        Me.cmdDrukuj = New System.Windows.Forms.Button()
        Me.cmdWysylaj = New System.Windows.Forms.Button()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.HorizScrollTimer = New System.Windows.Forms.Timer(Me.components)
        Me.cmdHarmonogramWizyt = New System.Windows.Forms.Button()
        Me.cbxWydruki = New System.Windows.Forms.ComboBox()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.cmdWybierzSchemat = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblIlSprzetu = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.lbleMail = New System.Windows.Forms.Label()
        Me.lblTelefon = New System.Windows.Forms.Label()
        Me.lblIdDomu = New System.Windows.Forms.Label()
        Me.txtTOFI = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblPotencjal = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblTransakcje = New System.Windows.Forms.Label()
        Me.lblNastKontakt = New System.Windows.Forms.Label()
        Me.lblOstKontakt = New System.Windows.Forms.Label()
        Me.lblOpiekun = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.lblRodzaj = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa = New System.Windows.Forms.Label()
        Me.lblUlica = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblMiejscowosc = New System.Windows.Forms.Label()
        Me.lblNrDomu = New System.Windows.Forms.Label()
        Me.lblKod = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblIdent = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.stbKlienci = New System.Windows.Forms.StatusBar()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.dagKlienci = New System.Windows.Forms.DataGridView()
        Me.menuEKlienci = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuEEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEUsun = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuESingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuEOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuENowaA = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEWybierzA = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEWybierzKilkaA = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEDodajDoA = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEUsunZA = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuEWybierzWgSprzed = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEWybierzWgOsoby = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEWybierzWgRZU = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEWybierzWgWzrostu = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuEUstawM = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEUsunM = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEZmienM = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuEUstawMP = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEUsunMP = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuEZmienMP = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuZmienWskAkt = New System.Windows.Forms.ToolStripMenuItem()
        Me.Panel4 = New System.Windows.Forms.Panel()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dagKlienci, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuEKlienci.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolTip1
        '
        Me.ToolTip1.AutoPopDelay = 5000
        Me.ToolTip1.InitialDelay = 100
        Me.ToolTip1.IsBalloon = True
        Me.ToolTip1.ReshowDelay = 100
        '
        'cmdDoExcela
        '
        Me.cmdDoExcela.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdDoExcela.ForeColor = System.Drawing.Color.Black
        Me.cmdDoExcela.Image = CType(resources.GetObject("cmdDoExcela.Image"), System.Drawing.Image)
        Me.cmdDoExcela.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDoExcela.Location = New System.Drawing.Point(144, 479)
        Me.cmdDoExcela.Name = "cmdDoExcela"
        Me.cmdDoExcela.Size = New System.Drawing.Size(24, 22)
        Me.cmdDoExcela.TabIndex = 178
        Me.cmdDoExcela.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdDoExcela, "Eksport do EXCELa informacji o klientach" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "widocznych na ekranie (przefiltrowanych" &
        ").")
        '
        'cmdCoDalej
        '
        Me.cmdCoDalej.ForeColor = System.Drawing.Color.Black
        Me.cmdCoDalej.Location = New System.Drawing.Point(8, 130)
        Me.cmdCoDalej.Name = "cmdCoDalej"
        Me.cmdCoDalej.Size = New System.Drawing.Size(130, 22)
        Me.cmdCoDalej.TabIndex = 177
        Me.cmdCoDalej.Text = "Co dalej ?"
        Me.ToolTip1.SetToolTip(Me.cmdCoDalej, "Wszystkie zaplanowane dzia³ania z klientem.")
        '
        'Button2
        '
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(8, 108)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(130, 22)
        Me.Button2.TabIndex = 175
        Me.Button2.Text = "Kontakty Handlowe"
        Me.ToolTip1.SetToolTip(Me.Button2, "Wszystkie Kontakty Handlowe z Klientami.")
        '
        'cmdClearColWidth
        '
        Me.cmdClearColWidth.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdClearColWidth.ForeColor = System.Drawing.Color.Black
        Me.cmdClearColWidth.Location = New System.Drawing.Point(0, 0)
        Me.cmdClearColWidth.Name = "cmdClearColWidth"
        Me.cmdClearColWidth.Size = New System.Drawing.Size(46, 20)
        Me.cmdClearColWidth.TabIndex = 174
        Me.cmdClearColWidth.Text = "|<==>|"
        Me.ToolTip1.SetToolTip(Me.cmdClearColWidth, "Powrót do pierwotnej szerokoœci kolumn")
        '
        'cmdOdswiez
        '
        Me.cmdOdswiez.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOdswiez.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdOdswiez.ForeColor = System.Drawing.Color.Black
        Me.cmdOdswiez.Image = CType(resources.GetObject("cmdOdswiez.Image"), System.Drawing.Image)
        Me.cmdOdswiez.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOdswiez.Location = New System.Drawing.Point(799, 480)
        Me.cmdOdswiez.Name = "cmdOdswiez"
        Me.cmdOdswiez.Size = New System.Drawing.Size(92, 24)
        Me.cmdOdswiez.TabIndex = 173
        Me.cmdOdswiez.Text = "Odœwie¿"
        Me.cmdOdswiez.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdOdswiez, "Odœwie¿enie informacji." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Pobranie Kartoteki Klientów z serwera aby pokazaæ wszyst" &
        "kie" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "równolegle wykonywane modyfikacje." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10))
        '
        'cmdAdrExport
        '
        Me.cmdAdrExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAdrExport.ForeColor = System.Drawing.Color.Black
        Me.cmdAdrExport.Image = CType(resources.GetObject("cmdAdrExport.Image"), System.Drawing.Image)
        Me.cmdAdrExport.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAdrExport.Location = New System.Drawing.Point(366, 479)
        Me.cmdAdrExport.Name = "cmdAdrExport"
        Me.cmdAdrExport.Size = New System.Drawing.Size(83, 24)
        Me.cmdAdrExport.TabIndex = 172
        Me.cmdAdrExport.Text = "Export adr."
        Me.cmdAdrExport.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdAdrExport, "Eksport do EXCELa danych adresowych klientów" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "widocznych na ekranie (przefiltrowa" &
        "nych).")
        '
        'cmdDrukuj
        '
        Me.cmdDrukuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDrukuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDrukuj.ForeColor = System.Drawing.Color.Black
        Me.cmdDrukuj.Image = CType(resources.GetObject("cmdDrukuj.Image"), System.Drawing.Image)
        Me.cmdDrukuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDrukuj.Location = New System.Drawing.Point(296, 480)
        Me.cmdDrukuj.Name = "cmdDrukuj"
        Me.cmdDrukuj.Size = New System.Drawing.Size(65, 24)
        Me.cmdDrukuj.TabIndex = 44
        Me.cmdDrukuj.Text = "Drukuj"
        Me.cmdDrukuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdDrukuj, "Wydruk zestawieñ dla wszystkich widocznych" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "na ekranie klientów (przefiltrowanych" &
        ").")
        '
        'cmdWysylaj
        '
        Me.cmdWysylaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWysylaj.ForeColor = System.Drawing.Color.Black
        Me.cmdWysylaj.Image = CType(resources.GetObject("cmdWysylaj.Image"), System.Drawing.Image)
        Me.cmdWysylaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWysylaj.Location = New System.Drawing.Point(456, 479)
        Me.cmdWysylaj.Name = "cmdWysylaj"
        Me.cmdWysylaj.Size = New System.Drawing.Size(80, 24)
        Me.cmdWysylaj.TabIndex = 16
        Me.cmdWysylaj.Text = "Wyœlij"
        Me.cmdWysylaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdWysylaj, "Wysy³anie listu eMail do wszystkich klientów" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "widocznych na ekranie (przefiltrowa" &
        "nych).")
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(698, 0)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 171
        Me.ToolTip1.SetToolTip(Me.cmdFi, "Unikalne wartoœci pól tej kolumny.")
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWszystko.ForeColor = System.Drawing.Color.Black
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(730, 0)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 20)
        Me.cmdWszystko.TabIndex = 115
        Me.ToolTip1.SetToolTip(Me.cmdWszystko, "Wyczyœæ wszystkie filtry.")
        '
        'Button4
        '
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.Location = New System.Drawing.Point(8, 64)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(130, 22)
        Me.Button4.TabIndex = 64
        Me.Button4.Text = "Obroty miesiêczne"
        Me.ToolTip1.SetToolTip(Me.Button4, "Obroty miesiêczne z klientem wskazywanym przez kursor.")
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ForeColor = System.Drawing.Color.Black
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(896, 480)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(80, 24)
        Me.cmdStop.TabIndex = 38
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdStop, "Koniec dzia³ania funkcji." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Powrót do ekranu z którego wywo³ano tê funkcjê.")
        '
        'cmdEdytuj
        '
        Me.cmdEdytuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytuj.ForeColor = System.Drawing.Color.Black
        Me.cmdEdytuj.Image = CType(resources.GetObject("cmdEdytuj.Image"), System.Drawing.Image)
        Me.cmdEdytuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytuj.Location = New System.Drawing.Point(713, 480)
        Me.cmdEdytuj.Name = "cmdEdytuj"
        Me.cmdEdytuj.Size = New System.Drawing.Size(80, 24)
        Me.cmdEdytuj.TabIndex = 42
        Me.cmdEdytuj.Text = "Edytuj"
        Me.cmdEdytuj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdEdytuj, "Edycja informacji o kliencie wskazywanym przez kursor.")
        '
        'cmdUsuñ
        '
        Me.cmdUsuñ.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUsuñ.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdUsuñ.ForeColor = System.Drawing.Color.Black
        Me.cmdUsuñ.Image = CType(resources.GetObject("cmdUsuñ.Image"), System.Drawing.Image)
        Me.cmdUsuñ.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUsuñ.Location = New System.Drawing.Point(627, 480)
        Me.cmdUsuñ.Name = "cmdUsuñ"
        Me.cmdUsuñ.Size = New System.Drawing.Size(80, 24)
        Me.cmdUsuñ.TabIndex = 41
        Me.cmdUsuñ.Text = "Usuñ"
        Me.cmdUsuñ.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdUsuñ, "Usuniecie informacji o kliencie wskazywanym przez kursor" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "z Kartoteki Klientów.")
        '
        'cmdDodaj
        '
        Me.cmdDodaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodaj.ForeColor = System.Drawing.Color.Black
        Me.cmdDodaj.Image = CType(resources.GetObject("cmdDodaj.Image"), System.Drawing.Image)
        Me.cmdDodaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodaj.Location = New System.Drawing.Point(541, 480)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 24)
        Me.cmdDodaj.TabIndex = 40
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdDodaj, "Dopisanie (dodanie) nowego klienta do Kartoteki.")
        '
        'Button3
        '
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(8, 86)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(130, 22)
        Me.Button3.TabIndex = 55
        Me.Button3.Text = "Obroty ost.miesi¹ca"
        Me.ToolTip1.SetToolTip(Me.Button3, "Obroty w ostatnim miesi¹cu z klientem wskazywanym przez kursor.")
        '
        'Button1
        '
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(8, 42)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(130, 22)
        Me.Button1.TabIndex = 53
        Me.Button1.Text = "Anal.czêst.kontaktów"
        Me.ToolTip1.SetToolTip(Me.Button1, "Analiza czêstoœci kontaktów z klientami")
        '
        'HorizScrollTimer
        '
        '
        'cmdHarmonogramWizyt
        '
        Me.cmdHarmonogramWizyt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdHarmonogramWizyt.ForeColor = System.Drawing.Color.Black
        Me.cmdHarmonogramWizyt.Image = CType(resources.GetObject("cmdHarmonogramWizyt.Image"), System.Drawing.Image)
        Me.cmdHarmonogramWizyt.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdHarmonogramWizyt.Location = New System.Drawing.Point(118, 479)
        Me.cmdHarmonogramWizyt.Name = "cmdHarmonogramWizyt"
        Me.cmdHarmonogramWizyt.Size = New System.Drawing.Size(24, 22)
        Me.cmdHarmonogramWizyt.TabIndex = 176
        Me.cmdHarmonogramWizyt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbxWydruki
        '
        Me.cbxWydruki.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbxWydruki.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxWydruki.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxWydruki.ForeColor = System.Drawing.Color.Purple
        Me.cbxWydruki.Location = New System.Drawing.Point(174, 481)
        Me.cbxWydruki.Name = "cbxWydruki"
        Me.cbxWydruki.Size = New System.Drawing.Size(117, 22)
        Me.cbxWydruki.TabIndex = 58
        '
        'stbFiltry
        '
        Me.stbFiltry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltry.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltry.Location = New System.Drawing.Point(0, 0)
        Me.stbFiltry.Name = "stbFiltry"
        Me.stbFiltry.ShowPanels = True
        Me.stbFiltry.Size = New System.Drawing.Size(968, 22)
        Me.stbFiltry.TabIndex = 67
        Me.stbFiltry.Text = "StatusBar1"
        '
        'picLogo
        '
        Me.picLogo.Image = CType(resources.GetObject("picLogo.Image"), System.Drawing.Image)
        Me.picLogo.Location = New System.Drawing.Point(552, 19)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(248, 50)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 66
        Me.picLogo.TabStop = False
        Me.picLogo.Visible = False
        '
        'cmdWybierzSchemat
        '
        Me.cmdWybierzSchemat.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzSchemat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierzSchemat.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierzSchemat.Location = New System.Drawing.Point(144, 292)
        Me.cmdWybierzSchemat.Name = "cmdWybierzSchemat"
        Me.cmdWybierzSchemat.Size = New System.Drawing.Size(56, 23)
        Me.cmdWybierzSchemat.TabIndex = 57
        Me.cmdWybierzSchemat.Text = "Wybierz"
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 8)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(130, 30)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 56
        Me.PictureBox2.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.picLogo)
        Me.Panel1.Controls.Add(Me.lblIlSprzetu)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.lbleMail)
        Me.Panel1.Controls.Add(Me.lblTelefon)
        Me.Panel1.Controls.Add(Me.lblIdDomu)
        Me.Panel1.Controls.Add(Me.txtTOFI)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.lblPotencjal)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.lblTransakcje)
        Me.Panel1.Controls.Add(Me.lblNastKontakt)
        Me.Panel1.Controls.Add(Me.lblOstKontakt)
        Me.Panel1.Controls.Add(Me.lblOpiekun)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.lblRodzaj)
        Me.Panel1.Controls.Add(Me.lblNazwaHandlowa)
        Me.Panel1.Controls.Add(Me.lblUlica)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.lblMiejscowosc)
        Me.Panel1.Controls.Add(Me.lblNrDomu)
        Me.Panel1.Controls.Add(Me.lblKod)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.lblIdent)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(144, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(832, 144)
        Me.Panel1.TabIndex = 46
        '
        'lblIlSprzetu
        '
        Me.lblIlSprzetu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIlSprzetu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIlSprzetu.ForeColor = System.Drawing.Color.Purple
        Me.lblIlSprzetu.Location = New System.Drawing.Point(64, 88)
        Me.lblIlSprzetu.Name = "lblIlSprzetu"
        Me.lblIlSprzetu.Size = New System.Drawing.Size(672, 48)
        Me.lblIlSprzetu.TabIndex = 40
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Navy
        Me.Label17.Location = New System.Drawing.Point(8, 88)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(88, 16)
        Me.Label17.TabIndex = 41
        Me.Label17.Text = "Urz¹dzenia . . . . . . . . "
        '
        'lbleMail
        '
        Me.lbleMail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbleMail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lbleMail.ForeColor = System.Drawing.Color.Purple
        Me.lbleMail.Location = New System.Drawing.Point(168, 72)
        Me.lbleMail.Name = "lbleMail"
        Me.lbleMail.Size = New System.Drawing.Size(232, 16)
        Me.lbleMail.TabIndex = 20
        '
        'lblTelefon
        '
        Me.lblTelefon.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefon.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTelefon.ForeColor = System.Drawing.Color.Purple
        Me.lblTelefon.Location = New System.Drawing.Point(168, 56)
        Me.lblTelefon.Name = "lblTelefon"
        Me.lblTelefon.Size = New System.Drawing.Size(232, 16)
        Me.lblTelefon.TabIndex = 9
        '
        'lblIdDomu
        '
        Me.lblIdDomu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdDomu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdDomu.ForeColor = System.Drawing.Color.Purple
        Me.lblIdDomu.Location = New System.Drawing.Point(355, 40)
        Me.lblIdDomu.Name = "lblIdDomu"
        Me.lblIdDomu.Size = New System.Drawing.Size(45, 16)
        Me.lblIdDomu.TabIndex = 39
        '
        'txtTOFI
        '
        Me.txtTOFI.Location = New System.Drawing.Point(64, 24)
        Me.txtTOFI.Multiline = True
        Me.txtTOFI.Name = "txtTOFI"
        Me.txtTOFI.ReadOnly = True
        Me.txtTOFI.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtTOFI.Size = New System.Drawing.Size(56, 48)
        Me.txtTOFI.TabIndex = 38
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(128, 57)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 35
        Me.Label2.Text = "Telefon . . "
        '
        'lblPotencjal
        '
        Me.lblPotencjal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPotencjal.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblPotencjal.ForeColor = System.Drawing.Color.Purple
        Me.lblPotencjal.Location = New System.Drawing.Point(504, 72)
        Me.lblPotencjal.Name = "lblPotencjal"
        Me.lblPotencjal.Size = New System.Drawing.Size(232, 16)
        Me.lblPotencjal.TabIndex = 33
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(408, 72)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(104, 16)
        Me.Label8.TabIndex = 34
        Me.Label8.Text = "Potencja³ . . . . . . . . ."
        '
        'lblTransakcje
        '
        Me.lblTransakcje.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTransakcje.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTransakcje.ForeColor = System.Drawing.Color.Purple
        Me.lblTransakcje.Location = New System.Drawing.Point(504, 56)
        Me.lblTransakcje.Name = "lblTransakcje"
        Me.lblTransakcje.Size = New System.Drawing.Size(232, 16)
        Me.lblTransakcje.TabIndex = 19
        '
        'lblNastKontakt
        '
        Me.lblNastKontakt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNastKontakt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNastKontakt.ForeColor = System.Drawing.Color.Purple
        Me.lblNastKontakt.Location = New System.Drawing.Point(504, 40)
        Me.lblNastKontakt.Name = "lblNastKontakt"
        Me.lblNastKontakt.Size = New System.Drawing.Size(232, 16)
        Me.lblNastKontakt.TabIndex = 18
        '
        'lblOstKontakt
        '
        Me.lblOstKontakt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOstKontakt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOstKontakt.ForeColor = System.Drawing.Color.Purple
        Me.lblOstKontakt.Location = New System.Drawing.Point(504, 24)
        Me.lblOstKontakt.Name = "lblOstKontakt"
        Me.lblOstKontakt.Size = New System.Drawing.Size(232, 16)
        Me.lblOstKontakt.TabIndex = 17
        '
        'lblOpiekun
        '
        Me.lblOpiekun.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOpiekun.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOpiekun.ForeColor = System.Drawing.Color.Purple
        Me.lblOpiekun.Location = New System.Drawing.Point(504, 8)
        Me.lblOpiekun.Name = "lblOpiekun"
        Me.lblOpiekun.Size = New System.Drawing.Size(232, 16)
        Me.lblOpiekun.TabIndex = 16
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Navy
        Me.Label15.Location = New System.Drawing.Point(8, 72)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(56, 16)
        Me.Label15.TabIndex = 32
        Me.Label15.Text = "Rodz.sp. . . . . . . . . "
        '
        'lblRodzaj
        '
        Me.lblRodzaj.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRodzaj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblRodzaj.ForeColor = System.Drawing.Color.Purple
        Me.lblRodzaj.Location = New System.Drawing.Point(64, 72)
        Me.lblRodzaj.Name = "lblRodzaj"
        Me.lblRodzaj.Size = New System.Drawing.Size(56, 16)
        Me.lblRodzaj.TabIndex = 31
        Me.lblRodzaj.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblNazwaHandlowa
        '
        Me.lblNazwaHandlowa.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwaHandlowa.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwaHandlowa.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwaHandlowa.Location = New System.Drawing.Point(168, 8)
        Me.lblNazwaHandlowa.Name = "lblNazwaHandlowa"
        Me.lblNazwaHandlowa.Size = New System.Drawing.Size(232, 16)
        Me.lblNazwaHandlowa.TabIndex = 2
        '
        'lblUlica
        '
        Me.lblUlica.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblUlica.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblUlica.ForeColor = System.Drawing.Color.Purple
        Me.lblUlica.Location = New System.Drawing.Point(168, 40)
        Me.lblUlica.Name = "lblUlica"
        Me.lblUlica.Size = New System.Drawing.Size(144, 16)
        Me.lblUlica.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(128, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 16)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Ulica . . "
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(8, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 16)
        Me.Label5.TabIndex = 26
        Me.Label5.Text = "Nr TOFI . . . . . . . . "
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(128, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 16)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Nazwa . . . . . . . . "
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Navy
        Me.Label14.Location = New System.Drawing.Point(128, 72)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(56, 16)
        Me.Label14.TabIndex = 24
        Me.Label14.Text = "eMail . ."
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Navy
        Me.Label13.Location = New System.Drawing.Point(408, 56)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(104, 16)
        Me.Label13.TabIndex = 23
        Me.Label13.Text = "Sprzed.Ost.Trans."
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Navy
        Me.Label12.Location = New System.Drawing.Point(408, 40)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(96, 16)
        Me.Label12.TabIndex = 22
        Me.Label12.Text = "Sprzed.Nast.Kont."
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(408, 24)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(104, 16)
        Me.Label11.TabIndex = 21
        Me.Label11.Text = "Sprzed.Ost.Kont."
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Navy
        Me.Label10.Location = New System.Drawing.Point(408, 8)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(96, 16)
        Me.Label10.TabIndex = 15
        Me.Label10.Text = "Opiekun . . . . . . . . . "
        '
        'lblMiejscowosc
        '
        Me.lblMiejscowosc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblMiejscowosc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblMiejscowosc.ForeColor = System.Drawing.Color.Purple
        Me.lblMiejscowosc.Location = New System.Drawing.Point(240, 24)
        Me.lblMiejscowosc.Name = "lblMiejscowosc"
        Me.lblMiejscowosc.Size = New System.Drawing.Size(160, 16)
        Me.lblMiejscowosc.TabIndex = 12
        '
        'lblNrDomu
        '
        Me.lblNrDomu.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNrDomu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNrDomu.ForeColor = System.Drawing.Color.Purple
        Me.lblNrDomu.Location = New System.Drawing.Point(311, 40)
        Me.lblNrDomu.Name = "lblNrDomu"
        Me.lblNrDomu.Size = New System.Drawing.Size(45, 16)
        Me.lblNrDomu.TabIndex = 11
        '
        'lblKod
        '
        Me.lblKod.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblKod.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblKod.ForeColor = System.Drawing.Color.Purple
        Me.lblKod.Location = New System.Drawing.Point(168, 24)
        Me.lblKod.Name = "lblKod"
        Me.lblKod.Size = New System.Drawing.Size(64, 16)
        Me.lblKod.TabIndex = 7
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Navy
        Me.Label7.Location = New System.Drawing.Point(128, 24)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(56, 16)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Adres"
        '
        'lblIdent
        '
        Me.lblIdent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdent.ForeColor = System.Drawing.Color.Purple
        Me.lblIdent.Location = New System.Drawing.Point(64, 8)
        Me.lblIdent.Name = "lblIdent"
        Me.lblIdent.Size = New System.Drawing.Size(56, 16)
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
        Me.Label1.Text = "Nr karty . . . . . . . . "
        '
        'stbKlienci
        '
        Me.stbKlienci.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbKlienci.Dock = System.Windows.Forms.DockStyle.None
        Me.stbKlienci.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbKlienci.Location = New System.Drawing.Point(0, 292)
        Me.stbKlienci.Name = "stbKlienci"
        Me.stbKlienci.ShowPanels = True
        Me.stbKlienci.Size = New System.Drawing.Size(968, 23)
        Me.stbKlienci.TabIndex = 43
        Me.stbKlienci.Text = "StatusBar1"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 479)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 39
        Me.PictureBox1.TabStop = False
        '
        'dagKlienci
        '
        Me.dagKlienci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagKlienci.ContextMenuStrip = Me.menuEKlienci
        Me.dagKlienci.Location = New System.Drawing.Point(0, 22)
        Me.dagKlienci.Name = "dagKlienci"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagKlienci.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagKlienci.Size = New System.Drawing.Size(968, 270)
        Me.dagKlienci.TabIndex = 181
        '
        'menuEKlienci
        '
        Me.menuEKlienci.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuEEdytuj, Me.menuEDodaj, Me.menuEUsun, Me.ToolStripSeparator2, Me.menuESingleL, Me.menuEMultiL, Me.ToolStripSeparator3, Me.menuEOdswiez, Me.ToolStripSeparator1, Me.menuENowaA, Me.menuEWybierzA, Me.menuEWybierzKilkaA, Me.menuEDodajDoA, Me.menuEUsunZA, Me.ToolStripSeparator4, Me.menuEWybierzWgSprzed, Me.menuEWybierzWgOsoby, Me.menuEWybierzWgRZU, Me.menuEWybierzWgWzrostu, Me.ToolStripSeparator5, Me.menuEUstawM, Me.menuEUsunM, Me.menuEZmienM, Me.ToolStripSeparator6, Me.menuEUstawMP, Me.menuEUsunMP, Me.menuEZmienMP, Me.ToolStripSeparator7, Me.menuZmienWskAkt})
        Me.menuEKlienci.Name = "menuEdytuj"
        Me.menuEKlienci.Size = New System.Drawing.Size(465, 530)
        '
        'menuEEdytuj
        '
        Me.menuEEdytuj.Name = "menuEEdytuj"
        Me.menuEEdytuj.Size = New System.Drawing.Size(464, 22)
        Me.menuEEdytuj.Text = "Edytuj"
        '
        'menuEDodaj
        '
        Me.menuEDodaj.Name = "menuEDodaj"
        Me.menuEDodaj.Size = New System.Drawing.Size(464, 22)
        Me.menuEDodaj.Text = "Dodaj"
        '
        'menuEUsun
        '
        Me.menuEUsun.Name = "menuEUsun"
        Me.menuEUsun.Size = New System.Drawing.Size(464, 22)
        Me.menuEUsun.Text = "Usuñ"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(461, 6)
        '
        'menuESingleL
        '
        Me.menuESingleL.Name = "menuESingleL"
        Me.menuESingleL.Size = New System.Drawing.Size(464, 22)
        Me.menuESingleL.Text = "Poka¿ w jednej linii"
        '
        'menuEMultiL
        '
        Me.menuEMultiL.Name = "menuEMultiL"
        Me.menuEMultiL.Size = New System.Drawing.Size(464, 22)
        Me.menuEMultiL.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(461, 6)
        '
        'menuEOdswiez
        '
        Me.menuEOdswiez.Name = "menuEOdswiez"
        Me.menuEOdswiez.Size = New System.Drawing.Size(464, 22)
        Me.menuEOdswiez.Text = "Odœwie¿ Tabelê"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(461, 6)
        '
        'menuENowaA
        '
        Me.menuENowaA.Name = "menuENowaA"
        Me.menuENowaA.Size = New System.Drawing.Size(464, 22)
        Me.menuENowaA.Text = "Nowa akcja marketingowa"
        '
        'menuEWybierzA
        '
        Me.menuEWybierzA.Name = "menuEWybierzA"
        Me.menuEWybierzA.Size = New System.Drawing.Size(464, 22)
        Me.menuEWybierzA.Text = "Wybierz akcje marketingow¹"
        '
        'menuEWybierzKilkaA
        '
        Me.menuEWybierzKilkaA.Name = "menuEWybierzKilkaA"
        Me.menuEWybierzKilkaA.Size = New System.Drawing.Size(464, 22)
        Me.menuEWybierzKilkaA.Text = "Wybierz kilka akcji marketingowych"
        '
        'menuEDodajDoA
        '
        Me.menuEDodajDoA.Name = "menuEDodajDoA"
        Me.menuEDodajDoA.Size = New System.Drawing.Size(464, 22)
        Me.menuEDodajDoA.Text = "Dodaj do akcji marketingowej"
        '
        'menuEUsunZA
        '
        Me.menuEUsunZA.Name = "menuEUsunZA"
        Me.menuEUsunZA.Size = New System.Drawing.Size(464, 22)
        Me.menuEUsunZA.Text = "Usuñ z akcji marketingowej"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(461, 6)
        '
        'menuEWybierzWgSprzed
        '
        Me.menuEWybierzWgSprzed.Name = "menuEWybierzWgSprzed"
        Me.menuEWybierzWgSprzed.Size = New System.Drawing.Size(464, 22)
        Me.menuEWybierzWgSprzed.Text = "Wybierz klientów wg wielkoœci sprzeda¿y"
        '
        'menuEWybierzWgOsoby
        '
        Me.menuEWybierzWgOsoby.Name = "menuEWybierzWgOsoby"
        Me.menuEWybierzWgOsoby.Size = New System.Drawing.Size(464, 22)
        Me.menuEWybierzWgOsoby.Text = "Wybierz klientów wg osoby kontaktowej"
        '
        'menuEWybierzWgRZU
        '
        Me.menuEWybierzWgRZU.Name = "menuEWybierzWgRZU"
        Me.menuEWybierzWgRZU.Size = New System.Drawing.Size(464, 22)
        Me.menuEWybierzWgRZU.Text = "Wybierz klientów którzy wybrali wskazane us³ugi w ramach RZU"
        '
        'menuEWybierzWgWzrostu
        '
        Me.menuEWybierzWgWzrostu.Name = "menuEWybierzWgWzrostu"
        Me.menuEWybierzWgWzrostu.Size = New System.Drawing.Size(464, 22)
        Me.menuEWybierzWgWzrostu.Text = "Wybierz klientów którzy zwiêkszyli/zmniejszyli obroty w wybranym okresie"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(461, 6)
        '
        'menuEUstawM
        '
        Me.menuEUstawM.Name = "menuEUstawM"
        Me.menuEUstawM.Size = New System.Drawing.Size(464, 22)
        Me.menuEUstawM.Text = "Ustaw marker"
        '
        'menuEUsunM
        '
        Me.menuEUsunM.Name = "menuEUsunM"
        Me.menuEUsunM.Size = New System.Drawing.Size(464, 22)
        Me.menuEUsunM.Text = "Wyma¿ marker"
        '
        'menuEZmienM
        '
        Me.menuEZmienM.Name = "menuEZmienM"
        Me.menuEZmienM.Size = New System.Drawing.Size(464, 22)
        Me.menuEZmienM.Text = "Zmieñ ustawienie markera"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(461, 6)
        '
        'menuEUstawMP
        '
        Me.menuEUstawMP.Name = "menuEUstawMP"
        Me.menuEUstawMP.Size = New System.Drawing.Size(464, 22)
        Me.menuEUstawMP.Text = "Ustaw marker podrzêdny"
        '
        'menuEUsunMP
        '
        Me.menuEUsunMP.Name = "menuEUsunMP"
        Me.menuEUsunMP.Size = New System.Drawing.Size(464, 22)
        Me.menuEUsunMP.Text = "Wyma¿ marker podrzêdny"
        '
        'menuEZmienMP
        '
        Me.menuEZmienMP.Name = "menuEZmienMP"
        Me.menuEZmienMP.Size = New System.Drawing.Size(464, 22)
        Me.menuEZmienMP.Text = "Zmieñ ustawienia markera podrzêdnego"
        '
        'ToolStripSeparator7
        '
        Me.ToolStripSeparator7.Name = "ToolStripSeparator7"
        Me.ToolStripSeparator7.Size = New System.Drawing.Size(461, 6)
        '
        'menuZmienWskAkt
        '
        Me.menuZmienWskAkt.ForeColor = System.Drawing.Color.Red
        Me.menuZmienWskAkt.Name = "menuZmienWskAkt"
        Me.menuZmienWskAkt.Size = New System.Drawing.Size(464, 22)
        Me.menuZmienWskAkt.Text = "Oznacz klienta jako NIEAKTYWNY"
        '
        'Panel4
        '
        Me.Panel4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel4.Controls.Add(Me.cmdClearColWidth)
        Me.Panel4.Controls.Add(Me.cmdWybierzSchemat)
        Me.Panel4.Controls.Add(Me.stbKlienci)
        Me.Panel4.Controls.Add(Me.dagKlienci)
        Me.Panel4.Controls.Add(Me.cmdWszystko)
        Me.Panel4.Controls.Add(Me.cmdFi)
        Me.Panel4.Controls.Add(Me.stbFiltry)
        Me.Panel4.Location = New System.Drawing.Point(8, 154)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(968, 319)
        Me.Panel4.TabIndex = 182
        '
        'KatalogKlientow
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(984, 508)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.cmdDoExcela)
        Me.Controls.Add(Me.cmdCoDalej)
        Me.Controls.Add(Me.cmdHarmonogramWizyt)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.cmdOdswiez)
        Me.Controls.Add(Me.cmdAdrExport)
        Me.Controls.Add(Me.cbxWydruki)
        Me.Controls.Add(Me.cmdDrukuj)
        Me.Controls.Add(Me.cmdWysylaj)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PictureBox1)
        Me.ForeColor = System.Drawing.Color.Teal
        Me.Location = New System.Drawing.Point(8, 8)
        Me.Name = "KatalogKlientow"
        Me.Text = "Katalog Klientów firmy PRYZMAT"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagKlienci, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuEKlienci.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
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





    Dim dbSelectKlienci As String = "SELECT * FROM Klienci"

    Dim dbConnection As OleDb.OleDbConnection
    Dim dbCommandSelect As OleDb.OleDbCommand
    Dim dbCommandDelete As OleDb.OleDbCommand
    Dim dbCommandUpdate As OleDb.OleDbCommand
    Dim dbCommandInsert As OleDb.OleDbCommand
    Dim dbDataAdapter As OleDb.OleDbDataAdapter

    Dim sqlConnectionKlienci As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienci As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienci As SqlClient.SqlCommand
    Dim sqlCommandUpdateKlienci As SqlClient.SqlCommand
    Dim sqlCommandInsertKlienci As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienci As SqlClient.SqlDataAdapter

    Dim DataSetKlienci As New DataSet
    Dim DataViewKlienci As New DataView
    Dim ConnectionState As System.Data.ConnectionState

    Dim _CoMamRobic As String = ""



    '====================================================
    'UWAGA - aby to zadzialalo trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    '====================================================
    Private Sub KatalogKlientow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me._closeClick Then
            'e.Cancel = True     'nie pozwalaj zamkn¹c forme
            'MessageBox.Show("Proszê zamkn¹æ formê klawiszami ZAPISZ lub WYCOFAJ SIÊ", _
            '    "Uwaga :", _
            '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    MessageBoxIcon.Information, _
            '    MessageBoxDefaultButton.Button1)

            ZapamietajSzerokosciKolumn()
        Else
            'MsgBox("ZAPISZ lub WYCOFAJ SIÊ")
        End If
    End Sub
    '====================================================




    Private Sub KatalogKlientow_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        oCzasStart = DateTime.UtcNow

        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        _CoMamRobic = oCoMamRobic

        IdentUzytkownika(Program_IdUzytkownika)
        If OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaPracownik Then
            cmdUsuñ.Enabled = False
            menuEUsun.Enabled = False
        End If

        If Program_IdUzytkownika = Program_AdministratorLogin Then
            cmdDoExcela.Visible = True
        ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaAdministrator Then
            cmdDoExcela.Visible = True
        ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaSzef Then
            cmdDoExcela.Visible = True
        ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaPracownikUprzywilejowany Then
            cmdDoExcela.Visible = False
        Else        'pracownik
            cmdDoExcela.Visible = False
        End If
        '---------------------------------------
        cbxWydruki.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or
                                        System.Windows.Forms.AnchorStyles.Right Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        cbxWydruki.Items.Add("Zestawienie klientów (wydruk dla promotora)")
        cbxWydruki.Items.Add("Zestawienie klientów (wydruk dla promotora - wersja skrócona)")
        cbxWydruki.Items.Add("Zestawienie klientów (generator wydruków)")
        cbxWydruki.Items.Add("Wydruk etykiet adresowych")
        cbxWydruki.Items.Add("Wydruk karty klienta")
        cbxWydruki.SelectedIndex = 0
        '------------------------------------------------




        ''---------------------------------------------------------------------
        ''Katalog Klientow
        'Public oIdentKli As String         'IDENTKLIENTA   Text, 6
        '---Public oNrTOFIKli As Long          'NRTOFI         Long
        'Public oNrTOFITxtKli As String     'NRTOFITXT      Text 500

        'Public oKartaPaybackKli As Boolean 'KARTAPAYBACK   Logical
        'Public oNryKartPaybackKli As String 'NRYKARTPAYBACK   memo

        'Public oNazwa1Kli As String        'NAZWA1         Text, 100
        'Public oKodPoczKli As String       'KODPOCZTOWY    Text, 10
        'Public oMiejscKli As String        'MIEJSCOWOSC    Text, 50
        'Public oUlicaKli As String         'ULICA          Text, 50
        'Public oNumNrDomuKli As Integer    'NUMNRDOMU      INTEGER
        'Public oNrDomuKli As String        'NRDOMU         Text, 10
        'Public oIDDomuKli As String        'IDDOMU         Text, 10
        'Public oNIPKli As String           'NIP            Text, 15
        'Public oTelefonKli As String       'TELEFON        Text, 50
        'Public oFaxKli As String           'FAX            Text, 50
        'Public oeMailKli As String         'EMAIL          Text, 50
        'Public oWpisalKli As String        'KTOWPISAL      Text, 10   uzytkownik
        'Public oRejonKli As String         'REJKLIENTA     Text, 20   
        'Public oPKontaktKli As String      'PKONTAKT       Text, 10   data rrrr-mm-dd

        'Public oBranzaKli As String        'BRANZA         Text, 100
        'Public oPodBranzaKli As String     'PODBRANZA      Text, 100
        'Public oLiczbaZaktrudnionychKli As Integer   'LICZBAZATRUDNIONYCH       INTEGER
        'Public oLiczbaPracZdalnieKli As Integer      'LICZBAPRACZDALNIE         INTEGER
        ''..............................
        'Public oWykazUrzadzenKli As String              'WYKLAZURZADZEN    Memo

        'Public oSposobWyboruDostawcyKli As String       'SPOSOBWYBORUDOSTAWCY    Text 30
        'Public oZainstalowanoMonitorKli As Boolean      'ZAINSTALOWANOMONITOR    Logical
        'Public oLiczbaUrzadzenKli As Integer            'LICZBAURZADZEN     Integer
        'Public oLiczbaWydrukowKli As Integer            'LICZBAWYDRUKOW     Integer
        'Public oPotencjalDrukuKli As String             'POTENCJALDRUKU     Text 2

        'Public oAktZakupMatEkspKli As Boolean           'AKTZAKUPMATEKSP    Logical
        'Public oAktRozliczaStronyDrukuKli As Boolean    'AKTROZLICZASTRONYDRUKU    Logical
        'Public oAktKorzystaZNajmuKli As Boolean         'AKTKORZYSTAZNAJMU  Logical
        'Public oZaintRozliczaniemStronDrukuKli As Boolean   'ZAINTROZLICZANIEMSTRONDRUKU    Logical
        'Public oZaintKorzystaniemZNajmuKli As Boolean   'ZAINTKORZYSTANIEMZNAJMU    Logical

        'Public oDataWeryfSegmentacjiKli As String          'DATAWERYFSEGMENTACJI  Text 10

        'Public oPlanDlugookresowyKli As Integer         'PLANDLUGOOKRESOWY     Integer
        'Public oPlanKrotkookresowyKli As Integer        'PLANKROTKOOKRESOWY    Integer
        ''......................................
        'Public oRodzSprzetuLKli As Boolean 'RODZSPRZETUL Logical
        'Public oRodzSprzetuAKli As Boolean 'RODZSPRZETUA Logical
        'Public oRodzSprzetuTKli As Boolean 'RODZSPRZETUT Logical
        'Public oOfertaTZlozeniaKli As String        'OFERTATZLOZENIA         Text, 4
        'Public oOfertaPlikKli As String    'OFERTAPLIK     Text, 100
        'Public oUwagikli As String         'UWAGI          Memo

        'Public oSkupUwagikli As String        'SKUPUWAGI      Memo
        'Public oSprzedOpiekunkli As String    'SPRZEDOPIEKUN    Text, 10   uzytkownik

        'Public oSprzedOKontaktRKli As String  'SPRZEDOKONTAKTR  Text, 4    rok
        'Public oSprzedOKontaktTKli As String  'SPRZEDOKONTAKTT  Integer    nr tygodnia
        'Public oSprzedOKontaktDKli As String  'SPRZEDOKONTAKTD  Text, 12   dzien tygodnia
        'Public oSprzedNKontaktRKli As String  'SPRZEDNKONTAKTR  Text, 4    rok
        'Public oSprzedNKontaktTKli As String  'SPRZEDNKONTAKTT  Integer    nr tygodnia
        'Public oSprzedNKontaktDKli As String  'SKUPNKONTAKTT    Text, 12    nr tygodnia

        'Public oSprzedUwagiKli As String      'SPRZEDUWAGI      Memo
        'Public oWersjaKli As Integer          'WERSJA           Integer
        'Public oMarkerKli As Boolean          'MARKER           boolean
        'Public oMarkerPKli As Boolean         'MARKERP          boolean
        'Public oAktywnyKli As Boolean         'AKTYWNY          boolean

        'Public oZakupPopRokKli As Double       'ZAKUPPOPROK        Double
        'Public oUslugiDodKli As String         'USLUGIDOD          Text, 200
        'Public oRZURap01 As String              'RZURAP01     Text, 100
        'Public oRZURap02 As String              'RZURAP02     Text, 100
        'Public oRZURap03 As String              'RZURAP03     Text, 100
        'Public oRZURap04 As String              'RZURAP04     Text, 100
        'Public oRZURap05 As String              'RZURAP05     Text, 100
        'Public oRZURap06 As String              'RZURAP06     Text, 100
        'Public oRZURap07 As String              'RZURAP07     Text, 100
        'Public oRZURap08 As String              'RZURAP08     Text, 100
        'Public oRZURap09 As String              'RZURAP09     Text, 100






        dbSelectKlienci = "SELECT " &
                                             "IDENTKLIENTA AS NRKARTY, " &
                                             "NIP," &
                                             "NRTOFITXT," &
                                             "KARTAPAYBACK," &
                                             "NRYKARTPAYBACK," &
                                             "NAZWA1 AS NAZWAKLIENTA," &
                                             "MIEJSCOWOSC," &
                                             "KODPOCZTOWY," &
                                             "ULICA," &
                                             "NUMNRDOMU, "

        If ParL_DataBaseType = "MS ACCESS" Then
            dbSelectKlienci &= "IIF([NUMNRDOMU] MOD 2=0,'True','False') AS PARZYSTE, "
        Else
            dbSelectKlienci &= "CASE WHEN [NUMNRDOMU]%2=0 THEN CAST('True' AS BIT) ELSE CAST('False' AS BIT) END AS PARZYSTE, "
        End If

        dbSelectKlienci &=
                                             "NRDOMU," &
                                             "IDDOMU," &
                                             "REJKLIENTA AS REJONKLIENTA," &
                                             "TELEFON," &
                                             "FAX," &
                                             "EMAIL," &
                                               "BRANZA," &
                                               "PODBRANZA," &
                                               "LICZBAZATRUDNIONYCH," &
                                               "LICZBAPRACZDALNIE," &
                                       "WYKAZURZADZEN," &
                                         "SPOSOBWYBORUDOSTAWCY," &
                                         "ZAINSTALOWANOMONITOR," &
                                         "LICZBAURZADZEN," &
                                         "LICZBAWYDRUKOW," &
                                         "POTENCJALDRUKU," &
                                       "AKTZAKUPMATEKSP," &
                                       "AKTROZLICZASTRONYDRUKU," &
                                       "AKTKORZYSTAZNAJMU," &
                                       "ZAINTROZLICZANIEMSTRONDRUKU," &
                                       "ZAINTKORZYSTANIEMZNAJMU," &
                                         "DATAWERYFSEGMENTACJI," &
                                         "PLANDLUGOOKRESOWY," &
                                         "PLANKROTKOOKRESOWY," &
                                               "RODZSPRZETUL," &
                                               "RODZSPRZETUA," &
                                               "RODZSPRZETUT," &
                                               "OFERTATZLOZENIA," &
                                               "OFERTAPLIK," &
                                               "PKONTAKT," &
                                               "SKUPUWAGI AS PROMOTORUWAGI," &
                                                 "SPRZEDOPIEKUN AS OPIEKUN," &
                                                 "SPRZEDOKONTAKTR AS OPIEKUNOSTKONTAKTROK," &
                                                 "SPRZEDOKONTAKTT AS OPIEKUNOSTKONTAKTTYDZ," &
                                                 "SPRZEDOKONTAKTD AS OPIEKUNOSTKONTAKTDZIEN," &
                                                 "SPRZEDNKONTAKTR AS OPIEKUNKOLKONTAKTROK," &
                                                 "SPRZEDNKONTAKTT AS OPIEKUNKOLKONTAKTTYDZ," &
                                                 "SPRZEDNKONTAKTD AS OPIEKUNKOLKONTAKTDZIEN," &
                                                 "SPRZEDUWAGI AS OPIEKUNUWAGI," &
                                               "KTOWPISAL," &
                                               "UWAGI," &
                                               "MARKER, " &
                                               "MARKERP, " &
                                               "AKTYWNY, " &
                                           "ZAKUPPOPROK, " &
                                           "USLUGIDOD, " &
                                           "RZURAP01, " &
                                           "RZURAP02, " &
                                           "RZURAP03, " &
                                           "RZURAP04, " &
                                           "RZURAP05, " &
                                           "RZURAP06, " &
                                           "RZURAP07, " &
                                           "RZURAP08, " &
                                           "RZURAP09, "
        dbSelectKlienci &=
                                               "WERSJA " &
                                            "FROM Klienci ORDER BY NRKARTY ASC"




        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            'dbConnection = New OleDb.OleDbConnection(sConnectionKlienci)
            'dbCommandSelect = New OleDb.OleDbCommand(dbSelectKlienci, dbConnection)
            'dbCommandDelete = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnection)
            'dbCommandUpdate = New OleDb.OleDbCommand(sUpdateSQLKlienci, dbConnection)
            'dbCommandInsert = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnection)
            'dbDataAdapter = New OleDb.OleDbDataAdapter(dbCommandSelect)

            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            'MapowanieTabeli = dbDataAdapter.TableMappings.Add("Table", "TABELA_Klienci")

            'DBDeleteKlienci(dbCommandDelete, dbDataAdapter)
            'DBUpdateKlienci(dbCommandUpdate, dbDataAdapter)
            'DBInsertKlienci(dbCommandInsert, dbDataAdapter)

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapter.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnection.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'Finally
            '    ConnectionState = dbConnection.State
            'End Try


        Else
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(dbSelectKlienci, sqlConnectionKlienci)
            sqlCommandDeleteKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnectionKlienci)
            sqlCommandUpdateKlienci = New SqlClient.SqlCommand(sUpdateSQLKlienci, sqlConnectionKlienci)
            sqlCommandInsertKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

            SQLDeleteKlienci(sqlCommandDeleteKlienci, sqlDataAdapterKlienci)
            SQLUpdateKlienci(sqlCommandUpdateKlienci, sqlDataAdapterKlienci)
            SQLInsertKlienci(sqlCommandInsertKlienci, sqlDataAdapterKlienci)

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienci.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

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
                    'dbDataAdapter.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    'dbDataAdapter.Fill(DataSetKlienci)
                    'dbConnection.Close()
                Else
                    sqlDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                End If

                'zdefiniuj unikalny klucz wyszukiwania w tabeli
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}

                'definiuj DataView
                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                DataViewKlienci.AllowDelete = False
                DataViewKlienci.AllowNew = False

                '-------------------------------------------------
                'parametry DataGridView
                '-------------------------------------------------
                ' Initialize basic DataGridView properties.
                dagKlienci.BackgroundColor = System.Drawing.SystemColors.Control
                dagKlienci.GridColor = System.Drawing.SystemColors.ControlDark
                dagKlienci.ForeColor = System.Drawing.Color.Purple
                dagKlienci.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagKlienci.BorderStyle = BorderStyle.Fixed3D
                'dagklienci.Dock = DockStyle.Fill
                dagKlienci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                ' Set property values appropriate for read-only display and limited interactivity. 
                dagKlienci.AllowUserToAddRows = False
                dagKlienci.AllowUserToDeleteRows = False
                dagKlienci.AllowUserToOrderColumns = True
                dagKlienci.AllowUserToResizeColumns = True
                dagKlienci.AllowUserToResizeRows = True
                dagKlienci.ReadOnly = True
                dagKlienci.SelectionMode = DataGridViewSelectionMode.CellSelect
                dagKlienci.MultiSelect = False

                'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                'dagklienci.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                dagKlienci.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                dagKlienci.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                dagKlienci.ColumnHeadersVisible = True
                dagKlienci.ColumnHeadersHeight = 70
                dagKlienci.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                dagKlienci.RowHeadersVisible = True
                dagKlienci.RowHeadersWidth = 50
                dagKlienci.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                ' Set the selection background color for all the cells.
                dagKlienci.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                dagKlienci.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                dagKlienci.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagKlienci.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                dagKlienci.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                dagKlienci.DefaultCellStyle.NullValue = ""
                dagKlienci.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                dagKlienci.DefaultCellStyle.WrapMode = DataGridViewTriState.False         'single...


                ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                dagKlienci.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                dagKlienci.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                ' Set the background color for all rows and for alternating rows.  
                ' The value for alternating rows overrides the value for all rows. 
                dagKlienci.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                dagKlienci.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                ' Set the row and column header styles.
                dagKlienci.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagKlienci.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                dagKlienci.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                dagKlienci.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                '-----------------------------------
                'wypelnienie DataGridView
                dagKlienci.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                'dagKlienci.DataMember = DataSetKlienci.Tables(0).TableName
                'dagklienci.DataSource = DataViewklienci
                '-----------------------------------






                'dbSelectKlienci = "SELECT " &
                '                             "IDENTKLIENTA AS NRKARTY, " &
                '                             "NIP," &
                '                             "NRTOFITXT," &
                '                             "KARTAPAYBACK," &
                '                             "NRYKARTPAYBACK," &
                '                             "NAZWA1 AS NAZWAKLIENTA," &
                '                             "MIEJSCOWOSC," &
                '                             "KODPOCZTOWY," &
                '                             "ULICA," &
                '                             "NUMNRDOMU, " &
                '"CASE WHEN [NUMNRDOMU]%2=0 THEN CAST('True' AS BIT) ELSE CAST('False' AS BIT) END AS PARZYSTE, " &
                '                             "NRDOMU," &
                '                             "IDDOMU," &
                '                             "REJKLIENTA AS REJONKLIENTA," &
                '                             "TELEFON," &
                '                             "FAX," &
                '                             "EMAIL," &
                '"BRANZA," &
                '"PODBRANZA," &
                '"LICZBAZATRUDNIONYCH," &
                '"LICZBAPRACZDALNIE," &
                '                       "WYKAZURZADZEN," &
                '                         "SPOSOBWYBORUDOSTAWCY," &
                '                         "ZAINSTALOWANOMONITOR," &
                '                         "LICZBAURZADZEN," &
                '                         "LICZBAWYDRUKOW," &
                '                         "POTENCJALDRUKU," &
                '                       "AKTZAKUPMATEKSP," &
                '                       "AKTROZLICZASTRONYDRUKU," &
                '                       "AKTKORZYSTAZNAJMU," &
                '                       "ZAINTROZLICZANIEMSTRONDRUKU," &
                '                       "ZAINTKORZYSTANIEMZNAJMU," &
                '                         "DATAWERYFSEGMENTACJI," &
                '                         "PLANDLUGOOKRESOWY," &
                '                         "PLANKROTKOOKRESOWY," &
                '                               "RODZSPRZETUL," &
                '                               "RODZSPRZETUA," &
                '                               "RODZSPRZETUT," &
                '                               "OFERTATZLOZENIA," &
                '                               "OFERTAPLIK," &
                '                               "PKONTAKT," &
                '                               "SKUPUWAGI AS PROMOTORUWAGI," &
                '                                 "SPRZEDOPIEKUN AS OPIEKUN," &
                '                                 "SPRZEDOKONTAKTR AS OPIEKUNOSTKONTAKTROK," &
                '                                 "SPRZEDOKONTAKTT AS OPIEKUNOSTKONTAKTTYDZ," &
                '                                 "SPRZEDOKONTAKTD AS OPIEKUNOSTKONTAKTDZIEN," &
                '                                 "SPRZEDNKONTAKTR AS OPIEKUNKOLKONTAKTROK," &
                '                                 "SPRZEDNKONTAKTT AS OPIEKUNKOLKONTAKTTYDZ," &
                '                                 "SPRZEDNKONTAKTD AS OPIEKUNKOLKONTAKTDZIEN," &
                '                                 "SPRZEDUWAGI AS OPIEKUNUWAGI," &
                '                               "KTOWPISAL," &
                '                               "UWAGI," &
                '                               "MARKER, " &
                '                               "MARKERP, " &
                '                               "AKTYWNY, " &
                '                           "ZAKUPPOPROK, " &
                '                           "USLUGIDOD, " &
                '                           "RZURAP01, " &
                '                           "RZURAP02, " &
                '                           "RZURAP03, " &
                '                           "RZURAP04, " &
                '                           "RZURAP05, " &
                '                           "RZURAP06, " &
                '                           "RZURAP07, " &
                '                           "RZURAP08, " &
                '                           "RZURAP09, " &
                '                                "WERSJA " &
                '                            "FROM Klienci ORDER BY IDENTKLIENTA ASC"





                dagKlienci.Columns.Clear()
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(0).ColumnName, "Nr karty", 50, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(1).ColumnName, "NIP klienta", 100, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(2).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)

                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(3).ColumnName, "PayBack", 60, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(4).ColumnName, "Nry Kart Payback", 100, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(5).ColumnName, "Nazwa klienta", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(6).ColumnName, "Miejscowoœæ", 150, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(7).ColumnName, "Kod poczt.", 60, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(8).ColumnName, "Ulica", 150, DataGridViewContentAlignment.MiddleLeft, True)
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(9).ColumnName, "Nr domu", 80, DataGridViewContentAlignment.MiddleCenter, "F0", True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(10).ColumnName, "Parzysty Nieparzysty", 40, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(11).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(12).ColumnName, "Id domu", 50, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(13).ColumnName, "Rejon klienta", 80, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(14).ColumnName, "Telefon", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(15).ColumnName, "Fax", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(16).ColumnName, "eMail", 200, DataGridViewContentAlignment.MiddleLeft, True)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(17).ColumnName, "Bran¿a", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(18).ColumnName, "Podbran¿a", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(19).ColumnName, "Iloœæ zatrudnionych", 80, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(20).ColumnName, "Iloœæ pracuj¹cych zdalnie", 80, DataGridViewContentAlignment.MiddleCenter, True)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(21).ColumnName, "Wykaz urz¹dzeñ", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(22).ColumnName, "Sp.wyboru dostawcy", 100, DataGridViewContentAlignment.MiddleLeft, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(23).ColumnName, "Zainst. monitor", 50, True)
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(24).ColumnName, "Liczba urz¹dzeñ", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(25).ColumnName, "Liczba wydruków", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(26).ColumnName, "Potencja³ wydruku", 40, DataGridViewContentAlignment.MiddleCenter, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(27).ColumnName, "Akt. zakup mat.eksp.", 50, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(28).ColumnName, "Akt. rozlicza strony druku", 50, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(29).ColumnName, "Akt. korzysta z najmu", 50, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(30).ColumnName, "Zaint. rozlicz. stron druku", 50, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(31).ColumnName, "Zaint. korzyst. z najmu", 50, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(32).ColumnName, "Data weryfikacji segmentacji", 80, DataGridViewContentAlignment.MiddleLeft, True)
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(33).ColumnName, "Plan d³ugookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(34).ColumnName, "Plan krótkookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)

                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(35).ColumnName, "Rodz.Sprzêtu L", 100, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(36).ColumnName, "Rodz.Sprzêtu A", 100, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(37).ColumnName, "Rodz.Sprzêtu A3", 100, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(38).ColumnName, "Oferta z³o¿enia", 40, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(39).ColumnName, "Plik z ofert¹", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(40).ColumnName, "I-kontakt", 100, DataGridViewContentAlignment.MiddleLeft, True)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(41).ColumnName, "Opis potencja³u", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(42).ColumnName, "Opiekun", 100, DataGridViewContentAlignment.MiddleLeft, True)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(43).ColumnName, "Opiekun Ost.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(44).ColumnName, "Opiekun Ost.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(45).ColumnName, "Opiekun Ost.kontakt dz.", 150, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(46).ColumnName, "Opiekun Kol.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(47).ColumnName, "Opiekun Kol.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(48).ColumnName, "Opiekun Kol.kontakt dz.", 150, DataGridViewContentAlignment.MiddleLeft, True)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(49).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(50).ColumnName, "Kto wpisa³ ?", 100, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(51).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleCenter, True)

                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Marker  ", 50, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Marker P.  ", 50, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "Aktywny  ", 50, True)

                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(55).ColumnName, "Zakup pop.rok   ", 60, DataGridViewContentAlignment.MiddleRight, "F2", True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(56).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(57).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(58).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(59).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(60).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(61).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(62).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(63).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(64).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(65).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(66).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

                Me.Cursor = Cursors.WaitCursor
                dagKlienci.DataSource = DataViewKlienci
                Me.Cursor = Cursors.Default

                'ustaw sie na pierwszym zapisie
                If DataSetKlienci.Tables(0).Rows.Count > 0 Then
                    dagKlienci.CurrentCell = dagKlienci(0, 0)
                    dagKlienci.CurrentCell.Selected = True
                End If

                UstalSzerokosciKolumn()





            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End Try

            '------------------------------------
            'dodaj StatusBar na koncu
            stbKlienci.Panels.Clear()

            stbKlienci.BackColor = System.Drawing.SystemColors.Control
            stbKlienci.ForeColor = System.Drawing.Color.DarkGreen
            stbKlienci.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            stbKlienci.Panels.Add("stbPanelIleRec")
            stbKlienci.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(0).Width = 100
            stbKlienci.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(0).Text = "Il.rec.: " & DataSetKlienci.Tables(0).Rows.Count.ToString

            stbKlienci.Panels.Add("stbPanelWiersz")
            stbKlienci.Panels(1).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(1).Width = 80
            stbKlienci.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(1).Text = "Wi: " & IIf(dagKlienci.CurrentCell Is Nothing, "", dagKlienci.CurrentCell.RowIndex.ToString)

            stbKlienci.Panels.Add("stbPanelKolumna")
            stbKlienci.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(2).Width = 80
            stbKlienci.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(2).Text = "Ko: " & IIf(dagKlienci.CurrentCell Is Nothing, "", dagKlienci.CurrentCell.ColumnIndex.ToString)

            stbKlienci.Panels.Add("stbPanelSort")
            stbKlienci.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(3).Width = 150
            stbKlienci.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(3).Text = "Sort: NRKARTY"

            stbKlienci.Panels.Add("stbPanelSzablon")
            stbKlienci.Panels(4).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(4).Width = 150
            stbKlienci.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(4).Text = "Szablon: "

            'dodaj ctrl do pobierania szablonu
            cmdWybierzSchemat.Location = New System.Drawing.Point(stbKlienci.Location.X +
                                                stbKlienci.Panels(0).Width +
                                                stbKlienci.Panels(1).Width +
                                                stbKlienci.Panels(2).Width +
                                                stbKlienci.Panels(3).Width +
                                                stbKlienci.Panels(4).Width, stbKlienci.Location.Y + 2)
            cmdWybierzSchemat.Size = New Size(60, 20)
            cmdWybierzSchemat.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)



            stbKlienci.Panels.Add("stbPanelReszta")
            stbKlienci.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
            'stbKlienci.Panels(5).Width = 150
            stbKlienci.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(5).Text = ""

            stbKlienci.ShowPanels = True
            '---------------------------------
            'dodaj StatusBar na koncu
            stbFiltry.Panels.Clear()

            stbFiltry.BackColor = System.Drawing.SystemColors.Control
            stbFiltry.ForeColor = System.Drawing.Color.DarkGreen
            stbFiltry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

            stbFiltry.Panels.Add("stbFiltryRowHead")
            stbFiltry.Panels(0).AutoSize = StatusBarPanelAutoSize.None
            stbFiltry.Panels(0).Width = 50  'dagKlienci.TableStyles(0).RowHeaderWidth
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

        End If
        InitKlienci() 'inicjuj zmienne
        '-----------------------------------------------------------------
        'If _OCoMamRobic = "WYBIERAJ" Then
        '    dagKlienci.ContextMenuStrip = Me.menuWybieraj
        '    cmdStop.Text = "Wybierz"
        '    menuWybieraj.Visible = False
        '    menuWybieraj.Enabled = True
        '    menuEdytuj.Visible = False
        '    menuEdytuj.Enabled = False
        'Else
        '    dagKlienci.ContextMenuStrip = Me.menuEdytuj
        '    cmdStop.Text = "Powrót"
        '    menuWybieraj.Visible = False
        '    menuWybieraj.Enabled = False
        '    menuEdytuj.Visible = False
        '    menuEdytuj.Enabled = True
        'End If
        '--------------------------------------------------------------------
        'set the header width to something apporpriate
        dagKlienci.RowHeadersWidth = 50       'use if tablestyle
        'Me.dagKlienci.RowHeaderWidth = 80 'use if no tablestyle
        cmdClearColWidth.Size = New Size(dagKlienci.RowHeadersWidth, dagKlienci.ColumnHeadersHeight)
        cmdClearColWidth.BackColor = System.Drawing.SystemColors.ScrollBar
        cmdClearColWidth.Location = New System.Drawing.Point(dagKlienci.Location.X,
                                                             dagKlienci.Location.Y)
        '--------------------------------------------------------------------
        'set topleft corner point so we can locate the toprow
        If DataSetKlienci.Tables(0).Rows.Count > 0 Then
            StartPointInCell00 = New System.Drawing.Point((dagKlienci.GetCellDisplayRectangle(0, 0, True).X + 4),
                                      (dagKlienci.GetCellDisplayRectangle(0, 0, True).Y + 4))
            ScrollXOffset = 10000 * dagKlienci.FirstDisplayedScrollingColumnIndex +
                        dagKlienci.GetCellDisplayRectangle(dagKlienci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                        StartPointInCell00.X
        Else
            'brak zapisow - poczatek DataGrid
            StartPointInCell00 = New System.Drawing.Point((dagKlienci.Left + 4), (dagKlienci.Top + 4))
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
        For ii = 0 To DataViewKlienci.Table.Columns.Count - 1
            'stworz zmienn¹ TXTBOX
            Dim CtrlT As New System.Windows.Forms.TextBox
            CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
            CtrlT.Visible = False
            Me.Panel4.Controls.Add(CtrlT)
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
            Me.Panel4.Controls.Add(CtrlB)
            AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXX_Click)
            AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXX_GotFocus)
            AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXX_LostFocus)
        Next
        Me.Refresh()
        '--------------------------------------------------
        StartFormy = False
        dagKlienci_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)

        'If stbKlienci.Panels.Count > 6 Then
        oCzasStop = DateTime.UtcNow
        If stbKlienci.Panels.Count > 0 Then
            stbKlienci.Panels(5).Alignment = HorizontalAlignment.Right
            stbKlienci.Panels(5).Text = "Czas ini.(milisek) : " & (oCzasStop - oCzasStart).TotalMilliseconds.ToString
        End If
        'End If

        OdswiezBaze()
    End Sub



    Private Sub KatalogKlientow_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub











    Private Sub dagKlienci_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.CurrentCellChanged
        Dim Txt As String = ""
        Dim words() As String
        Dim splitChars() As Char = {","c}
        Dim iwords As Integer = 0

        If Not StartFormy Then
            If dagKlienci.CurrentCell Is Nothing Then
                stbKlienci.Panels(1).Text = "Wi: "
                stbKlienci.Panels(2).Text = "Ko: "

                lblIdent.Text = ""
                lblNazwaHandlowa.Text = ""
                txtTOFI.Text = ""
                lblRodzaj.Text = ""

                lblKod.Text = ""
                lblMiejscowosc.Text = ""
                lblUlica.Text = ""
                lblNrDomu.Text = ""
                lblIdDomu.Text = ""
                lblTelefon.Text = ""
                lbleMail.Text = ""

                lblOstKontakt.Text = ""
                lblNastKontakt.Text = ""
                lblOpiekun.Text = ""
                lblPotencjal.Text = ""
                lblTransakcje.Text = ""
                lblIlSprzetu.Text = ""

            Else
                stbKlienci.Panels(1).Text = "Wi: " & dagKlienci.CurrentCell.RowIndex.ToString
                stbKlienci.Panels(2).Text = "Ko: " & dagKlienci.CurrentCell.ColumnIndex.ToString

                If DataViewKlienci.Count = 0 Then
                    lblIdent.Text = ""
                    lblNazwaHandlowa.Text = ""
                    txtTOFI.Text = ""
                    lblRodzaj.Text = ""

                    lblKod.Text = ""
                    lblMiejscowosc.Text = ""
                    lblUlica.Text = ""
                    lblNrDomu.Text = ""
                    lblIdDomu.Text = ""
                    lblTelefon.Text = ""
                    lbleMail.Text = ""

                    lblOstKontakt.Text = ""
                    lblNastKontakt.Text = ""
                    lblOpiekun.Text = ""
                    lblPotencjal.Text = ""
                    lblTransakcje.Text = ""
                    lblIlSprzetu.Text = ""

                Else
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(0).ColumnName, "Nr karty", 50, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(1).ColumnName, "NIP klienta", 100, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(2).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)

                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(3).ColumnName, "PayBack", 60, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(4).ColumnName, "Nry Kart Payback", 100, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(5).ColumnName, "Nazwa klienta", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(6).ColumnName, "Miejscowoœæ", 150, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(7).ColumnName, "Kod poczt.", 60, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(8).ColumnName, "Ulica", 150, DataGridViewContentAlignment.MiddleLeft, True)
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(9).ColumnName, "Nr domu", 80, DataGridViewContentAlignment.MiddleCenter, "F0", True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(10).ColumnName, "Parzysty Nieparzysty", 40, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(11).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(12).ColumnName, "Id domu", 50, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(13).ColumnName, "Rejon klienta", 80, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(14).ColumnName, "Telefon", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(15).ColumnName, "Fax", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(16).ColumnName, "eMail", 200, DataGridViewContentAlignment.MiddleLeft, True)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(17).ColumnName, "Bran¿a", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(18).ColumnName, "Podbran¿a", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(19).ColumnName, "Iloœæ zatrudnionych", 80, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(20).ColumnName, "Iloœæ pracuj¹cych zdalnie", 80, DataGridViewContentAlignment.MiddleCenter, True)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(21).ColumnName, "Wykaz urz¹dzeñ", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(22).ColumnName, "Sp.wyboru dostawcy", 100, DataGridViewContentAlignment.MiddleLeft, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(23).ColumnName, "Zainst. monitor", 50, True)
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(24).ColumnName, "Liczba urz¹dzeñ", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(25).ColumnName, "Liczba wydruków", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(26).ColumnName, "Potencja³ wydruku", 40, DataGridViewContentAlignment.MiddleCenter, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(27).ColumnName, "Akt. zakup mat.eksp.", 50, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(28).ColumnName, "Akt. rozlicza strony druku", 50, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(29).ColumnName, "Akt. korzysta z najmu", 50, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(30).ColumnName, "Zaint. rozlicz. stron druku", 50, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(31).ColumnName, "Zaint. korzyst. z najmu", 50, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(32).ColumnName, "Data weryfikacji segmentacji", 80, DataGridViewContentAlignment.MiddleLeft, True)
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(33).ColumnName, "Plan d³ugookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(34).ColumnName, "Plan krótkookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)

                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(35).ColumnName, "Rodz.Sprzêtu L", 100, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(36).ColumnName, "Rodz.Sprzêtu A", 100, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(37).ColumnName, "Rodz.Sprzêtu A3", 100, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(38).ColumnName, "Oferta z³o¿enia", 40, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(39).ColumnName, "Plik z ofert¹", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(40).ColumnName, "I-kontakt", 100, DataGridViewContentAlignment.MiddleLeft, True)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(41).ColumnName, "Opis potencja³u", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(42).ColumnName, "Opiekun", 100, DataGridViewContentAlignment.MiddleLeft, True)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(43).ColumnName, "Opiekun Ost.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(44).ColumnName, "Opiekun Ost.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(45).ColumnName, "Opiekun Ost.kontakt dz.", 150, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(46).ColumnName, "Opiekun Kol.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(47).ColumnName, "Opiekun Kol.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(48).ColumnName, "Opiekun Kol.kontakt dz.", 150, DataGridViewContentAlignment.MiddleLeft, True)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(49).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(50).ColumnName, "Kto wpisa³ ?", 100, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(51).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleCenter, True)

                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Marker  ", 50, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Marker P.  ", 50, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "Aktywny  ", 50, True)

                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(55).ColumnName, "Zakup pop.rok   ", 60, DataGridViewContentAlignment.MiddleRight, "F2", True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(56).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(57).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(58).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(59).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(60).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(61).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(62).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(63).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(64).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(65).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(66).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)




                    lblIdent.Text = GetTxtField(dagKlienci, 0)

                    'podziel linie na poszczególne czêœci
                    Txt = GetTxtField(dagKlienci, 2)
                    If Len(Txt) = 0 Then
                        txtTOFI.Text = ""
                    Else
                        words = Txt.Split(splitChars, StringSplitOptions.None)
                        If words.Length = 0 Then
                            txtTOFI.Text = ""
                        ElseIf words.Length = 1 Then
                            txtTOFI.Text = Txt
                        Else
                            For iwords = 0 To words.Length - 1
                                If Len(words(iwords)) > 0 Then
                                    txtTOFI.Text &= words(iwords) & vbCrLf
                                End If
                            Next
                        End If
                    End If

                    lblNazwaHandlowa.Text = GetTxtField(dagKlienci, 5)

                    oRodzSprzetuLKli = GetLogField(dagKlienci, 35)
                    oRodzSprzetuAKli = GetLogField(dagKlienci, 36)
                    oRodzSprzetuTKli = GetLogField(dagKlienci, 37)
                    lblRodzaj.Text = IIf(oRodzSprzetuLKli, "L", "") & IIf(oRodzSprzetuAKli, "A", "") & IIf(oRodzSprzetuTKli, "T", "")

                    lblKod.Text = GetTxtField(dagKlienci, 7)
                    lblMiejscowosc.Text = GetTxtField(dagKlienci, 6)
                    lblUlica.Text = GetTxtField(dagKlienci, 8)
                    lblNrDomu.Text = GetIntField(dagKlienci, 9).ToString("F0")
                    lblIdDomu.Text = GetTxtField(dagKlienci, 12)
                    lblTelefon.Text = GetTxtField(dagKlienci, 14)
                    lbleMail.Text = GetTxtField(dagKlienci, 16)

                    lblIlSprzetu.Text = GetTxtField(dagKlienci, 17)

                    lblOpiekun.Text = GetTxtField(dagKlienci, 42)
                    lblOstKontakt.Text = "Tydzieñ " & Trim(GetTxtField(dagKlienci, 44)) & " / " & GetTxtField(dagKlienci, 43)
                    lblNastKontakt.Text = "Tydzieñ " & Trim(GetTxtField(dagKlienci, 47)) & " / " & GetTxtField(dagKlienci, 46)
                    lblTransakcje.Text = OstTransakcje(lblIdent.Text, True)
                    lblPotencjal.Text = GetTxtField(dagKlienci, 41)

                End If
            End If
        End If
    End Sub


    '-=------------------------------------------
    'obsluga zegara scrollingu poziomego
    '-=------------------------------------------
    Private Sub HorizScrollTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HorizScrollTimer.Tick
        HorizScrollTimer.Enabled = False
        If Len(HorizScrollLastTime) = 0 Then
            'nie zainicjowano scrollingu
        Else
            'zainicjowano scrolling - sprawdz czy juz trwa 2 sec
            If IleSekundMinelo(HorizScrollLastTime, SysGodz()) >= 0 Then
                If dagKlienci.RowCount > 0 Then
                    If ScrollXOffset <> 10000 * dagKlienci.FirstDisplayedScrollingColumnIndex +
                                    dagKlienci.GetCellDisplayRectangle(dagKlienci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X Then

                        PokazFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                        '-----------------------------
                        ScrollXOffset = 10000 * dagKlienci.FirstDisplayedScrollingColumnIndex +
                                    dagKlienci.GetCellDisplayRectangle(dagKlienci.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                                    StartPointInCell00.X
                        HorizScrollLastTime = ""
                    End If
                Else
                    PokazFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Else
            End If
        End If
        HorizScrollTimer.Enabled = True
    End Sub


    Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        Dim rc As DataGridViewSelectedRowCollection = dagKlienci.SelectedRows
        If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
            If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
                Dim nextrow As Integer = rc(0).Index + 1
                dagKlienci.Rows(nextrow).Selected = True
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
                    PrzejdzDoDGV(dagKlienci, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXX_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormy Then
            FiltrujDataViewDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, DataViewKlienci, stbKlienci, oSchematFiltrowania)
        End If
    End Sub
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me.Panel4, dagKlienci, Mid(sender.name, 6, 3), "Klienci")
    End Sub
    Private Sub cmdFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiFiltraCmd(Me.Panel4, sender)
    End Sub
    Private Sub cmdFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraCmd(Me.Panel4, sender)
    End Sub


    '-------------------------------------------------
    ' filtrowanie 
    '-------------------------------------------------

    Private Sub cmdWszystko_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWszystko.Click
        Me.Cursor = Cursors.WaitCursor
        oLokalnyFiltr = ""
        oNazwaSchematu = ""
        oSchematFiltrowania = ""
        stbKlienci.Panels(4).Text = "Szablon : "
        Try
            StartFormy = True
            CzyscFiltryDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, DataViewKlienci, stbKlienci)
            StartFormy = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbKlienci.Panels(0).Text = "Il.zap.: " & DataViewKlienci.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagKlienci_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagKlienci.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagKlienci.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagKlienci.DefaultCellStyle.Font,
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
            PokazFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    Private Sub dagKlienci_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagKlienci.CellFormatting
        Dim NrTofi As String = ""
        Dim Akt As Boolean = True
        If Not StartFormy Then
            '-----------------------
            'zmieniamy ForeColor
            NrTofi = IIf(IsDBNull(DataViewKlienci.Item(e.RowIndex).Item("NRTOFITXT")), "", DataViewKlienci.Item(e.RowIndex).Item("NRTOFITXT"))
            Akt = IIf(IsDBNull(DataViewKlienci.Item(e.RowIndex).Item("AKTYWNY")), "", DataViewKlienci.Item(e.RowIndex).Item("AKTYWNY"))

            If Len(Trim(NrTofi)) = 0 Then
                'brak numeru TOFI
                e.CellStyle.ForeColor = Color.Red
            Else
                Dim words() As String
                Dim splitChars() As Char = {","c}
                words = NrTofi.Split(splitChars, StringSplitOptions.None)
                If words.Length = 1 Then
                    'jeden numer TOFI
                    e.CellStyle.ForeColor = Color.Purple
                Else
                    'wiêcej ni¿ jeden numer TOFI
                    e.CellStyle.ForeColor = Color.Green
                End If
            End If
            '-----------------------
            ''zmieniamy BackColor
            'Dim Wal As String = GetTxtField(dagKursyWalut, e.RowIndex, 0)
            If Not Akt Then
                'e.CellStyle.BackColor = System.Drawing.Color.FromArgb(CType(180, Byte), CType(255, Byte), CType(180, Byte))
                'e.CellStyle.BackColor = System.Drawing.Color.FromArgb(CType(200, Byte), CType(255, Byte), CType(200, Byte))
                'e.CellStyle.BackColor = System.Drawing.Color.FromArgb(CType(220, Byte), CType(255, Byte), CType(220, Byte))
                e.CellStyle.BackColor = System.Drawing.Color.FromArgb(CType(251, Byte), CType(196, Byte), CType(200, Byte))
            End If
            '-----------------------
        End If
    End Sub




    '*********************************************************
    '***** obsluga dag
    '*********************************************************

    Private Sub dagKlienci_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagKlienci.ColumnWidthChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKlienci_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.Resize
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKlienci_Scroll(sender As Object, e As ScrollEventArgs) Handles dagKlienci.Scroll
        'If (Not StartFormy) And (DataViewKlienci.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        '        End If
        '    End If
        'End If

        If (Not StartFormy) And (DataViewKlienci.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagKlienci_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.SizeChanged
        If Not StartFormy Then
            PokazFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        End If
    End Sub

    Private Sub dagKlienci_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagKlienci.MouseMove
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
                '    PokazFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormy Then
                    PokazFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagKlienci_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagKlienci.MouseUp
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
                stbKlienci.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbKlienci.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagKlienci(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbKlienci.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbKlienci.Panels(3).Text = "Sort: " & DataSetKlienci.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbKlienci.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagKlienci(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbKlienci.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbKlienci.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub




    Private Sub dagKlienci_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKlienci.DoubleClick
        If _CoMamRobic = "WYBIERAJ" Then
            cmdStop_Click(sender, e)
        Else
            cmdEdytuj_Click(sender, e)
        End If
    End Sub

    Private Sub dagKlienci_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagKlienci.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If _CoMamRobic = "WYBIERAJ" Then
                    cmdStop_Click(sender, e)
                Else
                    cmdEdytuj_Click(sender, e)
                End If
            Case Keys.Insert
                cmdDodaj_Click(sender, e)
            Case Keys.Delete
                cmdUsuñ_Click(sender, e)

            Case Keys.D1 AndAlso e.Control
                ZmianaUstawieniaMarkera("")
            Case Keys.D2 AndAlso e.Control
                ZmianaUstawieniaMarkeraP("")

            Case Else
        End Select
    End Sub

































    '*************************************************
    '*** przenoszenie danych miêdzy rekordem i zmiennymi
    '*************************************************


    Private Sub InitKlienci()
        oIdentKli = ""
        oNIPKli = ""
        oNrTOFIKli = 0
        oNrTOFITxtKli = ""
        oKartaPaybackKli = False
        oNryKartPaybackKli = ""

        oNazwa1Kli = ""
        oMiejscKli = ""
        oKodPoczKli = ""
        oUlicaKli = ""
        oNumNrDomuKli = 0
        oNrDomuKli = ""
        oIDDomuKli = ""
        oRejonKli = ""
        oTelefonKli = ""
        oFaxKli = ""
        oeMailKli = ""

        oBranzaKli = ""
        oPodBranzaKli = ""
        oLiczbaZaktrudnionychKli = 0
        oLiczbaPracZdalnieKli = 0

        oWykazUrzadzenKli = ""
        oSposobWyboruDostawcyKli = ""
        oZainstalowanoMonitorKli = False
        oLiczbaUrzadzenKli = 0
        oLiczbaWydrukowKli = 0
        oPotencjalDrukuKli = ""
        oAktZakupMatEkspKli = False
        oAktRozliczaStronyDrukuKli = False
        oAktKorzystaZNajmuKli = False
        oZaintRozliczaniemStronDrukuKli = False
        oZaintKorzystaniemZNajmuKli = False
        oDATAWERYFSEGMENTACJIkli = ""
        oPlanDlugookresowyKli = 0
        oPlanKrotkookresowyKli = 0

        oRodzSprzetuLKli = False
        oRodzSprzetuAKli = False
        oRodzSprzetuTKli = False
        oOfertaTZlozeniaKli = ""
        oOfertaPlikKli = ""
        oPKontaktKli = SysData()
        oSkupUwagikli = ""
        oSprzedOpiekunkli = ""
        oSprzedOKontaktRKli = Mid(SysData(), 1, 4)
        oSprzedOKontaktTKli = ""
        oSprzedOKontaktDKli = ""
        oSprzedNKontaktRKli = Mid(SysData(), 1, 4)
        oSprzedNKontaktTKli = ""
        oSprzedNKontaktDKli = ""
        oSprzedUwagiKli = ""
        oWpisalKli = ""
        oUwagikli = ""
        oMarkerKli = False
        oMarkerPKli = False
        oAktywnyKli = True
        oZakupPopRokKli = 0
        oUslugiDodKli = ""
        oRZURap01 = ""
        oRZURap02 = ""
        oRZURap03 = ""
        oRZURap04 = ""
        oRZURap05 = ""
        oRZURap06 = ""
        oRZURap07 = ""
        oRZURap08 = ""
        oRZURap09 = ""

        oWersjaKli = 1          'WERSJA         Integer, 2
    End Sub

    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(0).ColumnName, "Nr karty", 50, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(1).ColumnName, "NIP klienta", 100, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(2).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)

    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(3).ColumnName, "PayBack", 60, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(4).ColumnName, "Nry Kart Payback", 100, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(5).ColumnName, "Nazwa klienta", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(6).ColumnName, "Miejscowoœæ", 150, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(7).ColumnName, "Kod poczt.", 60, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(8).ColumnName, "Ulica", 150, DataGridViewContentAlignment.MiddleLeft, True)
    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(9).ColumnName, "Nr domu", 80, DataGridViewContentAlignment.MiddleCenter, "F0", True)
    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(10).ColumnName, "Parzysty Nieparzysty", 40, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(11).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(12).ColumnName, "Id domu", 50, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(13).ColumnName, "Rejon klienta", 80, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(14).ColumnName, "Telefon", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(15).ColumnName, "Fax", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(16).ColumnName, "eMail", 200, DataGridViewContentAlignment.MiddleLeft, True)

    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(17).ColumnName, "Wykaz urz¹dzeñ", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(18).ColumnName, "Sp.wyboru dostawcy", 100, DataGridViewContentAlignment.MiddleLeft, True)
    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(19).ColumnName, "Zainst. monitor", 50, True)
    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(20).ColumnName, "Liczba urz¹dzeñ", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(21).ColumnName, "Liczba wydruków", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(22).ColumnName, "Potencja³ wydruku", 40, DataGridViewContentAlignment.MiddleCenter, True)
    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(23).ColumnName, "Akt. zakup mat.eksp.", 50, True)
    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(24).ColumnName, "Akt. rozlicza strony druku", 50, True)
    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(25).ColumnName, "Akt. korzysta z najmu", 50, True)
    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(26).ColumnName, "Zaint. rozlicz. stron druku", 50, True)
    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(27).ColumnName, "Zaint. korzyst. z najmu", 50, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(28).ColumnName, "Data weryfikacji segmentacji", 80, DataGridViewContentAlignment.MiddleCenter, True)
    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(29).ColumnName, "Plan d³ugookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(30).ColumnName, "Plan krótkookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)

    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(31).ColumnName, "Rodz.Sprzêtu L", 100, True)
    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(32).ColumnName, "Rodz.Sprzêtu A", 100, True)
    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(33).ColumnName, "Rodz.Sprzêtu A3", 100, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(34).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(35).ColumnName, "Sprzêt - iloœæ", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(36).ColumnName, "Oferta z³o¿enia", 40, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(37).ColumnName, "Plik z ofert¹", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(38).ColumnName, "I-kontakt", 100, DataGridViewContentAlignment.MiddleLeft, True)

    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(39).ColumnName, "Opis potencja³u", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(40).ColumnName, "Opiekun", 100, DataGridViewContentAlignment.MiddleCenter, True)

    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(41).ColumnName, "Opiekun Ost.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(42).ColumnName, "Opiekun Ost.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(43).ColumnName, "Opiekun Ost.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(44).ColumnName, "Opiekun Kol.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(45).ColumnName, "Opiekun Kol.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(46).ColumnName, "Opiekun Kol.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(47).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(48).ColumnName, "Kto wpisa³ ?", 100, DataGridViewContentAlignment.MiddleCenter, True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(49).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleCenter, True)

    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(50).ColumnName, "Marker  ", 50, True)
    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(51).ColumnName, "Marker P.  ", 50, True)
    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Aktywny  ", 50, True)

    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Zakup pop.rok   ", 60, DataGridViewContentAlignment.MiddleRight, "F2", True)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(55).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(56).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(57).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(58).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(59).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(60).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(61).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(62).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(63).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(64).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)


    Private Sub PobierzKlienci()
        'pobierz wartosci przed aktualizacja
        oIdentKli = GetTxtField(dagKlienci, 0)
        oNIPKli = GetTxtField(dagKlienci, 1)
        oNrTOFITxtKli = GetTxtField(dagKlienci, 2)
        oKartaPaybackKli = GetLogField(dagKlienci, 3)
        oNryKartPaybackKli = GetTxtField(dagKlienci, 4)

        oNazwa1Kli = GetTxtField(dagKlienci, 5)
        oMiejscKli = GetTxtField(dagKlienci, 6)
        oKodPoczKli = GetTxtField(dagKlienci, 7)
        oUlicaKli = GetTxtField(dagKlienci, 8)
        oNumNrDomuKli = GetIntField(dagKlienci, 9)
        '10 parzyste/nieparzyste
        oNrDomuKli = GetTxtField(dagKlienci, 11)
        oIDDomuKli = GetTxtField(dagKlienci, 12)
        oRejonKli = GetTxtField(dagKlienci, 13)
        oTelefonKli = GetTxtField(dagKlienci, 14)
        oFaxKli = GetTxtField(dagKlienci, 15)
        oeMailKli = GetTxtField(dagKlienci, 16)

        oBranzaKli = GetTxtField(dagKlienci, 17)
        oPodBranzaKli = GetTxtField(dagKlienci, 18)
        oLiczbaZaktrudnionychKli = GetIntField(dagKlienci, 19)
        oLiczbaPracZdalnieKli = GetIntField(dagKlienci, 20)

        oWykazUrzadzenKli = GetTxtField(dagKlienci, 21)
        oSposobWyboruDostawcyKli = GetTxtField(dagKlienci, 22)
        oZainstalowanoMonitorKli = GetLogField(dagKlienci, 23)
        oLiczbaUrzadzenKli = GetNumField(dagKlienci, 24)
        oLiczbaWydrukowKli = GetNumField(dagKlienci, 25)
        oPotencjalDrukuKli = GetTxtField(dagKlienci, 26)
        oAktZakupMatEkspKli = GetLogField(dagKlienci, 27)
        oAktRozliczaStronyDrukuKli = GetLogField(dagKlienci, 28)
        oAktKorzystaZNajmuKli = GetLogField(dagKlienci, 29)
        oZaintRozliczaniemStronDrukuKli = GetLogField(dagKlienci, 30)
        oZaintKorzystaniemZNajmuKli = GetLogField(dagKlienci, 31)
        oDataWeryfSegmentacjiKli = GetTxtField(dagKlienci, 32)
        oPlanDlugookresowyKli = GetNumField(dagKlienci, 33)
        oPlanKrotkookresowyKli = GetNumField(dagKlienci, 34)

        oRodzSprzetuLKli = GetLogField(dagKlienci, 35)
        oRodzSprzetuAKli = GetLogField(dagKlienci, 36)
        oRodzSprzetuTKli = GetLogField(dagKlienci, 37)

        oOfertaTZlozeniaKli = GetTxtField(dagKlienci, 38)
        oOfertaPlikKli = GetTxtField(dagKlienci, 39)
        oPKontaktKli = GetTxtField(dagKlienci, 40)

        oSkupUwagikli = GetTxtField(dagKlienci, 41)
        oSprzedOpiekunkli = GetTxtField(dagKlienci, 42)

        oSprzedOKontaktRKli = GetTxtField(dagKlienci, 43)
        oSprzedOKontaktTKli = GetTxtField(dagKlienci, 44)
        oSprzedOKontaktDKli = GetTxtField(dagKlienci, 45)
        oSprzedNKontaktRKli = GetTxtField(dagKlienci, 46)
        oSprzedNKontaktTKli = GetTxtField(dagKlienci, 47)
        oSprzedNKontaktDKli = GetTxtField(dagKlienci, 48)

        oSprzedUwagiKli = GetTxtField(dagKlienci, 49)
        oWpisalKli = GetTxtField(dagKlienci, 50)
        oUwagikli = GetTxtField(dagKlienci, 51)
        oMarkerKli = GetLogField(dagKlienci, 52)
        oMarkerPKli = GetLogField(dagKlienci, 53)
        oAktywnyKli = GetLogField(dagKlienci, 54)

        oZakupPopRokKli = GetDblField(dagKlienci, 55)
        oUslugiDodKli = GetTxtField(dagKlienci, 56)
        oRZURap01 = GetTxtField(dagKlienci, 57)
        oRZURap02 = GetTxtField(dagKlienci, 58)
        oRZURap03 = GetTxtField(dagKlienci, 59)
        oRZURap04 = GetTxtField(dagKlienci, 60)
        oRZURap05 = GetTxtField(dagKlienci, 61)
        oRZURap06 = GetTxtField(dagKlienci, 62)
        oRZURap07 = GetTxtField(dagKlienci, 63)
        oRZURap08 = GetTxtField(dagKlienci, 64)
        oRZURap09 = GetTxtField(dagKlienci, 65)

        oWersjaKli = GetIntField(dagKlienci, 66)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdWybierzSchemat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierzSchemat.Click
        oCoMamRobic = "WYBIERAJ"
        oLokalnyFiltr = ""
        oNazwaSchematu = ""
        oSchematFiltrowania = ""

        Dim Form As New SchematFiltrowania("KatalogKlientow")
        Form.ShowDialog()

        'If Len(Trim(oSchematFiltrowania)) > 0 Then
        stbKlienci.Panels(4).Text = "Szablon : " & oNazwaSchematu
        Try
            'DataViewKlienci.RowFilter = ""
            ''DataViewKlienci.RowFilter = Trim(cbxFilter.SelectedItem.ToString) & " LIKE '" & Trim(txtFilter.Text) & "*'"
            'DataViewKlienci.RowFilter = BudujFiltr(Trim(oSchematFiltrowania), oLokalnyFiltr)
            'dagKlienci.DataSource = DataViewKlienci
            'stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString

            FiltrujDataViewDGV(Me.Panel4,
                            dagKlienci,
                            DataViewKlienci.Table.Columns.Count,
                            DataViewKlienci,
                            stbKlienci,
                            Trim(oSchematFiltrowania))

            'FiltrujDataView(Me.Panel4,
            '                dagKlienci,
            '                DataViewKlienci.Table.Columns.Count,
            '                DataViewKlienci,
            '                stbKlienci,
            '                Trim(oSchematFiltrowania))
            ''--------------------------------------------------------------------

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        End Try
        'Else
        '   stbKlienci.Panels(4).Text = "Szablon : "
        'End If

        If DataViewKlienci.Count > 0 Then
            'ustaw sie na poczatku DataGrid
            dagKlienci.CurrentCell = dagKlienci(0, 0)
            dagKlienci.CurrentCell.Selected = True
        End If
    End Sub




    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        ZapamietajSzerokosciKolumn()
        '-------------------------------
        If DataSetKlienci.Tables(0).Rows.Count > 0 Then
            oKlient = GetTxtField(dagKlienci, 0)
        Else
            oKlient = ""        'IDENT       Text, 10
        End If
        Me.Close()
    End Sub




    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        If DataSetKlienci.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else

            Do While True
                oOperacja = "EDYTUJ"
                '====================================
                'OdswiezBaze()
                '====================================
                PobierzKlienci()
                Dim Form As New EdytujKatalogKlientow
                oAktualizuj = False
                Form.ShowDialog()
                If Not oAktualizuj Then
                Else
                    Dim foundRow As DataRow
                    ' Create an array for the key values to find.
                    Dim findTheseVals(0) As Object
                    findTheseVals(0) = oIdentKli
                    foundRow = DataSetKlienci.Tables(0).Rows.Find(findTheseVals)
                    ' sprawdz czy jest...
                    If Not (foundRow Is Nothing) Then

                        'dbSelectKlienci = "SELECT " &
                        '                             "IDENTKLIENTA AS NRKARTY, " &
                        '                             "NIP," &
                        '                             "NRTOFITXT," &
                        '                             "KARTAPAYBACK," &
                        '                             "NRYKARTPAYBACK," &
                        '                             "NAZWA1 AS NAZWAKLIENTA," &
                        '                             "MIEJSCOWOSC," &
                        '                             "KODPOCZTOWY," &
                        '                             "ULICA," &
                        '                             "NUMNRDOMU, " &
                        '"CASE WHEN [NUMNRDOMU]%2=0 THEN CAST('True' AS BIT) ELSE CAST('False' AS BIT) END AS PARZYSTE, " &
                        '                             "NRDOMU," &
                        '                             "IDDOMU," &
                        '                             "REJKLIENTA AS REJONKLIENTA," &
                        '                             "TELEFON," &
                        '                             "FAX," &
                        '                             "EMAIL," &
                        '                       "WYKAZURZADZEN," &
                        '                         "SPOSOBWYBORUDOSTAWCY," &
                        '                         "ZAINSTALOWANOMONITOR," &
                        '                         "LICZBAURZADZEN," &
                        '                         "LICZBAWYDRUKOW," &
                        '                         "POTENCJALDRUKU," &
                        '                       "AKTZAKUPMATEKSP," &
                        '                       "AKTROZLICZASTRONYDRUKU," &
                        '                       "AKTKORZYSTAZNAJMU," &
                        '                       "ZAINTROZLICZANIEMSTRONDRUKU," &
                        '                       "ZAINTKORZYSTANIEMZNAJMU," &
                        '                         "DATAWERYFSEGMENTACJI," &
                        '                         "PLANDLUGOOKRESOWY," &
                        '                         "PLANKROTKOOKRESOWY," &
                        '                               "RODZSPRZETUL," &
                        '                               "RODZSPRZETUA," &
                        '                               "RODZSPRZETUT," &
                        '                               "OFERTATZLOZENIA," &
                        '                               "OFERTAPLIK," &
                        '                               "PKONTAKT," &
                        '                               "SKUPUWAGI AS PROMOTORUWAGI," &
                        '                               "SPRZEDOPIEKUN AS OPIEKUN," &
                        '                               "SPRZEDOKONTAKTR AS OPIEKUNOSTKONTAKTROK," &
                        '                               "SPRZEDOKONTAKTT AS OPIEKUNOSTKONTAKTTYDZ," &
                        '                               "SPRZEDOKONTAKTD AS OPIEKUNOSTKONTAKTDZIEN," &
                        '                               "SPRZEDNKONTAKTR AS OPIEKUNKOLKONTAKTROK," &
                        '                               "SPRZEDNKONTAKTT AS OPIEKUNKOLKONTAKTTYDZ," &
                        '                               "SPRZEDNKONTAKTD AS OPIEKUNKOLKONTAKTDZIEN," &
                        '                               "SPRZEDUWAGI AS OPIEKUNUWAGI," &
                        '                               "KTOWPISAL," &
                        '                               "UWAGI," &
                        '                               "MARKER, " &
                        '                               "MARKERP, " &
                        '                               "AKTYWNY, " &
                        '                           "ZAKUPPOPROK, " &
                        '                           "USLUGIDOD, " &
                        '                           "RZURAP01, " &
                        '                           "RZURAP02, " &
                        '                           "RZURAP03, " &
                        '                           "RZURAP04, " &
                        '                           "RZURAP05, " &
                        '                           "RZURAP06, " &
                        '                           "RZURAP07, " &
                        '                           "RZURAP08, " &
                        '                           "RZURAP09, " &
                        '                                "WERSJA " &
                        '                            "FROM Klienci ORDER BY IDENTKLIENTA ASC"

                        'foundRow("NRKARTY") = oIdentKli
                        foundRow("NIP") = oNIPKli
                        foundRow("NRTOFITXT") = oNrTOFITxtKli
                        foundRow("KARTAPAYBACK") = oKartaPaybackKli
                        foundRow("NRYKARTPAYBACK") = oNryKartPaybackKli

                        foundRow("NAZWAKLIENTA") = oNazwa1Kli
                        foundRow("MIEJSCOWOSC") = oMiejscKli
                        foundRow("KODPOCZTOWY") = oKodPoczKli
                        foundRow("ULICA") = oUlicaKli
                        foundRow("NUMNRDOMU") = oNumNrDomuKli
                        'parzyste
                        foundRow("NRDOMU") = oNrDomuKli
                        foundRow("IDDOMU") = oIDDomuKli
                        foundRow("REJONKLIENTA") = oRejonKli
                        foundRow("TELEFON") = oTelefonKli
                        foundRow("FAX") = oFaxKli
                        foundRow("EMAIL") = oeMailKli

                        foundRow("BRANZA") = oBranzaKli
                        foundRow("PODBRANZA") = oPodBranzaKli
                        foundRow("LICZBAZATRUDNIONYCH") = oLiczbaZaktrudnionychKli
                        foundRow("LICZBAPRACZDALNIE") = oLiczbaPracZdalnieKli

                        foundRow("WYKAZURZADZEN") = oWykazUrzadzenKli
                        foundRow("SPOSOBWYBORUDOSTAWCY") = oSposobWyboruDostawcyKli
                        foundRow("ZAINSTALOWANOMONITOR") = oZainstalowanoMonitorKli
                        foundRow("LICZBAURZADZEN") = oLiczbaUrzadzenKli
                        foundRow("LICZBAWYDRUKOW") = oLiczbaWydrukowKli
                        foundRow("POTENCJALDRUKU") = oPotencjalDrukuKli
                        foundRow("AKTZAKUPMATEKSP") = oAktZakupMatEkspKli
                        foundRow("AKTROZLICZASTRONYDRUKU") = oAktRozliczaStronyDrukuKli
                        foundRow("AKTKORZYSTAZNAJMU") = oAktKorzystaZNajmuKli
                        foundRow("ZAINTROZLICZANIEMSTRONDRUKU") = oZaintRozliczaniemStronDrukuKli
                        foundRow("ZAINTKORZYSTANIEMZNAJMU") = oZaintKorzystaniemZNajmuKli
                        foundRow("DATAWERYFSEGMENTACJI") = oDataWeryfSegmentacjiKli
                        foundRow("PLANDLUGOOKRESOWY") = oPlanDlugookresowyKli
                        foundRow("PLANKROTKOOKRESOWY") = oPlanKrotkookresowyKli

                        foundRow("RODZSPRZETUL") = oRodzSprzetuLKli
                        foundRow("RODZSPRZETUA") = oRodzSprzetuAKli
                        foundRow("RODZSPRZETUT") = oRodzSprzetuTKli
                        foundRow("OFERTATZLOZENIA") = oOfertaTZlozeniaKli
                        foundRow("OFERTAPLIK") = oOfertaPlikKli
                        foundRow("PKONTAKT") = oPKontaktKli
                        foundRow("PROMOTORUWAGI") = oSkupUwagikli
                        foundRow("OPIEKUN") = oSprzedOpiekunkli
                        foundRow("OPIEKUNOSTKONTAKTROK") = oSprzedOKontaktRKli
                        foundRow("OPIEKUNOSTKONTAKTTYDZ") = Val(oSprzedOKontaktTKli)
                        foundRow("OPIEKUNOSTKONTAKTDZIEN") = oSprzedOKontaktDKli
                        foundRow("OPIEKUNKOLKONTAKTROK") = oSprzedNKontaktRKli
                        foundRow("OPIEKUNKOLKONTAKTTYDZ") = Val(oSprzedNKontaktTKli)
                        foundRow("OPIEKUNKOLKONTAKTDZIEN") = oSprzedNKontaktDKli
                        foundRow("OPIEKUNUWAGI") = oSprzedUwagiKli
                        foundRow("KTOWPISAL") = oWpisalKli
                        foundRow("UWAGI") = oUwagikli
                        foundRow("MARKER") = oMarkerKli
                        foundRow("MARKERP") = oMarkerPKli
                        foundRow("AKTYWNY") = oAktywnyKli

                        foundRow("ZAKUPPOPROK") = oZakupPopRokKli
                        foundRow("USLUGIDOD") = oUslugiDodKli
                        foundRow("RZURAP01") = oRZURap01
                        foundRow("RZURAP02") = oRZURap02
                        foundRow("RZURAP03") = oRZURap03
                        foundRow("RZURAP04") = oRZURap04
                        foundRow("RZURAP05") = oRZURap05
                        foundRow("RZURAP06") = oRZURap06
                        foundRow("RZURAP07") = oRZURap07
                        foundRow("RZURAP08") = oRZURap08
                        foundRow("RZURAP09") = oRZURap09

                        foundRow("WERSJA") = ZmienWersje(oWersjaKli) 'Wersja         Integer, 2

                        'wyswietl ilosc rekordow
                        stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString
                        AktualizujBaze()
                        'aktualizuj DataGrid
                        dagKlienci.Update()
                        '---------------
                        dagKlienci_CurrentCellChanged(sender, e)
                    End If
                End If
                '--------------------------
                Select Case oCoDalej
                    Case "STÓJ"
                        Exit Do
                    Case "NASTÊPNY"
                        If DataViewKlienci.Count() > 0 Then
                            If dagKlienci.CurrentCell.RowIndex < DataViewKlienci.Count() - 1 Then
                                dagKlienci.CurrentCell = dagKlienci(0, dagKlienci.CurrentCell.RowIndex + 1)
                                dagKlienci.CurrentCell.Selected = True
                            End If
                        End If
                    Case "POPRZEDNI"
                        If DataViewKlienci.Count() > 0 Then
                            If dagKlienci.CurrentCell.RowIndex > 0 Then
                                dagKlienci.CurrentCell = dagKlienci(0, dagKlienci.CurrentCell.RowIndex - 1)
                                dagKlienci.CurrentCell.Selected = True
                            End If
                        End If
                End Select
                '--------------------------
            Loop
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
        If _MaUprawnieniaPracownika() Or _MaUprawnieniaPracownikaUprzywilejowanego() Then
            MessageBox.Show("Nie masz uprawnieñ do usuwania klientów." & vbCrLf &
                "Jeœli to konieczne - poproœ Szefa lub Administratora.",
                "Przykro mi...",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            If DataSetKlienci.Tables(0).Rows.Count <= 0 Then
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
                    PobierzKlienci()
                    '====================================
                    'OdswiezBaze()
                    '====================================
                    Dim foundRow As DataRow
                    ' Create an array for the key values to find.
                    Dim findTheseVals(0) As Object
                    findTheseVals(0) = oIdentKli
                    foundRow = DataSetKlienci.Tables(0).Rows.Find(findTheseVals)
                    ' sprawdz czy jest...
                    If Not (foundRow Is Nothing) Then
                        foundRow.Delete()
                        stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString
                        AktualizujBaze()
                        '---------------
                        dagKlienci_CurrentCellChanged(sender, e)
                        DataSetKlienci.AcceptChanges()
                        '--------------------------------
                        'usun zapisy z tablic podleglych...
                        UsunOsoby(oIdentKli)
                        UsunKontakty(oIdentKli)
                        UsunObrotyMies(oIdentKli)
                        UsunObroty(oIdentKli)
                        UsunAkcjeSpec(oIdentKli)
                        UsunCoDalej(oIdentKli)
                    Else
                        MessageBox.Show("Nie znalaz³em w katalogu zapisu dla którego " & vbCrLf &
                            "NrKarty = " & oIdentKli & vbCrLf &
                            "Ktoœ musia³ ju¿ ten zapis usun¹æ...",
                            "Uwaga :",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        oOperacja = "DODAJ"
        InitKlienci()
        '----------------------------------------------
        ''generuj IdentKlienta
        'Dim objRow As DataRow
        'Dim LastIdent As Long = 0
        'For Each objRow In DataSetKlienci.Tables(0).Rows
        '    If LastIdent < Val(objRow.Item(0)) Then
        '        LastIdent = Val(objRow.Item(0))
        '    End If
        'Next
        'oIdentKli = Mid(Trim(Str(1000000 + LastIdent + 1)), 2, 6)
        '----------------------------------------------
        Do While True
            oAktualizuj = False
            Dim Form As New EdytujKatalogKlientow
            Form.ShowDialog()
            Form.Dispose()
            If Not oAktualizuj Then
                Exit Do
            Else
                '====================================
                'OdswiezBaze()
                '====================================
                Dim foundRow As DataRow
                Dim NewRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(0) As Object
                findTheseVals(0) = oIdentKli
                foundRow = DataSetKlienci.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "NrKarty = " & oIdentKli & vbCrLf &
                        "Nie mogê dopisaæ tego rekordu...",
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                    'Exit Do
                Else
                    'stworz nowy rekord
                    NewRow = DataSetKlienci.Tables(0).NewRow()

                    NewRow("NRKARTY") = oIdentKli
                    NewRow("NIP") = oNIPKli
                    NewRow("NRTOFITXT") = oNrTOFITxtKli
                    NewRow("KARTAPAYBACK") = oKartaPaybackKli
                    NewRow("NRYKARTPAYBACK") = oNryKartPaybackKli

                    NewRow("NAZWAKLIENTA") = oNazwa1Kli
                    NewRow("MIEJSCOWOSC") = oMiejscKli
                    NewRow("KODPOCZTOWY") = oKodPoczKli
                    NewRow("ULICA") = oUlicaKli
                    NewRow("NUMNRDOMU") = oNumNrDomuKli
                    'parzyste
                    NewRow("NRDOMU") = oNrDomuKli
                    NewRow("IDDOMU") = oIDDomuKli
                    NewRow("REJONKLIENTA") = oRejonKli
                    NewRow("TELEFON") = oTelefonKli
                    NewRow("FAX") = oFaxKli
                    NewRow("EMAIL") = oeMailKli

                    NewRow("BRANZA") = oBranzaKli
                    NewRow("PODBRANZA") = oPodBranzaKli
                    NewRow("LICZBAZATRUDNIONYCH") = oLiczbaZaktrudnionychKli
                    NewRow("LICZBAPRACZDALNIE") = oLiczbaPracZdalnieKli

                    NewRow("WYKAZURZADZEN") = oWykazUrzadzenKli
                    NewRow("SPOSOBWYBORUDOSTAWCY") = oSposobWyboruDostawcyKli
                    NewRow("ZAINSTALOWANOMONITOR") = oZainstalowanoMonitorKli
                    NewRow("LICZBAURZADZEN") = oLiczbaUrzadzenKli
                    NewRow("LICZBAWYDRUKOW") = oLiczbaWydrukowKli
                    NewRow("POTENCJALDRUKU") = oPotencjalDrukuKli
                    NewRow("AKTZAKUPMATEKSP") = oAktZakupMatEkspKli
                    NewRow("AKTROZLICZASTRONYDRUKU") = oAktRozliczaStronyDrukuKli
                    NewRow("AKTKORZYSTAZNAJMU") = oAktKorzystaZNajmuKli
                    NewRow("ZAINTROZLICZANIEMSTRONDRUKU") = oZaintRozliczaniemStronDrukuKli
                    NewRow("ZAINTKORZYSTANIEMZNAJMU") = oZaintKorzystaniemZNajmuKli
                    NewRow("DATAWERYFSEGMENTACJI") = oDataWeryfSegmentacjiKli
                    NewRow("PLANDLUGOOKRESOWY") = oPlanDlugookresowyKli
                    NewRow("PLANKROTKOOKRESOWY") = oPlanKrotkookresowyKli

                    NewRow("RODZSPRZETUL") = oRodzSprzetuLKli
                    NewRow("RODZSPRZETUA") = oRodzSprzetuAKli
                    NewRow("RODZSPRZETUT") = oRodzSprzetuTKli
                    NewRow("OFERTATZLOZENIA") = oOfertaTZlozeniaKli
                    NewRow("OFERTAPLIK") = oOfertaPlikKli
                    NewRow("PKONTAKT") = oPKontaktKli
                    NewRow("PROMOTORUWAGI") = oSkupUwagikli
                    NewRow("OPIEKUN") = oSprzedOpiekunkli
                    NewRow("OPIEKUNOSTKONTAKTROK") = oSprzedOKontaktRKli
                    NewRow("OPIEKUNOSTKONTAKTTYDZ") = Val(oSprzedOKontaktTKli)
                    NewRow("OPIEKUNOSTKONTAKTDZIEN") = oSprzedOKontaktDKli
                    NewRow("OPIEKUNKOLKONTAKTROK") = oSprzedNKontaktRKli
                    NewRow("OPIEKUNKOLKONTAKTTYDZ") = Val(oSprzedNKontaktTKli)
                    NewRow("OPIEKUNKOLKONTAKTDZIEN") = oSprzedNKontaktDKli
                    NewRow("OPIEKUNUWAGI") = oSprzedUwagiKli
                    NewRow("KTOWPISAL") = oWpisalKli
                    NewRow("UWAGI") = oUwagikli
                    NewRow("MARKER") = oMarkerKli
                    NewRow("MARKERP") = oMarkerPKli
                    NewRow("AKTYWNY") = oAktywnyKli

                    NewRow("ZAKUPPOPROK") = oZakupPopRokKli
                    NewRow("USLUGIDOD") = oUslugiDodKli
                    NewRow("RZURAP01") = oRZURap01
                    NewRow("RZURAP02") = oRZURap02
                    NewRow("RZURAP03") = oRZURap03
                    NewRow("RZURAP04") = oRZURap04
                    NewRow("RZURAP05") = oRZURap05
                    NewRow("RZURAP06") = oRZURap06
                    NewRow("RZURAP07") = oRZURap07
                    NewRow("RZURAP08") = oRZURap08
                    NewRow("RZURAP09") = oRZURap09

                    NewRow("WERSJA") = 1                     'Wersja         Integer, 2
                    'dodaj rekord do DataSet
                    DataSetKlienci.Tables(0).Rows.Add(NewRow)

                    'wyswietl ilosc rekordow
                    stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString
                    AktualizujBaze()
                    'aktualizuj DataGrid
                    dagKlienci.Update()
                    '---------------
                    dagKlienci_CurrentCellChanged(sender, e)
                    Exit Do
                End If
            End If
        Loop
    End Sub

    Private Sub cmdWysylaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWysylaj.Click
        Dim Form As New eMailing(DataViewKlienci, "", "")
        Form.ShowDialog()
    End Sub



    '**************************************
    '* Drukowanie
    '**************************************
    Dim TextL As String = ""
    Dim TextP As String = ""
    Dim i100 As Integer = 0
    Dim TxtLDl As Long = 0
    Dim TxtPDl As Long = 0


    Dim DrukujLinieLNr As Integer = 0
    Dim DrukujLiniePNr As Integer = 0
    Dim IleLiniiTrzebaWydrukowac As Integer = 0
    Dim IleLiniiL As Integer = 0
    Dim IleLiniiP As Integer = 0

    Dim drWiersz As String = ""     'liczniki iloœci wierszy poszcz kolumn wydruku

    Private Sub cmdDrukuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDrukuj.Click
        DrukujLinieLNr = 0
        DrukujLiniePNr = 0
        IleLiniiTrzebaWydrukowac = 0
        IleLiniiL = 0
        IleLiniiP = 0
        TextL = ""
        TextP = ""

        ' Tworzymy dokument i do³¹czmy obs³ugê zdarzenia PrintPage
        Dim MyDoc As New Softart.myPrintDocument
        MyDoc.DocumentName = cbxWydruki.SelectedItem.ToString
        MyDoc.LineNumber = 0
        MyDoc.PageNumber = 0
        MyDoc.NextRowNumber = 0

        'podlacz hendler zaleznie od wybranego wydruku
        'cbxWydruki.Items.Add("Zestawienie klientów (wydruk dla promotora)")
        'cbxWydruki.Items.Add("Zestawienie klientów (wydruk dla promotora - wersja skrócona)")
        'cbxWydruki.Items.Add("Zestawienie klientów (generator wydruków)")
        'cbxWydruki.Items.Add("Wydruk etykiet adresowych")
        'cbxWydruki.Items.Add("Wydruk karty klienta")
        'cbxWydruki.SelectedIndex = 0

        Select Case cbxWydruki.SelectedIndex
            Case 0
                MyDoc.DefaultPageSettings.Landscape = True
                AddHandler MyDoc.PrintPage, AddressOf MyDoc_Zestawienie
            Case 1
                MyDoc.DefaultPageSettings.Landscape = True
                AddHandler MyDoc.PrintPage, AddressOf MyDoc_ZestawienieOld
            Case 2
                oGeneratorWydruku = ""
                Dim form As New GeneratorWydrukow(DataSetKlienci, DataViewKlienci)
                form.ShowDialog()
                '------------------------------------
                If Len(oGeneratorWydruku) > 0 Then
                    MyDoc.DefaultPageSettings.Landscape = True
                    AddHandler MyDoc.PrintPage, AddressOf MyDoc_ZestawienieZGeneratora
                Else
                    Return
                End If
            Case 3
                MyDoc.DefaultPageSettings.Landscape = False
                AddHandler MyDoc.PrintPage, AddressOf MyDoc_Etykiety
            Case 4
                MyDoc.DefaultPageSettings.Landscape = False
                AddHandler MyDoc.PrintPage, AddressOf MyDoc_KartaKlienta
        End Select

        Try
            'wybor drukarki...
            Dim PDialog As New PrintDialog
            PDialog.Document = MyDoc
            PDialog.AllowPrintToFile = True
            PDialog.PrintToFile = False
            PDialog.AllowSelection = True
            PDialog.ShowHelp = True
            PDialog.ShowNetwork = True
            PDialog.UseEXDialog = True

            Dim Result As DialogResult = PDialog.ShowDialog()
            ' Drukuj po nacisnieciu OK
            If Result = System.Windows.Forms.DialogResult.OK Then
                ' This method returns immediately, before the print job starts.
                ' The PrintPage event will fire asynchronously.
                'MyDoc.Print()

                drWiersz = ""

                'podglad wydruku
                Dim dlgPreview As New PrintPreviewDialog
                dlgPreview.Document = MyDoc
                dlgPreview.WindowState = FormWindowState.Maximized
                dlgPreview.UseAntiAlias = True  'lepsza jakosc wydruku
                dlgPreview.ShowDialog()
            End If
        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1)
        Finally
        End Try
    End Sub







    '**************************************************************
    '** wydruk z bazy danych
    '**************************************************************



    Private Sub MyDoc_ZestawienieZGeneratora(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        ' Retrieve the document that sent this event.
        ' You could store the document in a class member variable,
        ' but this approach allows you to use the same event handler
        ' to handle multiple print documents at once.
        Dim Doc As Softart.myPrintDocument = CType(sender, PrintDocument)
        Dim Text As String = ""
        Dim Text2 As String = ""
        Dim LiRow As Integer = 0
        Dim ii As Integer = 0
        Dim wi As Integer = 0
        Dim wiMax As Integer = 0
        Dim NowyWiersz As Boolean = True

        e.PageSettings.Landscape = True

        ' Define the font and determine the line height.
        Dim FontSize As Integer = 8
        Dim FontSizeTitle As Integer = 12
        Dim MyFontTitle As New System.Drawing.Font("Arial", FontSizeTitle, FontStyle.Bold)
        Dim LineHeightTitle As Single = MyFontTitle.GetHeight(e.Graphics)
        Dim MyFont As New System.Drawing.Font("Arial", FontSize, FontStyle.Regular)
        Dim LineHeight As Single = MyFont.GetHeight(e.Graphics)
        Dim DrawingPen As New Pen(Color.Black, 1)

        ' Create variables to hold position on page.
        Dim x As Single = e.MarginBounds.Left
        Dim y As Single = e.MarginBounds.Top

        Dim PageWidth As Single = e.MarginBounds.Width
        Dim PageHeight As Single = e.MarginBounds.Height
        Dim PageLeft As Single = cm2pts(0.5)    '.MarginBounds.Left
        Dim PageTop As Single = e.MarginBounds.Top

        Dim Pos As Integer = 0
        Dim IleKolumn As Integer = GWIleKolumn(oGeneratorWydruku)
        Dim ik As Integer = 0
        Dim odLewej As Integer = 0
        Dim KolNazwa As String = ""
        Dim KolTyp As String = ""
        Dim KolRozmiar As Integer = 0
        Dim KolNaglowek As String = ""
        Dim KolIleWierszy As Integer = 0
        Dim KolWyrownanie As String = ""

        ' Naglowek strony - drukuj tytul wydruku
        Doc.PageNumber += 1
        Text = GWNaglowek(oGeneratorWydruku)
        y = PageTop
        x = PageLeft
        CentrTxt(Text, x, y, PageWidth, MyFontTitle, e)
        y += 1.5 * MyFontTitle.GetHeight(e.Graphics)

        'drukujemy glowke tabeli
        x = PageLeft
        odLewej = 0
        For ik = 1 To IleKolumn
            KolRozmiar = GWKolRozmiar(oGeneratorWydruku, ik)
            odLewej += KolRozmiar
            e.Graphics.DrawRectangle(DrawingPen, x, y, mm2pts(odLewej), CType(1.5 * MyFont.GetHeight(e.Graphics), Single))
        Next
        y += 0.5 * MyFont.GetHeight(e.Graphics)

        x = PageLeft
        odLewej = 0
        For ik = 1 To IleKolumn
            KolRozmiar = GWKolRozmiar(oGeneratorWydruku, ik)
            KolNaglowek = GWKolNaglowek(oGeneratorWydruku, ik)
            CentrTxt(KolNaglowek, PageLeft + mm2pts(odLewej), y, mm2pts(KolRozmiar), MyFont, e)
            odLewej += KolRozmiar
        Next
        y += 1.5 * MyFont.GetHeight(e.Graphics)

        ' petla do konca strony
        If Doc.NextRowNumber < DataViewKlienci.Count Then
            Do
                NowyWiersz = True
                If Len(drWiersz) > 0 Then
                    For ik = 1 To IleKolumn
                        If Mid(drWiersz, (ik - 1) * 8 + 1, 4) <> "0000" Then
                            'ok jest koleja linia z tego rekordu
                            NowyWiersz = False
                            Exit For
                        End If
                    Next
                End If

                If Not NowyWiersz Then
                    'ustaw kursor na poprzednim wierszu DataGrid - musimy skonczyc drukowanie
                    LiRow = Doc.NextRowNumber - 1
                    dagKlienci.CurrentCell = dagKlienci(0, LiRow)
                    PobierzKlienci()

                    x = PageLeft 'od marginesu...
                Else
                    'ustaw kursor na nowym wierszu DataGrid
                    LiRow = Doc.NextRowNumber
                    dagKlienci.CurrentCell = dagKlienci(0, LiRow)
                    PobierzKlienci()

                    x = PageLeft 'od marginesu...
                    Doc.LineNumber += 1
                    Doc.NextRowNumber += 1
                    y += 0.5 * MyFont.GetHeight(e.Graphics)

                    drWiersz = ""
                    'Analizuj ilosc wierszy w kazdej kolumnie
                    For ik = 1 To IleKolumn
                        KolNazwa = GWKolNazwa(oGeneratorWydruku, ik)
                        KolTyp = GWKolTyp(oGeneratorWydruku, ik)
                        KolRozmiar = GWKolRozmiar(oGeneratorWydruku, ik)
                        KolWyrownanie = GWKolWyrownanie(oGeneratorWydruku, ik)
                        KolIleWierszy = GWKolIleWierszy(oGeneratorWydruku, ik)

                        'pobierz zawartosc pola tekstowego
                        Select Case KolNazwa
                            Case "NRKARTY"
                                Text = oIdentKli
                            Case "NIP"
                                Text = oNIPKli
                            Case "NRTOFI"
                                Text = Trim(Str(oNrTOFIKli))
                            Case "NRTOFITXT"
                                Text = Trim(oNrTOFITxtKli)
                            Case "KARTAPAYBACK"
                                Text = IIf(oKartaPaybackKli, "T", "N")
                            Case "NRYKARTPAYBACK"
                                Text = Trim(oNryKartPaybackKli)


                            Case "NAZWAKLIENTA"
                                Text = oNazwa1Kli
                            Case "MIEJSCOWOSC"
                                Text = oMiejscKli
                            Case "KODPOCZTOWY"
                                Text = oKodPoczKli
                            Case "ULICA"
                                Text = oUlicaKli
                            Case "NUMNRDOMU"
                                Text = Trim(Str(oNumNrDomuKli))
                            Case "NRDOMU"
                                Text = oNrDomuKli
                            Case "IDDOMU"
                                Text = oIDDomuKli
                            Case "REJONKLIENTA"
                                Text = oRejonKli
                            Case "TELEFON"
                                Text = oTelefonKli
                            Case "FAX"
                                Text = oFaxKli
                            Case "EMAIL"
                                Text = oeMailKli

                            Case "BRANZA"
                                Text = oeMailKli
                            Case "PODBRANZA"
                                Text = oeMailKli
                            Case "LICZBAZATRUDNIONYCH"
                                Text = oLiczbaZaktrudnionychKli
                            Case "LICZBAPRACZDALNIE"
                                Text = oLiczbaPracZdalnieKli


                            Case "WYKAZURZADZEN"
                                Text = oWykazUrzadzenKli
                            Case "SPOSOBWYBORUDOSTAWCY"
                                Text = oSposobWyboruDostawcyKli
                            Case "ZAINSTALOWANOMONITOR"
                                Text = oZainstalowanoMonitorKli
                            Case "LICZBAURZADZEN"
                                Text = Trim(Str(oLiczbaUrzadzenKli))
                            Case "LICZBAWYDRUKOW"
                                Text = Trim(Str(oLiczbaWydrukowKli))
                            Case "POTENCJALDRUKU"
                                Text = oPotencjalDrukuKli
                            Case "AKTZAKUPMATEKSP"
                                Text = IIf(oAktZakupMatEkspKli, "T", "N")
                            Case "AKTROZLICZASTRONYDRUKU"
                                Text = IIf(oAktRozliczaStronyDrukuKli, "T", "N")
                            Case "AKTKORZYSTAZNAJMU"
                                Text = IIf(oAktKorzystaZNajmuKli, "T", "N")
                            Case "ZAINTROZLICZANIEMSTRONDRUKU"
                                Text = IIf(oZaintRozliczaniemStronDrukuKli, "T", "N")
                            Case "ZAINTKORZYSTANIEMZNAJMU"
                                Text = IIf(oZaintKorzystaniemZNajmuKli, "T", "N")
                            Case "DATAWERYFSEGMENTACJI"
                                Text = oDataWeryfSegmentacjiKli
                            Case "PLANDLUGOOKRESOWY"
                                Text = Trim(Str(oPlanDlugookresowyKli))
                            Case "PLANKROTKOOKRESOWY"
                                Text = Trim(Str(oPlanKrotkookresowyKli))

                            Case "RODZSPRZETUL"
                                Text = IIf(oRodzSprzetuLKli, "T", "N")
                            Case "RODZSPRZETUA"
                                Text = IIf(oRodzSprzetuAKli, "T", "N")
                            Case "RODZSPRZETUT"
                                Text = IIf(oRodzSprzetuTKli, "T", "N")
                            Case "OFERTATZLOZENIA"
                                Text = oOfertaTZlozeniaKli
                            Case "OFERTAPLIK"
                                Text = oOfertaPlikKli
                            Case "PKONTAKT"
                                Text = oPKontaktKli
                            Case "PROMOTORUWAGI"
                                Text = oSkupUwagikli
                            Case "OPIEKUN"
                                Text = oSprzedOpiekunkli
                            Case "OPIEKUNOSTKONTAKTROK"
                                Text = oSprzedOKontaktRKli
                            Case "OPIEKUNOSTKONTAKTTYDZ"
                                Text = oSprzedOKontaktTKli
                            Case "OPIEKUNOSTKONTAKTDZIEN"
                                Text = oSprzedOKontaktDKli
                            Case "OPIEKUNKOLKONTAKTROK"
                                Text = oSprzedNKontaktRKli
                            Case "OPIEKUNKOLKONTAKTTYDZ"
                                Text = oSprzedNKontaktTKli
                            Case "OPIEKUNKOLKONTAKTDZIEN"
                                Text = oSprzedNKontaktDKli
                            Case "OPIEKUNUWAGI"
                                Text = oSprzedUwagiKli
                            Case "KTOWPISAL"
                                Text = oWpisalKli
                            Case "UWAGI"
                                Text = oUwagikli

                            Case "MARKER"
                                Text = IIf(oMarkerKli, "T", "N")
                            Case "MARKERP"
                                Text = IIf(oMarkerPKli, "T", "N")
                            Case "AKTYWNY"
                                Text = IIf(oAktywnyKli, "T", "N")

                            Case "WERSJA"
                                Text = Trim(Str(oWersjaKli))

                            Case "KONTAKTYHANDLOWE"
                                IdentOstatniKontakt(oIdentKli)
                                Text = Trim(oUwagiKon)

                            Case "CODALEJ"
                                IdentCoDalej(oIdentKli)
                                Text = Trim(cdOPIS)

                            Case Else
                                Text = ""
                        End Select

                        'sprawdzamy iloœc wierszy tylko w kolumnie tekstowej
                        Select Case KolTyp
                            Case "System.Int32"
                                wi = 1
                            Case "System.Boolean"
                                wi = 1
                            Case "System.String"
                                'wi = IleLiniiMaText(Text, mm2pts(KolRozmiar), MyFont, e)
                                wi = IleLiniiPotrzebaNaText(Text, mm2pts(KolRozmiar), MyFont, e)
                            Case Else
                                wi = 1
                        End Select
                        drWiersz &= Microsoft.VisualBasic.Right(Str(10000 + wi), 4) &
                                    Microsoft.VisualBasic.Right(Str(10000 + wi), 4)
                    Next
                End If
                '-----------------------------------
                'drukujemy kolejn¹ linie z tego rekordu
                odLewej = 0
                For ik = 1 To IleKolumn
                    KolNazwa = GWKolNazwa(oGeneratorWydruku, ik)
                    KolTyp = GWKolTyp(oGeneratorWydruku, ik)
                    KolRozmiar = GWKolRozmiar(oGeneratorWydruku, ik)
                    KolWyrownanie = GWKolWyrownanie(oGeneratorWydruku, ik)
                    KolIleWierszy = GWKolIleWierszy(oGeneratorWydruku, ik)

                    Select Case KolNazwa
                        Case "NRKARTY"
                            Text = oIdentKli
                        Case "NIP"
                            Text = oNIPKli
                        Case "NRTOFI"
                            Text = Trim(Str(oNrTOFIKli))
                        Case "NRTOFITXT"
                            Text = Trim(oNrTOFITxtKli)
                        Case "KARTAPAYBACK"
                            Text = IIf(oKartaPaybackKli, "T", "N")
                        Case "NRYKARTPAYBACK"
                            Text = Trim(oNryKartPaybackKli)

                        Case "NAZWAKLIENTA"
                            Text = oNazwa1Kli
                        Case "MIEJSCOWOSC"
                            Text = oMiejscKli
                        Case "KODPOCZTOWY"
                            Text = oKodPoczKli
                        Case "ULICA"
                            Text = oUlicaKli
                        Case "NUMNRDOMU"
                            Text = Trim(Str(oNumNrDomuKli))
                        Case "NRDOMU"
                            Text = oNrDomuKli
                        Case "IDDOMU"
                            Text = oIDDomuKli
                        Case "REJONKLIENTA"
                            Text = oRejonKli
                        Case "TELEFON"
                            Text = oTelefonKli
                        Case "FAX"
                            Text = oFaxKli
                        Case "EMAIL"
                            Text = oeMailKli


                        Case "BRANZA"
                            Text = oeMailKli
                        Case "PODBRANZA"
                            Text = oeMailKli
                        Case "LICZBAZATRUDNIONYCH"
                            Text = oLiczbaZaktrudnionychKli
                        Case "LICZBAPRACZDALNIE"
                            Text = oLiczbaPracZdalnieKli


                        Case "WYKAZURZADZEN"
                            Text = oWykazUrzadzenKli
                        Case "SPOSOBWYBORUDOSTAWCY"
                            Text = oSposobWyboruDostawcyKli
                        Case "ZAINSTALOWANOMONITOR"
                            Text = oZainstalowanoMonitorKli
                        Case "LICZBAURZADZEN"
                            Text = Trim(Str(oLiczbaUrzadzenKli))
                        Case "LICZBAWYDRUKOW"
                            Text = Trim(Str(oLiczbaWydrukowKli))
                        Case "POTENCJALDRUKU"
                            Text = oPotencjalDrukuKli
                        Case "AKTZAKUPMATEKSP"
                            Text = IIf(oAktZakupMatEkspKli, "T", "N")
                        Case "AKTROZLICZASTRONYDRUKU"
                            Text = IIf(oAktRozliczaStronyDrukuKli, "T", "N")
                        Case "AKTKORZYSTAZNAJMU"
                            Text = IIf(oAktKorzystaZNajmuKli, "T", "N")
                        Case "ZAINTROZLICZANIEMSTRONDRUKU"
                            Text = IIf(oZaintRozliczaniemStronDrukuKli, "T", "N")
                        Case "ZAINTKORZYSTANIEMZNAJMU"
                            Text = IIf(oZaintKorzystaniemZNajmuKli, "T", "N")
                        Case "DATAWERYFSEGMENTACJI"
                            Text = oDataWeryfSegmentacjiKli
                        Case "PLANDLUGOOKRESOWY"
                            Text = Trim(Str(oPlanDlugookresowyKli))
                        Case "PLANKROTKOOKRESOWY"
                            Text = Trim(Str(oPlanKrotkookresowyKli))


                        Case "RODZSPRZETUL"
                            Text = IIf(oRodzSprzetuLKli, "T", "N")
                        Case "RODZSPRZETUA"
                            Text = IIf(oRodzSprzetuAKli, "T", "N")
                        Case "RODZSPRZETUT"
                            Text = IIf(oRodzSprzetuTKli, "T", "N")

                        Case "OFERTATZLOZENIA"
                            Text = oOfertaTZlozeniaKli
                        Case "OFERTAPLIK"
                            Text = oOfertaPlikKli
                        Case "PKONTAKT"
                            Text = oPKontaktKli
                        Case "PROMOTORUWAGI"
                            Text = oSkupUwagikli
                        Case "OPIEKUN"
                            Text = oSprzedOpiekunkli
                        Case "OPIEKUNOSTKONTAKTROK"
                            Text = oSprzedOKontaktRKli
                        Case "OPIEKUNOSTKONTAKTTYDZ"
                            Text = oSprzedOKontaktTKli
                        Case "OPIEKUNOSTKONTAKTDZIEN"
                            Text = oSprzedOKontaktDKli
                        Case "OPIEKUNKOLKONTAKTROK"
                            Text = oSprzedNKontaktRKli
                        Case "OPIEKUNKOLKONTAKTTYDZ"
                            Text = oSprzedNKontaktTKli
                        Case "OPIEKUNKOLKONTAKTDZIEN"
                            Text = oSprzedNKontaktDKli
                        Case "OPIEKUNUWAGI"
                            Text = oSprzedUwagiKli
                        Case "KTOWPISAL"
                            Text = oWpisalKli
                        Case "UWAGI"
                            Text = oUwagikli
                        Case "MARKER"
                            Text = IIf(oMarkerKli, "T", "N")
                        Case "MARKERP"
                            Text = IIf(oMarkerPKli, "T", "N")
                        Case "WERSJA"
                            Text = Trim(Str(oWersjaKli))

                        Case "KONTAKTYHANDLOWE"
                            IdentOstatniKontakt(oIdentKli)
                            Text = Trim(oUwagiKon)

                        Case "CODALEJ"
                            IdentCoDalej(oIdentKli)
                            Text = Trim(cdOPIS)

                        Case Else
                            Text = ""
                    End Select

                    wi = CInt(Mid(drWiersz, (ik - 1) * 8 + 1, 4))
                    wiMax = CInt(Mid(drWiersz, (ik - 1) * 8 + 5, 4))
                    If wi > 0 Then
                        Select Case KolTyp
                            Case "System.Int32"
                                Text2 = Text
                            Case "System.Boolean"
                                Text2 = Text2
                            Case "System.String"
                                'Text2 = LiniaTxtNr(Text, wiMax - wi + 1, mm2pts(KolRozmiar), MyFont, e)
                                Text2 = DajLinieTextuNr(Text, wiMax - wi + 1, mm2pts(KolRozmiar), MyFont, e)
                            Case Else
                                Text2 = Text
                        End Select

                        If Len(Text2) > 0 Then
                            Select Case KolWyrownanie
                                Case "Do lewej"
                                    LeftTxt(Text2, x + mm2pts(odLewej), y, mm2pts(KolRozmiar), MyFont, e)
                                Case "Do prawej"
                                    RightTxt(Text2, x + mm2pts(odLewej), y, mm2pts(KolRozmiar), MyFont, e)
                                Case Else
                                    CentrTxt(Text2, x + mm2pts(odLewej), y, mm2pts(KolRozmiar), MyFont, e)
                            End Select
                        End If

                        'zapamietaj nowy licznik linii
                        If ik = 1 Then
                            drWiersz = Microsoft.VisualBasic.Right(Str(10000 + wi - 1), 4) & Microsoft.VisualBasic.Right(Str(10000 + wiMax), 4) & Mid(drWiersz, ik * 8 + 1)
                        Else
                            drWiersz = Mid(drWiersz, 1, (ik - 1) * 8) & Microsoft.VisualBasic.Right(Str(10000 + wi - 1), 4) & Microsoft.VisualBasic.Right(Str(10000 + wiMax), 4) & Mid(drWiersz, ik * 8 + 1)
                        End If
                    End If
                    odLewej += KolRozmiar
                Next
                y += LineHeight

                'e.Graphics.DrawLine(DrawingPen, x, y + CType(0.2 * LineHeight, Single), x + CType(25.0 * (100 / 2.54), Single), y + CType(0.5 * LineHeight, Single))
                'y += CType(0.5 * LineHeight, Single)
            Loop Until ((y + LineHeight) > e.MarginBounds.Bottom) Or (Doc.NextRowNumber >= DataViewKlienci.Count)
        End If

        If (Doc.NextRowNumber < DataViewKlienci.Count) Then
            ' There is still at least one more page.
            ' Signal this event to fire again.
            e.HasMorePages = True
        Else
            ' Printing is complete.
            y += LineHeight
            x = PageLeft 'od marginesu...
            Text = "Koniec zestawienia na pozycji " & Str(Doc.LineNumber)
            e.Graphics.DrawString(Text, MyFont, Brushes.Black, x + CType(0.2 * (100 / 2.54), Single), y)

            Doc.PageNumber = 0
            Doc.LineNumber = 0
            Doc.NextRowNumber = 0

            drWiersz = ""
            DrukujLinieLNr = 0
            DrukujLiniePNr = 0
            IleLiniiTrzebaWydrukowac = 0
            IleLiniiL = 0
            IleLiniiP = 0
            TextL = ""
            TextP = ""

            e.HasMorePages = False
        End If

    End Sub



    '**************************************************************
    '** wydruk z bazy danych
    '**************************************************************



    'Private Sub MyDoc_Zestawienie_______(ByVal sender As Object, ByVal e As PrintPageEventArgs)
    '    ' Retrieve the document that sent this event.
    '    ' You could store the document in a class member variable,
    '    ' but this approach allows you to use the same event handler
    '    ' to handle multiple print documents at once.
    '    Dim Doc As Softart.myPrintDocument = CType(sender, PrintDocument)
    '    Dim Text As String
    '    Dim LiRow As Integer

    '    Dim Uwagi As String = ""
    '    Dim Uwagi1 As String = ""
    '    Dim Uwagi2 As String = ""
    '    Dim Uwagi3 As String = ""
    '    Dim Uwagi4 As String = ""
    '    Dim Uwagi5 As String = ""
    '    Dim Uwagi6 As String = ""
    '    Dim UwagiX As String = ""
    '    Dim ii As Integer = 0

    '    e.PageSettings.Landscape = True

    '    ' Define the font and determine the line height.
    '    Dim FontSize As Integer = 8
    '    Dim FontSizeTitle As Integer = 12
    '    Dim MyFontTitle As New System.Drawing.Font("Arial", FontSizeTitle, FontStyle.Bold)
    '    Dim LineHeightTitle As Single = MyFontTitle.GetHeight(e.Graphics)
    '    Dim MyFont As New System.Drawing.Font("Arial", FontSize, FontStyle.Regular)
    '    Dim LineHeight As Single = MyFont.GetHeight(e.Graphics)
    '    Dim DrawingPen As New Pen(Color.Black, 1)

    '    ' Create variables to hold position on page.
    '    Dim x As Single = e.MarginBounds.Left
    '    Dim y As Single = e.MarginBounds.Top

    '    Dim PageWidth As Single = e.MarginBounds.Width
    '    Dim PageHeight As Single = e.MarginBounds.Height
    '    Dim PageLeft As Single = e.MarginBounds.Left
    '    Dim PageTop As Single = e.MarginBounds.Top

    '    Dim words() As String
    '    Dim splitChars() As Char = {","c}
    '    Dim Tofi1 As String = ""
    '    Dim Tofi2 As String = ""
    '    Dim Tofi3 As String = ""
    '    Dim Tofi4 As String = ""
    '    Dim Tofi5 As String = ""
    '    Dim Tofi6 As String = ""

    '    ' Naglowek strony
    '    Doc.PageNumber += 1
    '    Text = "Zestawienie klient¡w"
    '    y = e.MarginBounds.Top
    '    x = PageLeft + (PageWidth - e.Graphics.MeasureString(Text, MyFontTitle).Width) / 2
    '    e.Graphics.DrawString(Text, MyFontTitle, Brushes.Black, x, y)
    '    y += 2 * MyFontTitle.GetHeight(e.Graphics)

    '    x = PageLeft
    '    e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(1.3), CType(5.5 * MyFont.GetHeight(e.Graphics), Single))
    '    e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(8.0), CType(5.5 * MyFont.GetHeight(e.Graphics), Single))
    '    e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(13.0), CType(5.5 * MyFont.GetHeight(e.Graphics), Single))
    '    e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(14.0), CType(5.5 * MyFont.GetHeight(e.Graphics), Single))
    '    e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(15.0), CType(5.5 * MyFont.GetHeight(e.Graphics), Single))
    '    e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(25.0), CType(5.5 * MyFont.GetHeight(e.Graphics), Single))

    '    e.Graphics.DrawString("NrKarty", MyFont, Brushes.Black, x + cm2pts(0.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Nazwa", MyFont, Brushes.Black, x + cm2pts(1.4), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Sprzeda¬-wartoŠ", MyFont, Brushes.Black, x + cm2pts(8.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Sprzeda¬-wartoŠ", MyFont, Brushes.Black, x + cm2pts(8.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Skup", MyFont, Brushes.Black, x + cm2pts(13.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Sprz", MyFont, Brushes.Black, x + cm2pts(14.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Uwagi", MyFont, Brushes.Black, x + cm2pts(15.1), y + CType(0.1 * (100 / 2.54), Single))
    '    y += MyFont.GetHeight(e.Graphics)

    '    e.Graphics.DrawString("NrTofi", MyFont, Brushes.Black, x + cm2pts(0.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Osoba kontaktowa", MyFont, Brushes.Black, x + cm2pts(1.4), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Sprzeda¬-Plan", MyFont, Brushes.Black, x + cm2pts(8.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Ost.", MyFont, Brushes.Black, x + cm2pts(13.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Ost..", MyFont, Brushes.Black, x + cm2pts(14.1), y + CType(0.1 * (100 / 2.54), Single))
    '    y += MyFont.GetHeight(e.Graphics)

    '    e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(0.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("MiejscowoŠ", MyFont, Brushes.Black, x + cm2pts(1.4), y + CType(0.1 * (100 / 2.54), Single))
    '    RightTxt("Rejon", x + cm2pts(1.4), y + CType(0.1 * (100 / 2.54), Single), cm2pts(6.6), MyFont, e)
    '    e.Graphics.DrawString("Sprzeda¬-Ost.transakcje", MyFont, Brushes.Black, x + cm2pts(8.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Nast.", MyFont, Brushes.Black, x + cm2pts(13.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Nast.", MyFont, Brushes.Black, x + cm2pts(14.1), y + CType(0.1 * (100 / 2.54), Single))
    '    y += MyFont.GetHeight(e.Graphics)

    '    e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(0.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Ulica", MyFont, Brushes.Black, x + cm2pts(1.4), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Dzia-ania akcyjne", MyFont, Brushes.Black, x + cm2pts(8.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(13.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(14.1), y + CType(0.1 * (100 / 2.54), Single))
    '    y += MyFont.GetHeight(e.Graphics)

    '    e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(0.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Karty Payback", MyFont, Brushes.Black, x + cm2pts(1.4), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("Czyszczenie", MyFont, Brushes.Black, x + cm2pts(8.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(13.1), y + CType(0.1 * (100 / 2.54), Single))
    '    e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(14.1), y + CType(0.1 * (100 / 2.54), Single))
    '    y += 2 * MyFont.GetHeight(e.Graphics)

    '    ' petla do konca strony
    '    If Doc.NextRowNumber < DataViewKlienci.Count Then
    '        Do
    '            'ustaw kursor na wierszu DataGrid
    '            LiRow = Doc.NextRowNumber
    '            dagKlienci.CurrentCell = New DataGridCell(LiRow, 0)
    '            PobierzKlienci()
    '            '------------------------------------------
    '            Uwagi1 = ""
    '            Uwagi2 = ""
    '            Uwagi3 = ""
    '            Uwagi4 = ""
    '            Uwagi5 = ""
    '            Uwagi6 = ""
    '            Uwagi = oSprzedUwagiKli

    '            ii = InStr(Uwagi, vbLf)
    '            If ii > 0 Then
    '                Uwagi1 = Mid(Uwagi, 1, ii - 1)
    '                Uwagi = Mid(Uwagi, ii + 1)
    '                ii = InStr(Uwagi, vbLf)
    '                If ii > 0 Then
    '                    Uwagi2 = Mid(Uwagi, 1, ii - 1)
    '                    Uwagi = Mid(Uwagi, ii + 1)
    '                    ii = InStr(Uwagi, vbLf)
    '                    If ii > 0 Then
    '                        Uwagi3 = Mid(Uwagi, 1, ii - 1)
    '                        Uwagi = Mid(Uwagi, ii + 1)
    '                        ii = InStr(Uwagi, vbLf)
    '                        If ii > 0 Then
    '                            Uwagi4 = Mid(Uwagi, 1, ii - 1)
    '                            Uwagi = Mid(Uwagi, ii + 1)
    '                            ii = InStr(Uwagi, vbLf)
    '                            If ii > 0 Then
    '                                Uwagi5 = Mid(Uwagi, 1, ii - 1)
    '                                Uwagi = Mid(Uwagi, ii + 1)
    '                                ii = InStr(Uwagi, vbLf)
    '                                If ii > 0 Then
    '                                    Uwagi6 = Mid(Uwagi, 1, ii - 1)
    '                                    Uwagi = Mid(Uwagi, ii + 1)
    '                                Else
    '                                    Uwagi6 = Uwagi
    '                                End If
    '                            Else
    '                                Uwagi5 = Uwagi
    '                            End If
    '                        Else
    '                            Uwagi4 = Uwagi
    '                        End If
    '                    Else
    '                        Uwagi3 = Uwagi
    '                    End If
    '                Else
    '                    Uwagi2 = Uwagi
    '                End If
    '            Else
    '                Uwagi1 = Uwagi
    '            End If





    '            'sprawdz czy sie zmiesci
    '            UwagiX = ""
    '            Do While e.Graphics.MeasureString(Uwagi1, MyFont).Width > CType(10 * (100 / 2.54), Single)
    '                UwagiX = Mid(Uwagi1, Len(Uwagi1), 1) + UwagiX
    '                Uwagi1 = Mid(Uwagi1, 1, Len(Uwagi1) - 1)
    '            Loop
    '            If Len(UwagiX) > 0 Then
    '                Uwagi6 = Uwagi5
    '                Uwagi5 = Uwagi4
    '                Uwagi4 = Uwagi3
    '                Uwagi3 = Uwagi2
    '                Uwagi2 = UwagiX
    '                UwagiX = ""
    '            End If

    '            Do While e.Graphics.MeasureString(Uwagi2, MyFont).Width > CType(10 * (100 / 2.54), Single)
    '                UwagiX = Mid(Uwagi2, Len(Uwagi2), 1) + UwagiX
    '                Uwagi2 = Mid(Uwagi2, 1, Len(Uwagi2) - 1)
    '            Loop
    '            If Len(UwagiX) > 0 Then
    '                Uwagi6 = Uwagi5
    '                Uwagi5 = Uwagi4
    '                Uwagi4 = Uwagi3
    '                Uwagi3 = UwagiX
    '                UwagiX = ""
    '            End If

    '            Do While e.Graphics.MeasureString(Uwagi3, MyFont).Width > CType(10 * (100 / 2.54), Single)
    '                UwagiX = Mid(Uwagi3, Len(Uwagi3), 1) + UwagiX
    '                Uwagi3 = Mid(Uwagi3, 1, Len(Uwagi3) - 1)
    '            Loop
    '            If Len(UwagiX) > 0 Then
    '                Uwagi6 = Uwagi5
    '                Uwagi5 = Uwagi4
    '                Uwagi4 = UwagiX
    '                UwagiX = ""
    '            End If

    '            Do While e.Graphics.MeasureString(Uwagi4, MyFont).Width > CType(10 * (100 / 2.54), Single)
    '                UwagiX = Mid(Uwagi4, Len(Uwagi4), 1) + UwagiX
    '                Uwagi4 = Mid(Uwagi4, 1, Len(Uwagi4) - 1)
    '            Loop
    '            If Len(UwagiX) > 0 Then
    '                Uwagi6 = Uwagi5
    '                Uwagi5 = UwagiX
    '                UwagiX = ""
    '            End If

    '            Do While e.Graphics.MeasureString(Uwagi5, MyFont).Width > CType(10 * (100 / 2.54), Single)
    '                UwagiX = Mid(Uwagi5, Len(Uwagi5), 1) + UwagiX
    '                Uwagi5 = Mid(Uwagi5, 1, Len(Uwagi5) - 1)
    '            Loop
    '            If Len(UwagiX) > 0 Then
    '                Uwagi6 = Uwagi5
    '                UwagiX = ""
    '            End If

    '            Do While e.Graphics.MeasureString(Uwagi6, MyFont).Width > CType(10 * (100 / 2.54), Single)
    '                UwagiX = Mid(Uwagi6, Len(Uwagi6), 1) + UwagiX
    '                Uwagi6 = Mid(Uwagi6, 1, Len(Uwagi6) - 1)
    '            Loop
    '            If Len(UwagiX) > 0 Then
    '                UwagiX = ""
    '            End If







    '            Tofi1 = ""
    '            Tofi2 = ""
    '            Tofi3 = ""
    '            Tofi4 = ""
    '            Tofi5 = ""
    '            Tofi6 = ""
    '            words = oNrTOFITxtKli.Split(splitChars, StringSplitOptions.None)
    '            If words.Length = 0 Then
    '            ElseIf words.Length = 1 Then
    '                Tofi1 = oNrTOFITxtKli
    '            ElseIf words.Length = 2 Then
    '                Tofi1 = words(0)
    '                Tofi2 = words(1)
    '            ElseIf words.Length = 3 Then
    '                Tofi1 = words(0)
    '                Tofi2 = words(1)
    '                Tofi3 = words(2)
    '            ElseIf words.Length = 4 Then
    '                Tofi1 = words(0)
    '                Tofi2 = words(1)
    '                Tofi3 = words(2)
    '                Tofi4 = words(3)
    '            ElseIf words.Length = 5 Then
    '                Tofi1 = words(0)
    '                Tofi2 = words(1)
    '                Tofi3 = words(2)
    '                Tofi4 = words(3)
    '                Tofi5 = words(4)
    '            Else
    '                Tofi1 = words(0)
    '                Tofi2 = words(1)
    '                Tofi3 = words(2)
    '                Tofi4 = words(3)
    '                Tofi5 = words(4)
    '                Tofi6 = words(5)
    '            End If





    '            x = PageLeft 'od marginesu...
    '            Doc.LineNumber += 1
    '            Doc.NextRowNumber += 1

    '            'wiersz 1
    '            LeftTxt(oIdentKli, x + cm2pts(0.0), y, cm2pts(1.3), MyFont, e)
    '            LeftTxt(oNazwa1Kli, x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
    '            LeftTxt("", x + cm2pts(8.0), y, cm2pts(5.0), MyFont, e)
    '            
    '            LeftTxt(oSprzedOKontaktTKli & "/" & Mid(oSprzedOKontaktDKli, 1, 2), x + cm2pts(14.0), y, cm2pts(1.0), MyFont, e)
    '            LeftTxt(Uwagi1, x + cm2pts(15.0), y, cm2pts(10.0), MyFont, e)
    '            y += LineHeight

    '            'wiersz 2
    '            LeftTxt(Tofi1, x + cm2pts(0.0), y, cm2pts(1.3), MyFont, e)
    '            LeftTxt(Trim(OsobaKontaktowa(oIdentKli)), x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
    '            LeftTxt("", x + cm2pts(8.0), y, cm2pts(5.0), MyFont, e)
    '            LeftTxt(oSkupNKontaktTKli & "/" & Mid(oSkupNKontaktDKli, 1, 2), x + cm2pts(13.0), y, cm2pts(1.0), MyFont, e)
    '            LeftTxt(oSprzedNKontaktTKli & "/" & Mid(oSprzedNKontaktDKli, 1, 2), x + cm2pts(14.0), y, cm2pts(1.0), MyFont, e)
    '            LeftTxt(Uwagi2, x + cm2pts(15.0), y, cm2pts(10.0), MyFont, e)
    '            y += LineHeight

    '            'wiersz 3
    '            LeftTxt(Tofi2, x + cm2pts(0.0), y, cm2pts(1.3), MyFont, e)
    '            LeftTxt(Trim(oKodPoczKli) + " " + Trim(oMiejscKli), x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
    '            RightTxt(Trim(oRejonKli), x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
    '            LeftTxt(OstTransakcje(oIdentKli, True), x + cm2pts(8.0), y, cm2pts(5.0), MyFont, e)
    '            LeftTxt("", x + cm2pts(13.0), y, cm2pts(1.0), MyFont, e)
    '            LeftTxt("", x + cm2pts(14.0), y, cm2pts(1.0), MyFont, e)
    '            LeftTxt(Uwagi3, x + cm2pts(15.0), y, cm2pts(10.0), MyFont, e)
    '            y += LineHeight

    '            'wiersz 4
    '            If Len(Trim(Tofi3)) > 0 Or _
    '                (Len(Trim(oUlicaKli)) > 0 Or Len(Trim(Str(oNumNrDomuKli))) > 0 Or Len(Trim(oIDDomuKli)) > 0) Or _
    '                Len(Trim("")) > 0 Or _
    '                Len(Trim(Uwagi4)) > 0 Then

    '                LeftTxt(Tofi3, x + cm2pts(0.0), y, cm2pts(1.3), MyFont, e)
    '                If Len(oIDDomuKli) > 0 Then
    '                    LeftTxt(Trim(oUlicaKli) & " " & Trim(Str(oNumNrDomuKli)) & " " & Trim(oIDDomuKli), x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
    '                Else
    '                    LeftTxt(Trim(oUlicaKli) & " " & Trim(Str(oNumNrDomuKli)), x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
    '                End If
    '                LeftTxt("", x + cm2pts(8.0), y, cm2pts(5.0), MyFont, e)
    '                LeftTxt("", x + cm2pts(13.0), y, cm2pts(1.0), MyFont, e)
    '                LeftTxt("", x + cm2pts(14.0), y, cm2pts(1.0), MyFont, e)
    '                LeftTxt(Uwagi4, x + cm2pts(15.0), y, cm2pts(10.0), MyFont, e)
    '                y += LineHeight
    '            Else
    '                'jesli jest wiersz 5 - drukuj pusty
    '                If Len(Trim(Tofi4)) > 0 Or _
    '                    Len(Trim(oNryKartPaybackKli)) > 0 Or _
    '                    Len(Trim("")) > 0 Or _
    '                    Len(Trim(Uwagi5)) > 0 Then
    '                    ' pusta linia
    '                    y += LineHeight
    '                End If
    '            End If

    '            'wiersz 5
    '            If Len(Trim(Tofi4)) > 0 Or _
    '                Len(Trim(oNryKartPaybackKli)) > 0 Or _
    '                Len(Trim("")) > 0 Or _
    '                Len(Trim(Uwagi5)) > 0 Then

    '                LeftTxt(Tofi4, x + cm2pts(0.0), y, cm2pts(1.3), MyFont, e)
    '                LeftTxt(oNryKartPaybackKli, x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
    '                LeftTxt("", x + cm2pts(8.0), y, cm2pts(5.0), MyFont, e)
    '                LeftTxt("", x + cm2pts(13.0), y, cm2pts(1.0), MyFont, e)
    '                LeftTxt("", x + cm2pts(14.0), y, cm2pts(1.0), MyFont, e)
    '                LeftTxt(Uwagi5, x + cm2pts(15.0), y, cm2pts(10.0), MyFont, e)
    '                y += LineHeight
    '            End If

    '            e.Graphics.DrawLine(DrawingPen, x, y + CType(0.2 * LineHeight, Single), x + CType(25.0 * (100 / 2.54), Single), y + CType(0.5 * LineHeight, Single))
    '            y += CType(0.5 * LineHeight, Single)
    '        Loop Until ((y + LineHeight) > e.MarginBounds.Bottom) Or (Doc.NextRowNumber >= DataViewKlienci.Count)
    '    End If

    '    If (Doc.NextRowNumber < DataViewKlienci.Count) Then
    '        ' There is still at least one more page.
    '        ' Signal this event to fire again.
    '        e.HasMorePages = True
    '    Else
    '        ' Printing is complete.
    '        y += LineHeight
    '        x = PageLeft 'od marginesu...
    '        Text = "Koniec zestawienia na pozycji " & Str(Doc.LineNumber)
    '        e.Graphics.DrawString(Text, MyFont, Brushes.Black, x + CType(0.2 * (100 / 2.54), Single), y)

    '        Doc.PageNumber = 0
    '        Doc.LineNumber = 0
    '        Doc.NextRowNumber = 0

    '        DrukujLinieLNr = 0
    '        DrukujLiniePNr = 0
    '        IleLiniiTrzebaWydrukowac = 0
    '        IleLiniiL = 0
    '        IleLiniiP = 0
    '        TextL = ""
    '        TextP = ""

    '        e.HasMorePages = False
    '    End If

    'End Sub









    '**************************************************************
    '** wydruk z bazy danych
    '**************************************************************

    Private Sub MyDoc_ZestawienieOld(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        ' Retrieve the document that sent this event.
        ' You could store the document in a class member variable,
        ' but this approach allows you to use the same event handler
        ' to handle multiple print documents at once.
        Dim Doc As Softart.myPrintDocument = CType(sender, PrintDocument)
        Dim Text As String
        Dim LiRow As Integer

        Dim Uwagi As String = ""
        Dim Uwagi1 As String = ""
        Dim Uwagi2 As String = ""
        Dim Uwagi3 As String = ""
        Dim Uwagi4 As String = ""
        Dim Uwagi5 As String = ""
        Dim Uwagi6 As String = ""
        Dim UwagiX As String = ""
        Dim ii As Integer = 0

        e.PageSettings.Landscape = True

        ' Define the font and determine the line height.
        Dim FontSize As Integer = 8
        Dim FontSizeTitle As Integer = 12
        Dim MyFontTitle As New System.Drawing.Font("Arial", FontSizeTitle, FontStyle.Bold)
        Dim LineHeightTitle As Single = MyFontTitle.GetHeight(e.Graphics)
        Dim MyFont As New System.Drawing.Font("Arial", FontSize, FontStyle.Regular)
        Dim LineHeight As Single = MyFont.GetHeight(e.Graphics)
        Dim DrawingPen As New Pen(Color.Black, 1)

        ' Create variables to hold position on page.
        Dim x As Single = e.MarginBounds.Left
        Dim y As Single = e.MarginBounds.Top

        Dim PageWidth As Single = cm2pts(28)
        Dim PageHeight As Single = cm2pts(19)
        Dim PageLeft As Single = cm2pts(0.5)
        Dim PageTop As Single = cm2pts(1)

        Dim words() As String
        Dim splitChars() As Char = {","c}
        Dim Tofi1 As String = ""
        Dim Tofi2 As String = ""
        Dim Tofi3 As String = ""
        Dim Tofi4 As String = ""
        Dim Tofi5 As String = ""
        Dim Tofi6 As String = ""

        ' Naglowek strony
        Doc.PageNumber += 1
        Text = "Zestawienie klientów"
        y = e.MarginBounds.Top
        x = PageLeft + (PageWidth - e.Graphics.MeasureString(Text, MyFontTitle).Width) / 2
        e.Graphics.DrawString(Text, MyFontTitle, Brushes.Black, x, y)
        y += 2 * MyFontTitle.GetHeight(e.Graphics)

        x = PageLeft
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(1.3), CType(5.5 * MyFont.GetHeight(e.Graphics), Single))
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(8.0), CType(5.5 * MyFont.GetHeight(e.Graphics), Single))
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(13.0), CType(5.5 * MyFont.GetHeight(e.Graphics), Single))
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(14.0), CType(5.5 * MyFont.GetHeight(e.Graphics), Single))
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(15.0), CType(5.5 * MyFont.GetHeight(e.Graphics), Single))
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(28.0), CType(5.5 * MyFont.GetHeight(e.Graphics), Single))

        e.Graphics.DrawString("NrKarty", MyFont, Brushes.Black, x + cm2pts(0.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Nazwa", MyFont, Brushes.Black, x + cm2pts(1.4), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Sprzeda¿-wartoœæ", MyFont, Brushes.Black, x + cm2pts(8.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Skup", MyFont, Brushes.Black, x + cm2pts(13.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Sprz", MyFont, Brushes.Black, x + cm2pts(14.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Kontakty Handlowe", MyFont, Brushes.Black, x + cm2pts(15.1), y + CType(0.1 * (100 / 2.54), Single))
        y += MyFont.GetHeight(e.Graphics)

        e.Graphics.DrawString("NrTofi", MyFont, Brushes.Black, x + cm2pts(0.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Osoba kontaktowa", MyFont, Brushes.Black, x + cm2pts(1.4), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Sprzeda¿-Plan", MyFont, Brushes.Black, x + cm2pts(8.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Ost.", MyFont, Brushes.Black, x + cm2pts(13.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Ost..", MyFont, Brushes.Black, x + cm2pts(14.1), y + CType(0.1 * (100 / 2.54), Single))
        y += MyFont.GetHeight(e.Graphics)

        e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(0.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Adres", MyFont, Brushes.Black, x + cm2pts(1.4), y + CType(0.1 * (100 / 2.54), Single))
        RightTxt("Rejon", x + cm2pts(1.4), y + CType(0.1 * (100 / 2.54), Single), cm2pts(6.6), MyFont, e)
        e.Graphics.DrawString("Sprzeda¿-Ost.transakcje", MyFont, Brushes.Black, x + cm2pts(8.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Nast.", MyFont, Brushes.Black, x + cm2pts(13.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Nast.", MyFont, Brushes.Black, x + cm2pts(14.1), y + CType(0.1 * (100 / 2.54), Single))
        y += MyFont.GetHeight(e.Graphics)

        e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(0.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(1.4), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Dzia³ania akcyjne", MyFont, Brushes.Black, x + cm2pts(8.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(13.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(14.1), y + CType(0.1 * (100 / 2.54), Single))
        y += MyFont.GetHeight(e.Graphics)

        e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(0.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(1.4), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("Czyszczenie", MyFont, Brushes.Black, x + cm2pts(8.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(13.1), y + CType(0.1 * (100 / 2.54), Single))
        e.Graphics.DrawString("", MyFont, Brushes.Black, x + cm2pts(14.1), y + CType(0.1 * (100 / 2.54), Single))
        y += 2 * MyFont.GetHeight(e.Graphics)

        ' petla do konca strony
        If Doc.NextRowNumber < DataViewKlienci.Count Then
            Do
                'ustaw kursor na wierszu DataGrid
                LiRow = Doc.NextRowNumber

                dagKlienci.CurrentCell = dagKlienci(0, LiRow)

                PobierzKlienci()
                '------------------------------------------
                Uwagi1 = ""
                Uwagi2 = ""
                Uwagi3 = ""
                Uwagi4 = ""
                Uwagi5 = ""
                Uwagi6 = ""
                'Uwagi = oSprzedUwagiKli
                Uwagi = PobierzKontaktyDoTXT(oIdentKli, MyFont, e)

                ii = InStr(Uwagi, vbLf)
                If ii > 0 Then
                    Uwagi1 = Trim(Mid(Uwagi, 1, ii - 1).Replace("|"c, " "c))
                    Uwagi = Mid(Uwagi, ii + 1)
                    ii = InStr(Uwagi, vbLf)
                    If ii > 0 Then
                        Uwagi2 = Trim(Mid(Uwagi, 1, ii - 1).Replace("|"c, " "c))
                        Uwagi = Mid(Uwagi, ii + 1)
                        ii = InStr(Uwagi, vbLf)
                        If ii > 0 Then
                            Uwagi3 = Trim(Mid(Uwagi, 1, ii - 1).Replace("|"c, " "c))
                            Uwagi = Mid(Uwagi, ii + 1)
                            ii = InStr(Uwagi, vbLf)
                            If ii > 0 Then
                                Uwagi4 = Trim(Mid(Uwagi, 1, ii - 1).Replace("|"c, " "c))
                                Uwagi = Mid(Uwagi, ii + 1)
                                ii = InStr(Uwagi, vbLf)
                                If ii > 0 Then
                                    Uwagi5 = Trim(Mid(Uwagi, 1, ii - 1).Replace("|"c, " "c))
                                    Uwagi = Mid(Uwagi, ii + 1)
                                    ii = InStr(Uwagi, vbLf)
                                    If ii > 0 Then
                                        Uwagi6 = Trim(Mid(Uwagi, 1, ii - 1).Replace("|"c, " "c))
                                        Uwagi = Mid(Uwagi, ii + 1)
                                    Else
                                        Uwagi6 = Trim(Uwagi)
                                    End If
                                Else
                                    Uwagi5 = Uwagi
                                End If
                            Else
                                Uwagi4 = Uwagi
                            End If
                        Else
                            Uwagi3 = Uwagi
                        End If
                    Else
                        Uwagi2 = Uwagi
                    End If
                Else
                    Uwagi1 = Uwagi
                End If





                'sprawdz czy sie zmiesci
                UwagiX = ""
                Do While e.Graphics.MeasureString(Uwagi1, MyFont).Width > CType(13 * (100 / 2.54), Single)
                    UwagiX = Mid(Uwagi1, Len(Uwagi1), 1) + UwagiX
                    Uwagi1 = Mid(Uwagi1, 1, Len(Uwagi1) - 1)
                Loop
                If Len(UwagiX) > 0 Then
                    Uwagi6 = Uwagi5
                    Uwagi5 = Uwagi4
                    Uwagi4 = Uwagi3
                    Uwagi3 = Uwagi2
                    Uwagi2 = UwagiX
                    UwagiX = ""
                End If

                Do While e.Graphics.MeasureString(Uwagi2, MyFont).Width > CType(13 * (100 / 2.54), Single)
                    UwagiX = Mid(Uwagi2, Len(Uwagi2), 1) + UwagiX
                    Uwagi2 = Mid(Uwagi2, 1, Len(Uwagi2) - 1)
                Loop
                If Len(UwagiX) > 0 Then
                    Uwagi6 = Uwagi5
                    Uwagi5 = Uwagi4
                    Uwagi4 = Uwagi3
                    Uwagi3 = UwagiX
                    UwagiX = ""
                End If

                Do While e.Graphics.MeasureString(Uwagi3, MyFont).Width > CType(13 * (100 / 2.54), Single)
                    UwagiX = Mid(Uwagi3, Len(Uwagi3), 1) + UwagiX
                    Uwagi3 = Mid(Uwagi3, 1, Len(Uwagi3) - 1)
                Loop
                If Len(UwagiX) > 0 Then
                    Uwagi6 = Uwagi5
                    Uwagi5 = Uwagi4
                    Uwagi4 = UwagiX
                    UwagiX = ""
                End If

                Do While e.Graphics.MeasureString(Uwagi4, MyFont).Width > CType(13 * (100 / 2.54), Single)
                    UwagiX = Mid(Uwagi4, Len(Uwagi4), 1) + UwagiX
                    Uwagi4 = Mid(Uwagi4, 1, Len(Uwagi4) - 1)
                Loop
                If Len(UwagiX) > 0 Then
                    Uwagi6 = Uwagi5
                    Uwagi5 = UwagiX
                    UwagiX = ""
                End If

                Do While e.Graphics.MeasureString(Uwagi5, MyFont).Width > CType(13 * (100 / 2.54), Single)
                    UwagiX = Mid(Uwagi5, Len(Uwagi5), 1) + UwagiX
                    Uwagi5 = Mid(Uwagi5, 1, Len(Uwagi5) - 1)
                Loop
                If Len(UwagiX) > 0 Then
                    Uwagi6 = Uwagi5
                    UwagiX = ""
                End If

                Do While e.Graphics.MeasureString(Uwagi6, MyFont).Width > CType(13 * (100 / 2.54), Single)
                    UwagiX = Mid(Uwagi6, Len(Uwagi6), 1) + UwagiX
                    Uwagi6 = Mid(Uwagi6, 1, Len(Uwagi6) - 1)
                Loop
                If Len(UwagiX) > 0 Then
                    UwagiX = ""
                End If







                Tofi1 = ""
                Tofi2 = ""
                Tofi3 = ""
                Tofi4 = ""
                Tofi5 = ""
                Tofi6 = ""
                words = oNrTOFITxtKli.Split(splitChars, StringSplitOptions.None)
                If words.Length = 0 Then
                ElseIf words.Length = 1 Then
                    Tofi1 = oNrTOFITxtKli
                ElseIf words.Length = 2 Then
                    Tofi1 = words(0)
                    Tofi2 = words(1)
                ElseIf words.Length = 3 Then
                    Tofi1 = words(0)
                    Tofi2 = words(1)
                    Tofi3 = words(2)
                ElseIf words.Length = 4 Then
                    Tofi1 = words(0)
                    Tofi2 = words(1)
                    Tofi3 = words(2)
                    Tofi4 = words(3)
                ElseIf words.Length = 5 Then
                    Tofi1 = words(0)
                    Tofi2 = words(1)
                    Tofi3 = words(2)
                    Tofi4 = words(3)
                    Tofi5 = words(4)
                Else
                    Tofi1 = words(0)
                    Tofi2 = words(1)
                    Tofi3 = words(2)
                    Tofi4 = words(3)
                    Tofi5 = words(4)
                    Tofi6 = words(5)
                End If





                x = PageLeft 'od marginesu...
                Doc.LineNumber += 1
                Doc.NextRowNumber += 1

                'wiersz 1
                LeftTxt(oIdentKli, x + cm2pts(0.0), y, cm2pts(1.3), MyFont, e)
                LeftTxt(oNazwa1Kli, x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
                LeftTxt("", x + cm2pts(8.0), y, cm2pts(5.0), MyFont, e)

                LeftTxt(oSprzedOKontaktTKli & "/" & Mid(oSprzedOKontaktDKli, 1, 2), x + cm2pts(14.0), y, cm2pts(1.0), MyFont, e)
                LeftTxt(Uwagi1, x + cm2pts(15.0), y, cm2pts(13.0), MyFont, e)
                y += LineHeight

                'wiersz 2
                LeftTxt(Tofi1, x + cm2pts(0.0), y, cm2pts(1.3), MyFont, e)
                LeftTxt(Trim(OsobaKontaktowa(oIdentKli)), x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
                LeftTxt("", x + cm2pts(8.0), y, cm2pts(5.0), MyFont, e)

                LeftTxt(oSprzedNKontaktTKli & "/" & Mid(oSprzedNKontaktDKli, 1, 2), x + cm2pts(14.0), y, cm2pts(1.0), MyFont, e)
                LeftTxt(Uwagi2, x + cm2pts(15.0), y, cm2pts(13.0), MyFont, e)
                y += LineHeight

                'wiersz 3
                LeftTxt(Tofi2, x + cm2pts(0.0), y, cm2pts(1.3), MyFont, e)
                LeftTxt(Trim(oKodPoczKli) + " " + Trim(oMiejscKli), x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
                RightTxt(Trim(oRejonKli), x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
                LeftTxt(OstTransakcje(oIdentKli, True), x + cm2pts(8.0), y, cm2pts(5.0), MyFont, e)
                LeftTxt("", x + cm2pts(13.0), y, cm2pts(1.0), MyFont, e)
                LeftTxt("", x + cm2pts(14.0), y, cm2pts(1.0), MyFont, e)
                LeftTxt(Uwagi3, x + cm2pts(15.0), y, cm2pts(13.0), MyFont, e)
                y += LineHeight

                'wiersz 4
                If Len(Trim(Tofi3)) > 0 Or
                    (Len(Trim(oUlicaKli)) > 0 Or Len(Trim(Str(oNumNrDomuKli))) > 0 Or Len(Trim(oIDDomuKli)) > 0) Or
                    Len(Trim("")) > 0 Or
                    Len(Trim(Uwagi4)) > 0 Then

                    LeftTxt(Tofi3, x + cm2pts(0.0), y, cm2pts(1.3), MyFont, e)
                    If Len(oIDDomuKli) > 0 Then
                        LeftTxt(Trim(oUlicaKli) & " " & Trim(Str(oNumNrDomuKli)) & " " & Trim(oIDDomuKli), x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
                    Else
                        LeftTxt(Trim(oUlicaKli) & " " & Trim(Str(oNumNrDomuKli)), x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
                    End If
                    LeftTxt("", x + cm2pts(8.0), y, cm2pts(5.0), MyFont, e)
                    LeftTxt("", x + cm2pts(13.0), y, cm2pts(1.0), MyFont, e)
                    LeftTxt("", x + cm2pts(14.0), y, cm2pts(1.0), MyFont, e)
                    LeftTxt(Uwagi4, x + cm2pts(15.0), y, cm2pts(13.0), MyFont, e)
                    y += LineHeight
                Else
                    'jesli jest wiersz 5 - drukuj pusty
                    If Len(Trim(Tofi4)) > 0 Or
                        Len(Trim("")) > 0 Or
                        Len(Trim(Uwagi5)) > 0 Then
                        ' pusta linia
                        y += LineHeight
                    End If
                End If

                'wiersz 5
                If Len(Trim(Tofi4)) > 0 Or
                    Len(Trim("")) > 0 Or
                    Len(Trim(Uwagi5)) > 0 Then

                    LeftTxt(Tofi4, x + cm2pts(0.0), y, cm2pts(1.3), MyFont, e)
                    LeftTxt("", x + cm2pts(1.4), y, cm2pts(6.6), MyFont, e)
                    LeftTxt("", x + cm2pts(8.0), y, cm2pts(5.0), MyFont, e)
                    LeftTxt("", x + cm2pts(13.0), y, cm2pts(1.0), MyFont, e)
                    LeftTxt("", x + cm2pts(14.0), y, cm2pts(1.0), MyFont, e)
                    LeftTxt(Uwagi5, x + cm2pts(15.0), y, cm2pts(13.0), MyFont, e)
                    y += LineHeight
                End If

                e.Graphics.DrawLine(DrawingPen, x, y + CType(0.2 * LineHeight, Single), x + CType(28.0 * (100 / 2.54), Single), y + CType(0.5 * LineHeight, Single))
                y += CType(0.5 * LineHeight, Single)
            Loop Until ((y + LineHeight) > e.MarginBounds.Bottom) Or (Doc.NextRowNumber >= DataViewKlienci.Count)
        End If

        If (Doc.NextRowNumber < DataViewKlienci.Count) Then
            ' There is still at least one more page.
            ' Signal this event to fire again.
            e.HasMorePages = True
        Else
            ' Printing is complete.
            y += LineHeight
            x = PageLeft 'od marginesu...
            Text = "Koniec zestawienia na pozycji " & Str(Doc.LineNumber)
            e.Graphics.DrawString(Text, MyFont, Brushes.Black, x + CType(0.2 * (100 / 2.54), Single), y)

            Doc.PageNumber = 0
            Doc.LineNumber = 0
            Doc.NextRowNumber = 0

            e.HasMorePages = False
        End If

    End Sub





    Private Sub MyDoc_Zestawienie(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        ' Retrieve the document that sent this event.
        ' You could store the document in a class member variable,
        ' but this approach allows you to use the same event handler
        ' to handle multiple print documents at once.
        Dim Doc As Softart.myPrintDocument = CType(sender, PrintDocument)
        Dim Text As String
        Dim LiRow As Integer

        Dim Uwagi As String = ""
        Dim Uwagi1 As String = ""
        Dim Uwagi2 As String = ""
        Dim Uwagi3 As String = ""
        Dim Uwagi4 As String = ""
        Dim Uwagi5 As String = ""
        Dim Uwagi6 As String = ""
        Dim UwagiX As String = ""
        Dim ii As Integer = 0

        e.PageSettings.Landscape = True

        ' Define the font and determine the line height.
        Dim FontSizeTitle As Integer = 12
        Dim MyFontTitle As New System.Drawing.Font("Arial", FontSizeTitle, FontStyle.Bold)
        Dim LineHeightTitle As Single = MyFontTitle.GetHeight(e.Graphics)

        Dim FontSize As Integer = 8
        Dim MyFont As New System.Drawing.Font("Arial", FontSize, FontStyle.Regular)
        Dim LineHeight As Single = MyFont.GetHeight(e.Graphics)

        Dim DrawingPen As New Pen(Color.Black, 1)

        ' Create variables to hold position on page.
        Dim x As Single = e.MarginBounds.Left
        Dim y As Single = e.MarginBounds.Top

        Dim PageWidth As Single = cm2pts(28)        'e.MarginBounds.Width
        Dim PageHeight As Single = cm2pts(19)       'e.MarginBounds.Height
        Dim PageLeft As Single = cm2pts(0.5)          'e.MarginBounds.Left
        Dim PageTop As Single = cm2pts(1)           'e.MarginBounds.Top

        Dim words() As String
        Dim splitChars() As Char = {","c}
        Dim Tofi1 As String = ""
        Dim Tofi2 As String = ""
        Dim Tofi3 As String = ""
        Dim Tofi4 As String = ""
        Dim Tofi5 As String = ""
        Dim Tofi6 As String = ""

        'tekst z katalogu KONTAKTY HANDLOWE
        ' KLIENT | DATA | OPERATOR | RODZAJ KONTaktu | UWAGI
        Dim khSepar As Integer = 0
        Dim khText As String = ""

        Dim khKlient As String = ""
        Dim khData As String = ""
        Dim khOperator As String = ""
        Dim khRodzajK As String = ""
        Dim khUwagi As String = ""


        ' Naglowek strony
        Doc.PageNumber += 1
        x = PageLeft
        y = PageTop

        RightTxt("Str. " & Trim(Str(Doc.PageNumber)), x + cm2pts(0), y, PageWidth, MyFont, e)
        y += MyFont.GetHeight(e.Graphics)

        Text = "Zestawienie klientów"
        x = PageLeft + (PageWidth - e.Graphics.MeasureString(Text, MyFontTitle).Width) / 2
        e.Graphics.DrawString(Text, MyFontTitle, Brushes.Black, x, y)
        y += 2 * MyFontTitle.GetHeight(e.Graphics)



        x = PageLeft
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(1.5), 4 * LineHeight)
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(7.5), 4 * LineHeight)
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(13.5), 4 * LineHeight)
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(19.5), 4 * LineHeight)
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(25.5), 4 * LineHeight)
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(27.0), 4 * LineHeight)
        e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(28.5), 4 * LineHeight)

        e.Graphics.DrawRectangle(DrawingPen, x, y + 4 * LineHeight, cm2pts(1.5), 2 * LineHeight)
        e.Graphics.DrawRectangle(DrawingPen, x, y + 4 * LineHeight, cm2pts(14.0), 2 * LineHeight)
        e.Graphics.DrawRectangle(DrawingPen, x, y + 4 * LineHeight, cm2pts(15.5), 2 * LineHeight)
        e.Graphics.DrawRectangle(DrawingPen, x, y + 4 * LineHeight, cm2pts(28.5), 2 * LineHeight)
        y += 0.5 * MyFont.GetHeight(e.Graphics)

        CentrTxt("Nr Karty", x + cm2pts(0), y, cm2pts(1.5), MyFont, e)
        CentrTxt("Nazwa", x + cm2pts(1.5), y, cm2pts(6.0), MyFont, e)
        CentrTxt("Miejscowosc", x + cm2pts(7.5), y, cm2pts(6.0), MyFont, e)
        CentrTxt("Karty PAYBACK", x + cm2pts(13.5), y, cm2pts(6.0), MyFont, e)
        CentrTxt("Sprzeda¿-wartoœæ", x + cm2pts(19.5), y, cm2pts(6), MyFont, e)
        CentrTxt("Skup", x + cm2pts(25.5), y, cm2pts(1.5), MyFont, e)
        CentrTxt("Sprzed", x + cm2pts(27.0), y, cm2pts(1.5), MyFont, e)
        y += 1.0 * MyFont.GetHeight(e.Graphics)

        CentrTxt("Nr Tofi", x + cm2pts(0), y, cm2pts(1.5), MyFont, e)
        CentrTxt("Osoba kontaktowa", x + cm2pts(1.5), y, cm2pts(6.0), MyFont, e)
        CentrTxt("Ulica", x + cm2pts(7.5), y, cm2pts(6.0), MyFont, e)
        CentrTxt("Dzia³ania akcyjne", x + cm2pts(13.5), y, cm2pts(6.0), MyFont, e)
        CentrTxt("Sprzeda¿-plan", x + cm2pts(19.5), y, cm2pts(6.0), MyFont, e)
        CentrTxt("Ost.", x + cm2pts(25.5), y, cm2pts(1.5), MyFont, e)
        CentrTxt("Ost.", x + cm2pts(27.0), y, cm2pts(1.5), MyFont, e)
        y += 1.0 * MyFont.GetHeight(e.Graphics)

        CentrTxt("", x + cm2pts(0), y, cm2pts(1.5), MyFont, e)
        CentrTxt("", x + cm2pts(1.5), y, cm2pts(6.0), MyFont, e)
        CentrTxt("", x + cm2pts(7.5), y, cm2pts(6.0), MyFont, e)
        CentrTxt("Czyszczenie", x + cm2pts(13.5), y, cm2pts(6.0), MyFont, e)
        CentrTxt("Sprzeda¿-Ost.transakcje", x + cm2pts(19.5), y, cm2pts(6.0), MyFont, e)
        CentrTxt("Nast.", x + cm2pts(25.5), y, cm2pts(1.5), MyFont, e)
        CentrTxt("Nast.", x + cm2pts(27.0), y, cm2pts(1.5), MyFont, e)
        y += 2 * MyFont.GetHeight(e.Graphics)

        'CentrTxt("Uwagi", x + cm2pts(1.5), y, cm2pts(12.5), MyFont, e)
        CentrTxt("Iloœæ sprzêtu", x + cm2pts(1.5), y, cm2pts(12.5), MyFont, e)
        CentrTxt("Kontakty Handlowe", x + cm2pts(15.5), y, cm2pts(12.0), MyFont, e)
        y += 2 * MyFont.GetHeight(e.Graphics)





        ' petla do konca strony
        If Doc.NextRowNumber < DataViewKlienci.Count Then
            Do
                '----------------------------------------
                'kontynuujemy drukujemy dwie kolumny - z lewej UWAGI
                Do While IleLiniiTrzebaWydrukowac > 0
                    If DrukujLinieLNr < IleLiniiL Then
                        DrukujLinieLNr += 1
                        LeftTxt(DajLinieTextuNr(TextL, DrukujLinieLNr, cm2pts(12.5), MyFont, e), x + cm2pts(1.5), y, cm2pts(12.5), MyFont, e)
                    End If
                    If DrukujLiniePNr < IleLiniiP Then
                        DrukujLiniePNr += 1
                        khText = DajLinieTextuNr(TextP, DrukujLiniePNr, cm2pts(12.5), MyFont, e)
                        '------------------------------------
                        'analizuj linie tekstu
                        ' KLIENT | DATA | OPERATOR | RODZAJ KONTaktu | UWAGI
                        khKlient = ""
                        khData = ""
                        khOperator = ""
                        khRodzajK = ""
                        khUwagi = ""

                        'oddziel pola separator=|
                        khSepar = InStr(khText, "|")
                        If khSepar > 0 Then
                            khKlient = Mid(khText, 1, khSepar - 1)
                            khText = Mid(khText, khSepar + 1)

                            khSepar = InStr(khText, "|")
                            If khSepar > 0 Then
                                khData = Mid(khText, 1, khSepar - 1)
                                khText = Mid(khText, khSepar + 1)

                                khSepar = InStr(khText, "|")
                                If khSepar > 0 Then
                                    khOperator = Mid(khText, 1, khSepar - 1)
                                    khText = Mid(khText, khSepar + 1)

                                    khSepar = InStr(khText, "|")
                                    If khSepar > 0 Then
                                        khRodzajK = Mid(khText, 1, khSepar - 1)
                                        khUwagi = Mid(khText, khSepar + 1)
                                    End If
                                End If
                            End If
                        End If

                        LeftTxt(khData, x + cm2pts(15.5), y, cm2pts(2), MyFont, e)
                        LeftTxt(khOperator, x + cm2pts(17.5), y, cm2pts(2), MyFont, e)
                        LeftTxt(khRodzajK, x + cm2pts(19.5), y, cm2pts(2), MyFont, e)
                        LeftTxt(khUwagi, x + cm2pts(21.5), y, cm2pts(6.5), MyFont, e)
                        '------------------------------------
                    End If
                    IleLiniiTrzebaWydrukowac -= 1
                    y += LineHeight

                    If ((y + 5 * LineHeight) > e.MarginBounds.Bottom) Then
                        Exit Do
                    End If
                Loop


                'sprawdz czy mozna drukowaæ dalej klientów...
                If IleLiniiTrzebaWydrukowac > 0 Then
                    'jeszcze nie skonczylismy - nastepna strona
                Else
                    'koniec poprzedniego klienta
                    If (Doc.NextRowNumber < DataViewKlienci.Count) Then
                        'mozna pobrac nastepnego klienta
                        If Doc.NextRowNumber > 0 Then
                            y += CType(0.5 * LineHeight, Single)
                            e.Graphics.DrawLine(DrawingPen, x, y + CType(0.2 * LineHeight, Single), x + CType(28.5 * (100 / 2.54), Single), y + CType(0.5 * LineHeight, Single))
                            y += CType(0.5 * LineHeight, Single)
                        End If


                        'ustaw kursor na wierszu DataGrid
                        LiRow = Doc.NextRowNumber
                        dagKlienci.CurrentCell = dagKlienci(0, LiRow)
                        PobierzKlienci()
                        '------------------------------------------





                        'Uwagi1 = ""
                        'Uwagi2 = ""
                        'Uwagi3 = ""
                        'Uwagi4 = ""
                        'Uwagi5 = ""
                        'Uwagi6 = ""
                        'Uwagi = oSprzedUwagiKli

                        'ii = InStr(Uwagi, vbLf)
                        'If ii > 0 Then
                        '    Uwagi1 = Mid(Uwagi, 1, ii - 1)
                        '    Uwagi = Mid(Uwagi, ii + 1)
                        '    ii = InStr(Uwagi, vbLf)
                        '    If ii > 0 Then
                        '        Uwagi2 = Mid(Uwagi, 1, ii - 1)
                        '        Uwagi = Mid(Uwagi, ii + 1)
                        '        ii = InStr(Uwagi, vbLf)
                        '        If ii > 0 Then
                        '            Uwagi3 = Mid(Uwagi, 1, ii - 1)
                        '            Uwagi = Mid(Uwagi, ii + 1)
                        '            ii = InStr(Uwagi, vbLf)
                        '            If ii > 0 Then
                        '                Uwagi4 = Mid(Uwagi, 1, ii - 1)
                        '                Uwagi = Mid(Uwagi, ii + 1)
                        '                ii = InStr(Uwagi, vbLf)
                        '                If ii > 0 Then
                        '                    Uwagi5 = Mid(Uwagi, 1, ii - 1)
                        '                    Uwagi = Mid(Uwagi, ii + 1)
                        '                    ii = InStr(Uwagi, vbLf)
                        '                    If ii > 0 Then
                        '                        Uwagi6 = Mid(Uwagi, 1, ii - 1)
                        '                        Uwagi = Mid(Uwagi, ii + 1)
                        '                    Else
                        '                        Uwagi6 = Uwagi
                        '                    End If
                        '                Else
                        '                    Uwagi5 = Uwagi
                        '                End If
                        '            Else
                        '                Uwagi4 = Uwagi
                        '            End If
                        '        Else
                        '            Uwagi3 = Uwagi
                        '        End If
                        '    Else
                        '        Uwagi2 = Uwagi
                        '    End If
                        'Else
                        '    Uwagi1 = Uwagi
                        'End If





                        ''sprawdz czy sie zmiesci
                        'UwagiX = ""
                        'Do While e.Graphics.MeasureString(Uwagi1, MyFont).Width > CType(10 * (100 / 2.54), Single)
                        '    UwagiX = Mid(Uwagi1, Len(Uwagi1), 1) + UwagiX
                        '    Uwagi1 = Mid(Uwagi1, 1, Len(Uwagi1) - 1)
                        'Loop
                        'If Len(UwagiX) > 0 Then
                        '    Uwagi6 = Uwagi5
                        '    Uwagi5 = Uwagi4
                        '    Uwagi4 = Uwagi3
                        '    Uwagi3 = Uwagi2
                        '    Uwagi2 = UwagiX
                        '    UwagiX = ""
                        'End If

                        'Do While e.Graphics.MeasureString(Uwagi2, MyFont).Width > CType(10 * (100 / 2.54), Single)
                        '    UwagiX = Mid(Uwagi2, Len(Uwagi2), 1) + UwagiX
                        '    Uwagi2 = Mid(Uwagi2, 1, Len(Uwagi2) - 1)
                        'Loop
                        'If Len(UwagiX) > 0 Then
                        '    Uwagi6 = Uwagi5
                        '    Uwagi5 = Uwagi4
                        '    Uwagi4 = Uwagi3
                        '    Uwagi3 = UwagiX
                        '    UwagiX = ""
                        'End If

                        'Do While e.Graphics.MeasureString(Uwagi3, MyFont).Width > CType(10 * (100 / 2.54), Single)
                        '    UwagiX = Mid(Uwagi3, Len(Uwagi3), 1) + UwagiX
                        '    Uwagi3 = Mid(Uwagi3, 1, Len(Uwagi3) - 1)
                        'Loop
                        'If Len(UwagiX) > 0 Then
                        '    Uwagi6 = Uwagi5
                        '    Uwagi5 = Uwagi4
                        '    Uwagi4 = UwagiX
                        '    UwagiX = ""
                        'End If

                        'Do While e.Graphics.MeasureString(Uwagi4, MyFont).Width > CType(10 * (100 / 2.54), Single)
                        '    UwagiX = Mid(Uwagi4, Len(Uwagi4), 1) + UwagiX
                        '    Uwagi4 = Mid(Uwagi4, 1, Len(Uwagi4) - 1)
                        'Loop
                        'If Len(UwagiX) > 0 Then
                        '    Uwagi6 = Uwagi5
                        '    Uwagi5 = UwagiX
                        '    UwagiX = ""
                        'End If

                        'Do While e.Graphics.MeasureString(Uwagi5, MyFont).Width > CType(10 * (100 / 2.54), Single)
                        '    UwagiX = Mid(Uwagi5, Len(Uwagi5), 1) + UwagiX
                        '    Uwagi5 = Mid(Uwagi5, 1, Len(Uwagi5) - 1)
                        'Loop
                        'If Len(UwagiX) > 0 Then
                        '    Uwagi6 = Uwagi5
                        '    UwagiX = ""
                        'End If

                        'Do While e.Graphics.MeasureString(Uwagi6, MyFont).Width > CType(10 * (100 / 2.54), Single)
                        '    UwagiX = Mid(Uwagi6, Len(Uwagi6), 1) + UwagiX
                        '    Uwagi6 = Mid(Uwagi6, 1, Len(Uwagi6) - 1)
                        'Loop
                        'If Len(UwagiX) > 0 Then
                        '    UwagiX = ""
                        'End If











                        Tofi1 = ""
                        Tofi2 = ""
                        Tofi3 = ""
                        Tofi4 = ""
                        Tofi5 = ""
                        Tofi6 = ""
                        words = oNrTOFITxtKli.Split(splitChars, StringSplitOptions.None)
                        If words.Length = 0 Then
                        ElseIf words.Length = 1 Then
                            Tofi1 = oNrTOFITxtKli
                        ElseIf words.Length = 2 Then
                            Tofi1 = words(0)
                            Tofi2 = words(1)
                        ElseIf words.Length = 3 Then
                            Tofi1 = words(0)
                            Tofi2 = words(1)
                            Tofi3 = words(2)
                        ElseIf words.Length = 4 Then
                            Tofi1 = words(0)
                            Tofi2 = words(1)
                            Tofi3 = words(2)
                            Tofi4 = words(3)
                        ElseIf words.Length = 5 Then
                            Tofi1 = words(0)
                            Tofi2 = words(1)
                            Tofi3 = words(2)
                            Tofi4 = words(3)
                            Tofi5 = words(4)
                        Else
                            Tofi1 = words(0)
                            Tofi2 = words(1)
                            Tofi3 = words(2)
                            Tofi4 = words(3)
                            Tofi5 = words(4)
                            Tofi6 = words(5)
                        End If









                        x = PageLeft 'od marginesu...
                        Doc.LineNumber += 1
                        Doc.NextRowNumber += 1






                        'CentrTxt("Nr Karty", x + cm2pts(0), y, cm2pts(1.5), MyFont, e)
                        'CentrTxt("Nazwa", x + cm2pts(1.5), y, cm2pts(6.0), MyFont, e)
                        'CentrTxt("Miejscowosc", x + cm2pts(7.5), y, cm2pts(6.0), MyFont, e)
                        'CentrTxt("Karty PAYBACK", x + cm2pts(13.5), y, cm2pts(6.0), MyFont, e)
                        'CentrTxt("Sprzeda¿-wartoœæ", x + cm2pts(19.5), y, cm2pts(6), MyFont, e)
                        'CentrTxt("Skup", x + cm2pts(25.5), y, cm2pts(1.5), MyFont, e)
                        'CentrTxt("Sprzed", x + cm2pts(27.0), y, cm2pts(1.5), MyFont, e)

                        'wiersz 1
                        LeftTxt(oIdentKli, x + cm2pts(0.0), y, cm2pts(1.5), MyFont, e)
                        LeftTxt(oNazwa1Kli, x + cm2pts(1.5), y, cm2pts(6.0), MyFont, e)
                        LeftTxt(Trim(oKodPoczKli) + " " + Trim(oMiejscKli), x + cm2pts(7.5), y, cm2pts(6.0), MyFont, e)
                        LeftTxt(oNryKartPaybackKli, x + cm2pts(13.5), y, cm2pts(6.0), MyFont, e)
                        LeftTxt("", x + cm2pts(19.5), y, cm2pts(6.0), MyFont, e)

                        LeftTxt(oSprzedOKontaktTKli & "/" & Mid(oSprzedOKontaktDKli, 1, 2), x + cm2pts(27.0), y, cm2pts(1.5), MyFont, e)
                        y += LineHeight








                        'y += 1.0 * MyFont.GetHeight(e.Graphics)
                        'CentrTxt("Nr Tofi", x + cm2pts(0), y, cm2pts(1.5), MyFont, e)
                        'CentrTxt("Osoba kontaktowa", x + cm2pts(1.5), y, cm2pts(6.0), MyFont, e)
                        'CentrTxt("Ulica", x + cm2pts(7.5), y, cm2pts(6.0), MyFont, e)
                        'CentrTxt("Dzia³ania akcyjne", x + cm2pts(13.5), y, cm2pts(6.0), MyFont, e)
                        'CentrTxt("Sprzeda¿-plan", x + cm2pts(19.5), y, cm2pts(6.0), MyFont, e)
                        'CentrTxt("Ost.", x + cm2pts(25.5), y, cm2pts(1.5), MyFont, e)
                        'CentrTxt("Ost.", x + cm2pts(27.0), y, cm2pts(1.5), MyFont, e)

                        'wiersz 2
                        LeftTxt(Tofi1, x + cm2pts(0.0), y, cm2pts(1.5), MyFont, e)
                        LeftTxt(Trim(OsobaKontaktowa(oIdentKli)), x + cm2pts(1.5), y, cm2pts(6.0), MyFont, e)
                        If Len(oIDDomuKli) > 0 Then
                            LeftTxt(Trim(oUlicaKli) & " " & Trim(Str(oNumNrDomuKli)) & " " & Trim(oIDDomuKli), x + cm2pts(7.5), y, cm2pts(6.0), MyFont, e)
                        Else
                            LeftTxt(Trim(oUlicaKli) & " " & Trim(Str(oNumNrDomuKli)), x + cm2pts(7.5), y, cm2pts(6.0), MyFont, e)
                        End If
                        LeftTxt("", x + cm2pts(13.5), y, cm2pts(6.0), MyFont, e)

                        IdentCoDalej(oIdentKli)
                        LeftTxt(DajLinieTextuNr(cdOPIS, 1, cm2pts(6.0), MyFont, e), x + cm2pts(19.5), y, cm2pts(6.0), MyFont, e)

                        LeftTxt(oSprzedNKontaktTKli & "/" & Mid(oSprzedNKontaktDKli, 1, 2), x + cm2pts(27.0), y, cm2pts(1.5), MyFont, e)
                        y += LineHeight






                        'y += 1.0 * MyFont.GetHeight(e.Graphics)
                        'CentrTxt("", x + cm2pts(0), y, cm2pts(1.5), MyFont, e)
                        'CentrTxt("", x + cm2pts(1.5), y, cm2pts(6.0), MyFont, e)
                        'CentrTxt("", x + cm2pts(7.5), y, cm2pts(6.0), MyFont, e)
                        'CentrTxt("Czyszczenie", x + cm2pts(13.5), y, cm2pts(6.0), MyFont, e)
                        'CentrTxt("Sprzeda¿-Ost.transakcje", x + cm2pts(19.5), y, cm2pts(6.0), MyFont, e)
                        'CentrTxt("Nast.", x + cm2pts(25.5), y, cm2pts(1.5), MyFont, e)
                        'CentrTxt("Nast.", x + cm2pts(27.0), y, cm2pts(1.5), MyFont, e)

                        'wiersz 3
                        LeftTxt(Tofi3, x + cm2pts(0.0), y, cm2pts(1.5), MyFont, e)
                        LeftTxt("", x + cm2pts(1.5), y, cm2pts(6.0), MyFont, e)
                        LeftTxt(Trim(oRejonKli), x + cm2pts(7.5), y, cm2pts(6.0), MyFont, e)
                        LeftTxt("", x + cm2pts(13.5), y, cm2pts(6.0), MyFont, e)
                        LeftTxt(OstTransakcje(oIdentKli, True), x + cm2pts(19.5), y, cm2pts(9.0), MyFont, e)
                        y += LineHeight







                        '----------------------------------------
                        ''drukujemy dwie kolumny - z lewej UWAGI
                        'TextL = oSprzedUwagiKli
                        'i100 = 0
                        'TxtLDl = e.Graphics.MeasureString(TextL, MyFont).Width
                        'If TxtLDl > 6 * cm2pts(12.5) Then
                        '    'tekst dluzszy niz mozna wydrukowac...
                        '    Text = TextL
                        '    i100 = 0
                        '    TextL = Mid(Text, i100 * 100 + 1, 100)
                        '    Do While e.Graphics.MeasureString(TextL, MyFont).Width < 6 * cm2pts(12.5)
                        '        i100 += 1
                        '        TextL &= Mid(Text, i100 * 100 + 1, 100)
                        '    Loop
                        'End If
                        'IleLiniiL = IleLiniiPotrzebaNaText(TextL, cm2pts(12.5), MyFont, e)


                        'txtKatOpisPRYZMAT.Text = oKatOpisPRYZMATKli
                        'txtKatOpisORYG.Text = oKatOpisORYGKli
                        'txtKatOpisNAJEM.Text = oKatOpisNAJEMKli
                        'txtKatOpisSTRONY.Text = oKatOpisSTRONYKli
                        'txtKatOpisDRUKARKI.Text = oKatOpisDRUKARKIKli

                        TextL = oWykazUrzadzenKli

                        i100 = 0
                        TxtLDl = e.Graphics.MeasureString(TextL, MyFont).Width
                        If TxtLDl > 6 * cm2pts(12.5) Then
                            'tekst dluzszy niz mozna wydrukowac...
                            Text = TextL
                            i100 = 0
                            TextL = Mid(Text, i100 * 100 + 1, 100)
                            Do While e.Graphics.MeasureString(TextL, MyFont).Width < 6 * cm2pts(12.5)
                                i100 += 1
                                TextL &= Mid(Text, i100 * 100 + 1, 100)
                            Loop
                        End If
                        IleLiniiL = IleLiniiPotrzebaNaText(TextL, cm2pts(12.5), MyFont, e)







                        TextP = PobierzKontaktyDoTXT(oIdentKli, MyFont, e)
                        i100 = 0
                        TxtPDl = e.Graphics.MeasureString(TextP, MyFont).Width
                        If TxtPDl > 6 * cm2pts(12.5) Then
                            'tekst dluzszy niz mozna wydrukowac...
                            Text = TextP
                            i100 = 0
                            TextP = Mid(Text, i100 * 100 + 1, 100)
                            Do While e.Graphics.MeasureString(TextP, MyFont).Width < 6 * cm2pts(12.5)
                                i100 += 1
                                TextP &= Mid(Text, i100 * 100 + 1, 100)
                            Loop
                        End If
                        IleLiniiP = IleLiniiPotrzebaNaText(TextP, cm2pts(12.5), MyFont, e)

                        'drukuj nie wiêcej ni¿ 6 linii
                        If IleLiniiL > 6 Then IleLiniiL = 6
                        If IleLiniiP > 6 Then IleLiniiP = 6

                        DrukujLinieLNr = 0
                        DrukujLiniePNr = 0

                        If IleLiniiL > IleLiniiP Then
                            IleLiniiTrzebaWydrukowac = IleLiniiL
                        Else
                            IleLiniiTrzebaWydrukowac = IleLiniiP
                        End If

                        If IleLiniiTrzebaWydrukowac > 0 Then
                            y += CType(0.5 * LineHeight, Single)
                            Do While IleLiniiTrzebaWydrukowac > 0
                                If DrukujLinieLNr < IleLiniiL Then
                                    DrukujLinieLNr += 1
                                    LeftTxt(DajLinieTextuNr(TextL, DrukujLinieLNr, cm2pts(12.5), MyFont, e), x + cm2pts(1.5), y, cm2pts(12.5), MyFont, e)
                                End If

                                If DrukujLiniePNr < IleLiniiP Then
                                    DrukujLiniePNr += 1
                                    khText = DajLinieTextuNr(TextP, DrukujLiniePNr, cm2pts(12.5), MyFont, e)
                                    '------------------------------------
                                    'analizuj linie tekstu
                                    ' KLIENT | DATA | OPERATOR | RODZAJ KONTaktu | UWAGI
                                    khKlient = ""
                                    khData = ""
                                    khOperator = ""
                                    khRodzajK = ""
                                    khUwagi = ""

                                    'oddziel pola separator=|
                                    khSepar = InStr(khText, "|")
                                    If khSepar > 0 Then
                                        khKlient = Mid(khText, 1, khSepar - 1)
                                        khText = Mid(khText, khSepar + 1)

                                        khSepar = InStr(khText, "|")
                                        If khSepar > 0 Then
                                            khData = Mid(khText, 1, khSepar - 1)
                                            khText = Mid(khText, khSepar + 1)

                                            khSepar = InStr(khText, "|")
                                            If khSepar > 0 Then
                                                khOperator = Mid(khText, 1, khSepar - 1)
                                                khText = Mid(khText, khSepar + 1)

                                                khSepar = InStr(khText, "|")
                                                If khSepar > 0 Then
                                                    khRodzajK = Mid(khText, 1, khSepar - 1)
                                                    khUwagi = Mid(khText, khSepar + 1)
                                                End If
                                            End If
                                        End If
                                    End If

                                    LeftTxt(khData, x + cm2pts(15.5), y, cm2pts(2), MyFont, e)
                                    LeftTxt(khOperator, x + cm2pts(17.5), y, cm2pts(2), MyFont, e)
                                    LeftTxt(khRodzajK, x + cm2pts(19.5), y, cm2pts(2), MyFont, e)
                                    LeftTxt(khUwagi, x + cm2pts(21.5), y, cm2pts(6.5), MyFont, e)
                                    '------------------------------------
                                End If
                                IleLiniiTrzebaWydrukowac -= 1
                                y += LineHeight

                                If ((y + 5 * LineHeight) > e.MarginBounds.Bottom) Then
                                    Exit Do
                                End If
                            Loop
                        End If

                        ''sprawdz czy mozna drukowaæ dalej klientów...
                        'If IleLiniiTrzebaWydrukowac > 0 Then
                        '    'jeszcze nie skonczylismy - nastepna strona
                        'Else
                        '    'koniec poprzedniego klienta
                        '    y += CType(0.5 * LineHeight, Single)
                        '    e.Graphics.DrawLine(DrawingPen, x, y + CType(0.2 * LineHeight, Single), x + CType(28.5 * (100 / 2.54), Single), y + CType(0.5 * LineHeight, Single))
                        '    y += CType(0.5 * LineHeight, Single)
                        'End If

                    End If
                End If

            Loop Until ((y + 5 * LineHeight) > e.MarginBounds.Bottom) Or (Doc.NextRowNumber >= DataViewKlienci.Count)
        End If

        If (Doc.NextRowNumber < DataViewKlienci.Count) Then
            ' There is still at least one more page.
            ' Signal this event to fire again.
            e.HasMorePages = True
        Else
            ' Printing is complete.
            y += LineHeight
            x = PageLeft 'od marginesu...
            y += CType(0.5 * LineHeight, Single)
            e.Graphics.DrawLine(DrawingPen, x, y + CType(0.2 * LineHeight, Single), x + CType(28.5 * (100 / 2.54), Single), y + CType(0.5 * LineHeight, Single))
            y += CType(0.5 * LineHeight, Single)
            Text = "Koniec zestawienia na pozycji " & Str(Doc.LineNumber)
            e.Graphics.DrawString(Text, MyFont, Brushes.Black, x + CType(0.2 * (100 / 2.54), Single), y)

            Doc.PageNumber = 0
            Doc.LineNumber = 0
            Doc.NextRowNumber = 0

            DrukujLinieLNr = 0
            DrukujLiniePNr = 0
            IleLiniiTrzebaWydrukowac = 0
            IleLiniiL = 0
            IleLiniiP = 0
            TextL = ""
            TextP = ""

            e.HasMorePages = False
        End If

    End Sub


























    Private Sub MyDoc_Etykiety(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        ' Retrieve the document that sent this event.
        ' You could store the document in a class member variable,
        ' but this approach allows you to use the same event handler
        ' to handle multiple print documents at once.
        Dim Doc As Softart.myPrintDocument = CType(sender, PrintDocument)
        Dim LiRow As Integer
        Dim LiEty As Integer

        ' Define the font and determine the line height.
        Dim FontSize As Integer = 10
        Dim MyFont As New System.Drawing.Font("Arial", FontSize, FontStyle.Regular)
        Dim LineHeight As Single = MyFont.GetHeight(e.Graphics)
        Dim DrawingPen As New Pen(Color.Black, 1)

        Dim Nazwisko(2) As String
        Dim Firma1(2) As String
        Dim Firma2(2) As String
        Dim Ulica(2) As String
        Dim Miasto(2) As String
        Dim ii As Integer
        Dim jj As Integer
        Dim ostSP As Integer
        Dim ch As String

        ' Create variables to hold position on page.
        Dim x As Single = e.PageBounds.Left
        Dim y As Single = e.PageBounds.Top

        Dim PageWidth As Single = CType(21 * (100 / 2.54), Single)  'e.PageBounds.Width
        Dim PageHeight As Single = CType(29.5 * (100 / 2.54), Single)
        Dim PageLeft As Single = e.PageBounds.Left
        Dim PageTop As Single = e.PageBounds.Top

        LiEty = 0
        Doc.PageNumber += 1
        x = PageLeft

        'y = PageTop
        y = CType((PageHeight / 8 * LiEty), Single)
        ' petla do konca strony
        If Doc.NextRowNumber < DataViewKlienci.Count Then
            Do
                LiRow = Doc.NextRowNumber
                For ii = 0 To 2
                    If LiRow + ii <= DataViewKlienci.Count - 1 Then
                        dagKlienci.CurrentCell = dagKlienci(0, LiRow)
                        PobierzKlienci()
                        Nazwisko(ii) = OsobaKontaktowa(oIdentKli)

                        Firma1(ii) = ""
                        Firma2(ii) = ""
                        ostSP = 0
                        For jj = 1 To Len(oNazwa1Kli)
                            ch = Mid(oNazwa1Kli, jj, 1)
                            If e.Graphics.MeasureString(Firma1(ii), MyFont).Width > CType(PageWidth / 3 - (1.0 * (100 / 2.54)), Single) Then
                                Firma2(ii) &= ch
                            Else
                                Firma1(ii) &= ch
                                Select Case ch
                                    Case " ", ".", ",", "-", "_"
                                        ostSP = jj
                                    Case Else
                                End Select
                            End If
                        Next
                        If ostSP > 0 Then
                            Firma2(ii) = Mid(Firma1(ii), ostSP + 1) + Firma2(ii)
                            Firma1(ii) = Mid(Firma1(ii), 1, ostSP)
                        End If
                        If Len(oIDDomuKli) > 0 Then
                            Ulica(ii) = Trim(oUlicaKli) & " " & Trim(Str(oNumNrDomuKli)) & " " & Trim(oIDDomuKli)
                        Else
                            Ulica(ii) = Trim(oUlicaKli) & " " & Trim(Str(oNumNrDomuKli))
                        End If
                        Miasto(ii) = Trim(oKodPoczKli) & " " & Trim(oMiejscKli)
                        Doc.NextRowNumber += 1
                    Else
                        Nazwisko(ii) = ""
                        Firma1(ii) = ""
                        Firma2(ii) = ""
                        Ulica(ii) = ""
                        Miasto(ii) = ""
                    End If
                Next ii

                x = PageLeft 'od marginesu...
                y += 2 * LineHeight

                For ii = 0 To 2
                    If Len(Nazwisko(ii)) > 0 Then
                        e.Graphics.DrawString("Sz. P.", MyFont, Brushes.Black, CType((PageWidth / 3 * ii) + 0.5 * (100 / 2.54), Single), y)
                    End If
                Next

                y += 1 * LineHeight
                For ii = 0 To 2
                    If Len(Nazwisko(ii)) > 0 Then
                        e.Graphics.DrawString(Nazwisko(ii), MyFont, Brushes.Black, CType((PageWidth / 3 * ii) + 0.5 * (100 / 2.54), Single), y)
                    End If
                Next

                y += 1 * LineHeight
                For ii = 0 To 2
                    e.Graphics.DrawString(Firma1(ii), MyFont, Brushes.Black, CType((PageWidth / 3 * ii) + 0.5 * (100 / 2.54), Single), y)
                Next

                y += 1 * LineHeight
                For ii = 0 To 2
                    If Len(Firma2(ii)) > 0 Then
                        e.Graphics.DrawString(Firma2(ii), MyFont, Brushes.Black, CType((PageWidth / 3 * ii) + 0.5 * (100 / 2.54), Single), y)
                    Else
                        e.Graphics.DrawString(Ulica(ii), MyFont, Brushes.Black, CType((PageWidth / 3 * ii) + 0.5 * (100 / 2.54), Single), y)
                    End If
                Next

                y += 1 * LineHeight
                For ii = 0 To 2
                    If Len(Firma2(ii)) > 0 Then
                        e.Graphics.DrawString(Ulica(ii), MyFont, Brushes.Black, CType((PageWidth / 3 * ii) + 0.5 * (100 / 2.54), Single), y)
                    Else
                        e.Graphics.DrawString(Miasto(ii), MyFont, Brushes.Black, CType((PageWidth / 3 * ii) + 0.5 * (100 / 2.54), Single), y)
                    End If
                Next

                y += 1 * LineHeight
                For ii = 0 To 2
                    If Len(Firma2(ii)) > 0 Then
                        e.Graphics.DrawString(Miasto(ii), MyFont, Brushes.Black, CType((PageWidth / 3 * ii) + 0.5 * (100 / 2.54), Single), y)
                    Else
                    End If
                Next

                Doc.LineNumber += 1
                LiEty += 1
                y = CType((PageHeight / 8 * LiEty), Single)
            Loop Until ((y + LineHeight) > e.PageBounds.Bottom) Or (Doc.NextRowNumber >= DataViewKlienci.Count)
        End If

        If (Doc.NextRowNumber < DataViewKlienci.Count) Then
            ' There is still at least one more page.
            ' Signal this event to fire again.
            e.HasMorePages = True
        Else
            ' Printing is complete.
            'y += LineHeight
            'x = PageLeft 'od marginesu...
            'Text = "Koniec zestawienia na pozycji " & Str(Doc.LineNumber)
            'e.Graphics.DrawString(Text, MyFont, Brushes.Black, x + CType(0.2 * (100 / 2.54), Single), y)

            Doc.PageNumber = 0
            Doc.LineNumber = 0
            Doc.NextRowNumber = 0

            e.HasMorePages = False
        End If

    End Sub




    Private Sub MyDoc_KartaKlienta(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        ' Retrieve the document that sent this event.
        ' You could store the document in a class member variable,
        ' but this approach allows you to use the same event handler
        ' to handle multiple print documents at once.
        Dim Doc As Softart.myPrintDocument = CType(sender, PrintDocument)
        Dim LiRow As Integer
        Dim Uwagi As String
        Dim Uwagi1 As String
        Dim Uwagi2 As String
        Dim Uwagi3 As String
        Dim Uwagi4 As String
        Dim ii As Integer

        e.PageSettings.Landscape = False

        ' Define the font and determine the line height.
        Dim FontSizeTresc As Integer = 10
        Dim MyFontTresc As New System.Drawing.Font("Arial", FontSizeTresc, FontStyle.Bold)
        Dim LineHeightTresc As Single = MyFontTresc.GetHeight(e.Graphics)

        Dim FontSizeNazwa As Integer = 10
        Dim MyFontNazwa As New System.Drawing.Font("Arial", FontSizeNazwa, FontStyle.Regular)
        Dim LineHeightNazwa As Single = MyFontNazwa.GetHeight(e.Graphics)

        Dim FontSize As Integer = 12
        Dim MyFont As New System.Drawing.Font("Arial", FontSize, FontStyle.Bold)
        Dim LineHeight As Single = MyFont.GetHeight(e.Graphics)

        Dim DrawingPen As New Pen(Color.Black, 1)

        ' Create variables to hold position on page.
        Dim x As Single = e.MarginBounds.Left
        Dim y As Single = e.MarginBounds.Top

        Dim PageWidth As Single = e.MarginBounds.Width
        Dim PageHeight As Single = e.MarginBounds.Height
        Dim PageLeft As Single = e.MarginBounds.Left
        Dim PageTop As Single = e.MarginBounds.Top

        Dim words() As String
        Dim splitChars() As Char = {","c}

        ' petla do konca strony
        If Doc.NextRowNumber < DataViewKlienci.Count Then
            'ustaw kursor na wierszu DataGrid
            LiRow = Doc.NextRowNumber
            dagKlienci.CurrentCell = dagKlienci(0, LiRow)

            PobierzKlienci()
            y = cm2pts(2.0)  'PageTop
            x = PageLeft
            Doc.NextRowNumber += 1



            Uwagi1 = ""
            Uwagi2 = ""
            Uwagi3 = ""
            Uwagi4 = ""
            Uwagi = oSprzedUwagiKli

            ii = InStr(Uwagi, vbLf)
            If ii > 0 Then
                Uwagi1 = Mid(Uwagi, 1, ii - 1)
                Uwagi = Mid(Uwagi, ii + 2)
                ii = InStr(Uwagi, vbLf)
                If ii > 0 Then
                    Uwagi2 = Mid(Uwagi, 1, ii - 1)
                    Uwagi = Mid(Uwagi, ii + 2)
                    ii = InStr(Uwagi, vbLf)
                    If ii > 0 Then
                        Uwagi3 = Mid(Uwagi, 1, ii - 1)
                        Uwagi = Mid(Uwagi, ii + 2)
                        ii = InStr(Uwagi, vbLf)
                        If ii > 0 Then
                            Uwagi4 = Mid(Uwagi, 1, ii - 1)
                            Uwagi = Mid(Uwagi, ii + 2)
                        Else
                            Uwagi4 = Uwagi
                        End If
                    Else
                        Uwagi3 = Uwagi
                    End If
                Else
                    Uwagi2 = Uwagi
                End If
            Else
                Uwagi1 = Uwagi
            End If

            words = oNrTOFITxtKli.Split(splitChars, StringSplitOptions.None)
            If words.Length > 1 Then
            Else
            End If



            'drukuj logo
            'e.Graphics.DrawRectangle(DrawingPen, cm2pts(15), y, cm2pts(5), cm2pts(1))
            'Rectangle(left,top,widthy,height)
            Dim r As System.Drawing.Rectangle
            r = New System.Drawing.Rectangle(cm2pts(14), y, cm2pts(4.0), cm2pts(0.6))
            e.Graphics.DrawImage(picLogo.Image, r)


            'PoleKarty("NUMER KLIENTA ", oIdentKli, cm2pts(6.0), x, y, MyFontNazwa, MyFontTresc, MyFont, e)

            PoleKarty("DATA I KONTAKTU  ", oPKontaktKli, cm2pts(6.0), x, y, MyFontNazwa, MyFontTresc, MyFont, e)
            If words.Length > 1 Then
                PoleKarty("KOD TOFI ", words(0), cm2pts(4.0), x + cm2pts(6.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            Else
                PoleKarty("KOD TOFI ", oNrTOFITxtKli, cm2pts(4.0), x + cm2pts(6.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            End If
            y += 1.5 * LineHeight




            'e.Graphics.DrawRectangle(DrawingPen, x, y + CType(16.7 * LineHeight, Single), cm2pts(16), CType(33 * LineHeight, Single))
            'TytulKarty(" Notatki ", x, y, cm2pts(16.0), MyFontNazwa, e)

            ' numer obrócony o 90 stp
            e.Graphics.TranslateTransform(cm2pts(19.3), y)
            e.Graphics.RotateTransform(90)
            If words.Length > 1 Then
                e.Graphics.DrawString(oIdentKli & "  NUMER KLIENTA          DODATKOWE NUMERY TOFI : " & oNrTOFITxtKli, MyFont, Brushes.Black, 0, 0)
            Else
                e.Graphics.DrawString(oIdentKli & "  NUMER KLIENTA ", MyFont, Brushes.Black, 0, 0)
            End If
            e.Graphics.ResetTransform()



            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(16), CType(((5 * 1.8) + 0.3) * LineHeight, Single))
            TytulKarty(" Identyfikacja ", x, y, cm2pts(16.0), MyFontNazwa, e)

            x = PageLeft
            y += LineHeight
            PoleKarty("Firma ", oNazwa1Kli, cm2pts(15.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.8 * LineHeight
            If Len(oIDDomuKli) > 0 Then
                PoleKarty("Adres ", oUlicaKli & " " & Trim(Str(oNumNrDomuKli)) & " " & oIDDomuKli & ", " & oKodPoczKli & " " & oMiejscKli & ", " & oRejonKli, cm2pts(15.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            Else
                PoleKarty("Adres ", oUlicaKli & " " & Trim(Str(oNumNrDomuKli)) & ", " & oKodPoczKli & " " & oMiejscKli & ", " & oRejonKli, cm2pts(15.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            End If
            y += 1.8 * LineHeight
            PoleKarty("Osoba kontaktowa ", OsobaKontaktowa(oIdentKli), cm2pts(15.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.8 * LineHeight
            PoleKarty("Telefon ", oTelefonKli, cm2pts(15.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.8 * LineHeight
            PoleKarty("eMail ", oeMailKli, cm2pts(15.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.6 * LineHeight



            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(16), CType(((10 * 1.8) + 0.3) * LineHeight, Single))
            TytulKarty(" Charakterystyka ", x, y, cm2pts(16.0), MyFontNazwa, e)

            x = PageLeft
            y += LineHeight
            TytulKarty(" Wykaz urz¹dzeñ/zu¿ycie (m-c,rok) ", x + cm2pts(0.5), y, cm2pts(7.0), MyFontNazwa, e)
            'oWykazUrzadzenKli                      DajLinieTextuNr(oWykazUrzadzenKli, NrLinii, cm2pts(19), MyFont, e)
            PoleKarty("Data weryfikacji potencja³u : ", oDataWeryfSegmentacjiKli, cm2pts(7.0), x + cm2pts(8.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.8 * LineHeight
            PoleKarty("", DajLinieTextuNr(oWykazUrzadzenKli, 1, cm2pts(7), MyFontNazwa, e), cm2pts(7.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            PoleKarty("Liczba urz¹dzeñ : ", oLiczbaUrzadzenKli.ToString("F0"), cm2pts(7.0), x + cm2pts(8.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.8 * LineHeight
            PoleKarty("", DajLinieTextuNr(oWykazUrzadzenKli, 2, cm2pts(7), MyFontNazwa, e), cm2pts(7.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            PoleKarty("Liczba wydruków/rok (tys) : ", oLiczbaWydrukowKli.ToString("F0"), cm2pts(7.0), x + cm2pts(8.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.8 * LineHeight
            PoleKarty("", DajLinieTextuNr(oWykazUrzadzenKli, 3, cm2pts(7), MyFontNazwa, e), cm2pts(7.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            TytulKarty(" Sposób wyboru dostawcy ", x + cm2pts(8.5), y, cm2pts(7.0), MyFontNazwa, e)
            y += 1.8 * LineHeight
            PoleKarty("", DajLinieTextuNr(oWykazUrzadzenKli, 4, cm2pts(7), MyFontNazwa, e), cm2pts(7.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)

            'Lista.Items.Add("Brak danych")
            'Lista.Items.Add("Wolny zakup")
            'Lista.Items.Add("Zapytanie ofertowe")
            'Lista.Items.Add("Przetarg")
            'Lista.Items.Add("Platforma zakupowa")

            leftTextKropki("Wolny zakup", x + cm2pts(8.5), y, cm2pts(2.5), MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(11.0), y, (oSposobWyboruDostawcyKli = "Wolny zakup"), MyFontNazwa, e)
            leftTextKropki("Przetarg", x + cm2pts(12.5), y, cm2pts(2.5), MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(15.0), y, (oSposobWyboruDostawcyKli = "Przetarg"), MyFontNazwa, e)
            y += 1.8 * LineHeight

            PoleKarty("", DajLinieTextuNr(oWykazUrzadzenKli, 5, cm2pts(7), MyFontNazwa, e), cm2pts(7.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            leftTextKropki("Zapyt. ofert.", x + cm2pts(8.5), y, cm2pts(2.5), MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(11.0), y, (oSposobWyboruDostawcyKli = "Zapytanie ofertowe"), MyFontNazwa, e)
            leftTextKropki("Platforma zak.", x + cm2pts(12.5), y, cm2pts(2.5), MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(15.0), y, (oSposobWyboruDostawcyKli = "Platforma zakupowa"), MyFontNazwa, e)
            y += 1.8 * LineHeight

            PoleKarty("", DajLinieTextuNr(oWykazUrzadzenKli, 7, cm2pts(7), MyFontNazwa, e), cm2pts(7.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            TytulKarty(" Forma rozliczenia druku ", x + cm2pts(8.5), y, cm2pts(7.0), MyFontNazwa, e)
            y += 0.9 * LineHeight
            CentrText("Zakup mat.", x + cm2pts(10.0), y, cm2pts(2.0), MyFontNazwa, e)
            CentrText("Rozliczenie", x + cm2pts(12.0), y, cm2pts(2.0), MyFontNazwa, e)
            CentrText("Najem", x + cm2pts(14.0), y, cm2pts(2.0), MyFontNazwa, e)

            y += 0.9 * LineHeight
            PoleKarty("", DajLinieTextuNr(oWykazUrzadzenKli, 8, cm2pts(7), MyFontNazwa, e), cm2pts(7.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            CentrText("eksploat.", x + cm2pts(10.0), y, cm2pts(2.0), MyFontNazwa, e)
            CentrText("w stronach", x + cm2pts(12.0), y, cm2pts(2.0), MyFontNazwa, e)
            CentrText("", x + cm2pts(14.0), y, cm2pts(2.0), MyFontNazwa, e)

            'chkAktZakupMatEksp.Checked = oAktZakupMatEkspKli
            'chkAktRozliczaStrony.Checked = oAktRozliczaStronyDrukuKli
            'chkAktKorzystaZNajmu.Checked = oAktKorzystaZNajmuKli
            'chkZaintRozliczaStrony.Checked = oZaintRozliczaniemStronDrukuKli
            'chkZaintKorzystaZNajmu.Checked = oZaintKorzystaniemZNajmuKli

            y += 0.9 * LineHeight
            leftTextKropki("Stan obecny", x + cm2pts(8.5), y, cm2pts(2.5), MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(11.0), y, oAktZakupMatEkspKli, MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(13.0), y, oAktRozliczaStronyDrukuKli, MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(15.0), y, oAktKorzystaZNajmuKli, MyFontNazwa, e)

            y += 0.9 * LineHeight
            PoleKarty("", DajLinieTextuNr(oWykazUrzadzenKli, 9, cm2pts(7), MyFontNazwa, e), cm2pts(7.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)

            y += 0.9 * LineHeight
            leftTextKropki("Zainteresowanie", x + cm2pts(8.5), y, cm2pts(4.5), MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(13.0), y, oZaintRozliczaniemStronDrukuKli, MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(15.0), y, oZaintKorzystaniemZNajmuKli, MyFontNazwa, e)

            y += 0.9 * LineHeight
            PoleKarty("Wartoœæ rocznych zakupów (z³): ", oZakupPopRokKli.ToString("F2"), cm2pts(7.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.6 * LineHeight




            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(16), CType(((4 * 1.8) + 0.3) * LineHeight, Single))
            TytulKarty(" Rozszerzony Zakres Us³ug ", x, y, cm2pts(16.0), MyFontNazwa, e)

            'chkU1.Checked = (Mid(oUslugiDodKli, 1, 1) = "T")
            'chkU2.Checked = (Mid(oUslugiDodKli, 2, 1) = "T")
            'chkU3.Checked = (Mid(oUslugiDodKli, 3, 1) = "T")
            'chkU4.Checked = (Mid(oUslugiDodKli, 4, 1) = "T")
            'chkU5.Checked = (Mid(oUslugiDodKli, 5, 1) = "T")
            '  chkU6.Checked = (Mid(oUslugiDodKli, 6, 1) = "T")
            '  chkU7.Checked = (Mid(oUslugiDodKli, 7, 1) = "T")
            '  chkU8.Checked = (Mid(oUslugiDodKli, 8, 1) = "T")
            '  chkU9.Checked = (Mid(oUslugiDodKli, 9, 1) = "T")

            x = PageLeft
            y += LineHeight
            leftTextKropki("Platforma druku", x + cm2pts(0.5), y, cm2pts(6.5), MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(7.0), y, (Mid(oUslugiDodKli, 1, 1) = "T"), MyFontNazwa, e)
            leftTextKropki("Przegl¹d techniczny", x + cm2pts(8.5), y, cm2pts(6.5), MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(15.0), y, (Mid(oUslugiDodKli, 5, 1) = "T"), MyFontNazwa, e)
            y += 1.8 * LineHeight
            leftTextKropki("Wsparcie informatyczne", x + cm2pts(0.5), y, cm2pts(6.5), MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(7.0), y, (Mid(oUslugiDodKli, 2, 1) = "T"), MyFontNazwa, e)
            leftTextKropki("Czyszczenie drukarek laserowych", x + cm2pts(8.5), y, cm2pts(6.5), MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(15.0), y, (Mid(oUslugiDodKli, 6, 1) = "T"), MyFontNazwa, e)
            y += 1.8 * LineHeight
            leftTextKropki("Audyt druku", x + cm2pts(0.5), y, cm2pts(6.5), MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(7.0), y, (Mid(oUslugiDodKli, 3, 1) = "T"), MyFontNazwa, e)
            leftTextKropki("Naprawa drukarek laserowych", x + cm2pts(8.5), y, cm2pts(6.5), MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(15.0), y, (Mid(oUslugiDodKli, 7, 1) = "T"), MyFontNazwa, e)
            y += 1.8 * LineHeight
            leftTextKropki("Dobór urz¹dzenia drukuj¹cego", x + cm2pts(0.5), y, cm2pts(6.5), MyFontNazwa, e)
            PoleKratka("", "", x + cm2pts(7.0), y, (Mid(oUslugiDodKli, 4, 1) = "T"), MyFontNazwa, e)
            PoleKarty("Termin obowi¹zywania RZU: ", IIf(Len(Trim(Mid(oUslugiDodKli, 101, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 101, 10)), cm2pts(7.0), x + cm2pts(8.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.6 * LineHeight



            e.Graphics.DrawRectangle(DrawingPen, x, y, cm2pts(16), CType(((7 * 1.8) + 0.3) * LineHeight, Single))
            TytulKarty(" Notatki ", x, y, cm2pts(16.0), MyFontNazwa, e)

            x = PageLeft
            y += LineHeight
            PoleKarty("Co klient robi z pustymi kartrid¿ami ", "", cm2pts(15.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.8 * LineHeight
            PoleKarty("Z jakich materia³ów klient korzysta ", "", cm2pts(15.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.8 * LineHeight
            PoleKarty("Gdzie obecnie siê zaopatruje / najmuje ", "", cm2pts(15.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.8 * LineHeight
            PoleKarty("Oddzia³owoœæ / Lokalizacja Centrali ", "", cm2pts(15.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.8 * LineHeight
            PoleKarty("Dostawa / Forma p³atnoœci ", "", cm2pts(15.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.8 * LineHeight
            PoleKarty("Ustalenia ", "", cm2pts(15.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.7 * LineHeight
            PoleKarty("", "", cm2pts(15.0), x + cm2pts(0.5), y, MyFontNazwa, MyFontTresc, MyFont, e)
            y += 1.6 * LineHeight




            'y += 1 * LineHeight
            'PoleKarty("Sporz¹dzi³ ", oWpisalKli, cm2pts(8), x + cm2pts(7), y, MyFontNazwa, MyFontTresc, MyFont, e)
        End If

        If (Doc.NextRowNumber < DataViewKlienci.Count) Then
            ' There is still at least one more page.
            ' Signal this event to fire again.
            e.HasMorePages = True
        Else
            Doc.PageNumber = 0
            Doc.LineNumber = 0
            Doc.NextRowNumber = 0

            DrukujLinieLNr = 0
            DrukujLiniePNr = 0
            IleLiniiTrzebaWydrukowac = 0
            IleLiniiL = 0
            IleLiniiP = 0
            TextL = ""
            TextP = ""

            e.HasMorePages = False
        End If
    End Sub







    Private Function OgraniczTxt(ByVal TrescPola As String,
                         ByVal DlugoscPola As Long,
                         ByVal FontPola As System.Drawing.Font,
                         ByVal pp As PrintPageEventArgs) As String
        Dim DrawPen As New Pen(Color.Black, 1)
        Dim DlugoscTekstu As Long = pp.Graphics.MeasureString(TrescPola, FontPola).Width
        Do While DlugoscTekstu > DlugoscPola
            TrescPola = Mid(TrescPola, 1, Len(TrescPola) - 1)
            DlugoscTekstu = pp.Graphics.MeasureString(TrescPola, FontPola).Width
        Loop
        Return (TrescPola)
    End Function




    Private Sub CentrText(ByVal Tytul As String,
                          ByVal posX As Single,
                          ByVal posY As Single,
                          ByVal dlugoscPola As Single,
                          ByVal FontTytulu As System.Drawing.Font,
                          ByVal pp As PrintPageEventArgs)

        Tytul = OgraniczTxt(Tytul, dlugoscPola, FontTytulu, pp)
        Dim DlugoscTytulu As Long = pp.Graphics.MeasureString(Tytul, FontTytulu).Width
        If DlugoscTytulu >= dlugoscPola Then
            pp.Graphics.DrawString(Tytul, FontTytulu, Brushes.Black, posX, posY)
        Else
            pp.Graphics.DrawString(Tytul, FontTytulu, Brushes.Black, posX + (dlugoscPola - DlugoscTytulu) / 2, posY)
        End If
    End Sub

    Private Sub leftTextKropki(ByVal Tytul As String,
                          ByVal posX As Single,
                          ByVal posY As Single,
                          ByVal dlugoscPola As Single,
                          ByVal FontTytulu As System.Drawing.Font,
                          ByVal pp As PrintPageEventArgs)

        Tytul = OgraniczTxt(Tytul & " . . . . . . . . . . . . . . . . . .", dlugoscPola, FontTytulu, pp)
        Dim DlugoscTytulu As Long = pp.Graphics.MeasureString(Tytul, FontTytulu).Width
        pp.Graphics.DrawString(Tytul, FontTytulu, Brushes.Black, posX, posY)
    End Sub


    Private Sub TytulKarty(ByVal Tytul As String,
                          ByVal posX As Single,
                          ByVal posY As Single,
                          ByVal dlugoscPola As Single,
                          ByVal FontTytulu As System.Drawing.Font,
                          ByVal pp As PrintPageEventArgs)

        Dim DrawPen As New Pen(Color.Black, 1)
        CentrText(Tytul, posX, posY, dlugoscPola, FontTytulu, pp)
        pp.Graphics.DrawRectangle(DrawPen, posX, posY, dlugoscPola, FontTytulu.GetHeight(pp.Graphics))
    End Sub

    Private Sub PoleKarty(ByVal NazwaPola As String,
                          ByVal TrescPola As String,
                          ByVal DlugoscPola As Long,
                          ByVal posX As Single,
                          ByVal posY As Single,
                          ByVal FontNazwy As System.Drawing.Font,
                          ByVal FontTresci As System.Drawing.Font,
                          ByVal FontPola As System.Drawing.Font,
                          ByVal pp As PrintPageEventArgs)

        Dim DrawPen As New Pen(Color.Black, 1)
        Dim DlugoscNazwy As Long = pp.Graphics.MeasureString(NazwaPola, FontNazwy).Width
        TrescPola = OgraniczTxt(TrescPola, DlugoscPola - DlugoscNazwy, FontTresci, pp)
        pp.Graphics.DrawString(NazwaPola, FontNazwy, Brushes.Black, posX, posY)
        pp.Graphics.DrawString(TrescPola, FontTresci, Brushes.Black, posX + DlugoscNazwy, posY)
        If DlugoscPola > DlugoscNazwy Then
            pp.Graphics.DrawLine(DrawPen, posX + DlugoscNazwy, posY + FontPola.GetHeight(pp.Graphics), posX + DlugoscPola, posY + FontPola.GetHeight(pp.Graphics))
        End If
    End Sub

    Private Sub PoleKartyBezLinii(ByVal NazwaPola As String,
                          ByVal TrescPola As String,
                          ByVal DlugoscPola As Long,
                          ByVal posX As Single,
                          ByVal posY As Single,
                          ByVal FontNazwy As System.Drawing.Font,
                          ByVal FontTresci As System.Drawing.Font,
                          ByVal FontPola As System.Drawing.Font,
                          ByVal pp As PrintPageEventArgs)

        Dim DrawPen As New Pen(Color.Black, 1)
        Dim DlugoscNazwy As Long = pp.Graphics.MeasureString(NazwaPola, FontNazwy).Width
        TrescPola = OgraniczTxt(TrescPola, DlugoscPola - DlugoscNazwy, FontTresci, pp)
        pp.Graphics.DrawString(NazwaPola, FontNazwy, Brushes.Black, posX, posY)
        pp.Graphics.DrawString(TrescPola, FontTresci, Brushes.Black, posX + DlugoscNazwy, posY)
        'If DlugoscPola > DlugoscNazwy Then
        '    pp.Graphics.DrawLine(DrawPen, posX + DlugoscNazwy, posY + FontPola.GetHeight(pp.Graphics), posX + DlugoscPola, posY + FontPola.GetHeight(pp.Graphics))
        'End If
    End Sub

    Private Sub PoleKratka(ByVal txtPrzed As String,
                           ByVal txtPo As String,
                           ByVal posX As Single,
                           ByVal posY As Single,
                           ByVal Zaznacz As Boolean,
                           ByVal Font As System.Drawing.Font,
                           ByVal pp As PrintPageEventArgs)

        Dim DrawPen As New Pen(Color.Black, 1)
        Dim DlugoscNazwy As Long = pp.Graphics.MeasureString(txtPrzed, Font).Width
        pp.Graphics.DrawString(txtPrzed, Font, Brushes.Black, posX, posY)
        pp.Graphics.DrawRectangle(DrawPen, posX + DlugoscNazwy, posY, Font.GetHeight(pp.Graphics), Font.GetHeight(pp.Graphics))
        pp.Graphics.DrawString(txtPo, Font, Brushes.Black, posX + DlugoscNazwy + Font.GetHeight(pp.Graphics), posY)
        If Zaznacz Then
            pp.Graphics.DrawLine(DrawPen, posX + DlugoscNazwy, posY, posX + DlugoscNazwy + Font.GetHeight(pp.Graphics), posY + Font.GetHeight(pp.Graphics))
            pp.Graphics.DrawLine(DrawPen, posX + DlugoscNazwy, posY + Font.GetHeight(pp.Graphics), posX + DlugoscNazwy + Font.GetHeight(pp.Graphics), posY)
        End If
    End Sub









    '**************************************************************
    '** aktualizuj baze danych
    '**************************************************************

    Private Sub AktualizujBaze()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnection.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnection.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapter.Update(DataSetKlienci, "TABELA_Klienci")
            '    Catch Ex As System.Exception
            '        'ViewInLog(ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapter.Fill(DataSetKlienci)
            '    End If
            '    dbConnection.Close()
            'End If
        Else
            Try
                sqlConnectionKlienci.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionKlienci.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterKlienci.Update(DataSetKlienci, "TABELA_Klienci")
                Catch Ex As System.Exception
                    'ViewInLog(ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                End If
                sqlConnectionKlienci.Close()
            End If
        End If

        stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString
    End Sub

    Private Sub OdswiezBaze()
        DataSetKlienci.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnection.Open()
            '    If dbConnection.State = ConnectionState.Open Then
            '        dbDataAdapter.Fill(DataSetKlienci)
            '        dbConnection.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnectionKlienci.Open()
                If sqlConnectionKlienci.State = ConnectionState.Open Then
                    sqlDataAdapterKlienci.Fill(DataSetKlienci)
                    sqlConnectionKlienci.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub






    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'definiuj delegat
        Dim deleg As delegateAktualizuj = AddressOf AktualizujBaze

        oCoMamRobic = "EDYTUJ"
        Dim Form As New KatalogKlientowIKontaktow(DataSetKlienci, deleg)
        Form.ShowDialog()
        oLokalnyFiltr = ""
        oNazwaSchematu = ""
        oSchematFiltrowania = ""
        stbKlienci.Panels(4).Text = "Szablon : "
        OdswiezBaze()
        'dagKlienci.Update()

        'aktualizuj Katalog klientow
        CzyscFiltryDGV(Me.Panel4,
                    dagKlienci,
                    DataViewKlienci.Table.Columns.Count,
                    DataViewKlienci,
                    stbKlienci)

        'CzyscFiltry(Me.Panel4,
        '            dagKlienci,
        '            DataViewKlienci.Table.Columns.Count,
        '            DataViewKlienci,
        '            stbKlienci,
        '            Trim(oSchematFiltrowania))

        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Marker  ", 50, True)
        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Marker P.  ", 50, True)
        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "Aktywny  ", 50, True)

        Dim ctrl As Control
        Dim LiMark As Integer = 0
        For Each ctrl In Me.Panel4.Controls
            If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                If ctrl.Name = "txtFi052" Or ctrl.Name = "txtFi053" Then
                    ctrl.Text = "TRUE"
                    LiMark += 1
                    If LiMark = 2 Then
                        Exit For
                    End If
                End If
            End If
        Next
        dagKlienci.Update()
    End Sub



    Private Sub cmdCoDalej_Click(sender As Object, e As EventArgs) Handles cmdCoDalej.Click
        'definiuj delegat
        Dim deleg As delegateAktualizuj = AddressOf AktualizujBaze

        oCoMamRobic = "EDYTUJ"
        Dim Form As New KatalogKlientowICoDalej(DataSetKlienci, deleg)
        Form.ShowDialog()
        oLokalnyFiltr = ""
        oNazwaSchematu = ""
        oSchematFiltrowania = ""
        stbKlienci.Panels(4).Text = "Szablon : "
        OdswiezBaze()
        'dagKlienci.Update()

        'aktualizuj Katalog klientow
        CzyscFiltryDGV(Me.Panel4,
                    dagKlienci,
                    DataViewKlienci.Table.Columns.Count,
                    DataViewKlienci,
                    stbKlienci)

        'CzyscFiltry(Me.Panel4,
        '            dagKlienci,
        '            DataViewKlienci.Table.Columns.Count,
        '            DataViewKlienci,
        '            stbKlienci,
        '            Trim(oSchematFiltrowania))

        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Marker  ", 50, True)
        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Marker P.  ", 50, True)
        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "Aktywny  ", 50, True)

        Dim ctrl As Control
        Dim LiMark As Integer = 0
        For Each ctrl In Me.Panel4.Controls
            If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                If ctrl.Name = "txtFi052" Or ctrl.Name = "txtFi053" Then
                    ctrl.Text = "TRUE"
                    LiMark += 1
                    If LiMark = 2 Then
                        Exit For
                    End If
                End If
            End If
        Next
        dagKlienci.Update()
    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim Form As New frmAnalizaCzestKontaktowGrupowa
        'Form.ShowDialog()

        StartFormy = True
        'definiuj delegat
        Dim deleg As delegateAktualizuj = AddressOf AktualizujBaze
        Dim Form As New frmAnalizaCzestKontaktowGrupowa(DataSetKlienci, deleg)
        Form.ShowDialog()

        oLokalnyFiltr = ""
        oNazwaSchematu = ""
        oSchematFiltrowania = ""
        stbKlienci.Panels(4).Text = "Szablon : "
        OdswiezBaze()
        'dagKlienci.Update()
        StartFormy = False

        'aktualizuj Katalog klientow
        CzyscFiltryDGV(Me.Panel4,
                    dagKlienci,
                    DataViewKlienci.Table.Columns.Count,
                    DataViewKlienci,
                    stbKlienci)

        'CzyscFiltry(Me.Panel4,
        '            dagKlienci,
        '            DataViewKlienci.Table.Columns.Count,
        '            DataViewKlienci,
        '            stbKlienci,
        '            Trim(oSchematFiltrowania))

        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Marker  ", 50, True)
        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Marker P.  ", 50, True)
        'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "Aktywny  ", 50, True)

        Dim ctrl As Control
        Dim LiMark As Integer = 0
        For Each ctrl In Me.Panel4.Controls
            If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                If ctrl.Name = "txtFi052" Or ctrl.Name = "txtFi053" Then
                    ctrl.Text = "TRUE"
                    LiMark += 1
                    If LiMark = 2 Then
                        Exit For
                    End If
                End If
            End If
        Next
        dagKlienci.Update()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If DataSetKlienci.Tables(0).Rows.Count <= 0 Then
            MessageBox.Show("Brak informacji o kliencie." & vbCrLf &
                "Proszê najpierw dopisaæ klienta do Kartoteki...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            PobierzKlienci()
            Dim Form As New KatalogObrotow(False)
            Form.ShowDialog()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If DataSetKlienci.Tables(0).Rows.Count <= 0 Then
            MessageBox.Show("Brak informacji o kliencie." & vbCrLf &
                "Proszê najpierw dopisaæ klienta do Kartoteki...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            PobierzKlienci()
            Dim Form As New KatalogObrotowMies(False)
            Form.ShowDialog()
        End If
    End Sub

















    '*************************************************
    '** usuwanie zapisow z katalogu OSOBY
    '*************************************************

    Dim sSelectSQLOsoby As String = "SELECT * FROM Osoby WHERE IDENTKLIENTA='" & oIdentKli & "'"

    Dim dbConnectionOsoby As OleDb.OleDbConnection
    Dim dbCommandSelectOsoby As OleDb.OleDbCommand
    Dim dbCommandDeleteOsoby As OleDb.OleDbCommand
    Dim dbDataAdapterOsoby As OleDb.OleDbDataAdapter

    Dim sqlConnectionKlienciOsoby As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienciOsoby As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienciOsoby As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienciOsoby As SqlClient.SqlDataAdapter
    Dim DataSetOsoby As New DataSet
    Dim DataViewOsoby As New DataView

    Private Sub UsunOsoby(ByVal IdKl As String)
        sSelectSQLOsoby = "SELECT * FROM Osoby WHERE IDENTKLIENTA='" & IdKl & "'"

        DataSetOsoby = New DataSet
        DataSetOsoby.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            ''OleDbConnection-polaczenie
            'dbConnectionOsoby = New OleDb.OleDbConnection(sConnectionOsoby)
            'dbCommandSelectOsoby = New OleDb.OleDbCommand(sSelectSQLOsoby, dbConnectionOsoby)
            'dbCommandDeleteOsoby = New OleDb.OleDbCommand(sDeleteSQLOsoby, dbConnectionOsoby)
            'dbDataAdapterOsoby = New OleDb.OleDbDataAdapter(dbCommandSelectOsoby)
            ''---------------------------------------------
            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeliOsoby As System.Data.Common.DataTableMapping
            'MapowanieTabeliOsoby = dbDataAdapterOsoby.TableMappings.Add("Table", "TABELA_Osoby")

            ''komendy DataAdapter
            'Dim objParamOsoby As OleDb.OleDbParameter
            ''------------------------------------------
            ''komenda DELETE
            ''parametry filtrowania
            'objParamOsoby = dbCommandDeleteOsoby.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
            'objParamOsoby.SourceVersion = DataRowVersion.Original
            'objParamOsoby = dbCommandDeleteOsoby.Parameters.Add("@orygOsoba", OleDb.OleDbType.VarChar, 50, "OSOBA")
            'objParamOsoby.SourceVersion = DataRowVersion.Original
            'objParamOsoby = dbCommandDeleteOsoby.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            'objParamOsoby.SourceVersion = DataRowVersion.Original
            'dbDataAdapterOsoby.DeleteCommand = dbCommandDeleteOsoby

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterOsoby.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionOsoby.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'End Try

            'If dbConnectionOsoby.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterOsoby.FillSchema(DataSetOsoby, SchemaType.Mapped, "TABELA_Osoby")
            '        dbDataAdapterOsoby.Fill(DataSetOsoby)
            '        dbConnectionOsoby.Close()

            '        'zdefiniuj unikalny klucz wyszukiwania w tabeli
            '        DataSetOsoby.Tables(0).PrimaryKey = New DataColumn() {DataSetOsoby.Tables(0).Columns("IDENTKLIENTA"), DataSetOsoby.Tables(0).Columns("OSOBA")}

            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '        '    System.Windows.Forms.MessageBoxButtons.OK, _
            '        '    MessageBoxIcon.Information, _
            '        '    MessageBoxDefaultButton.Button1)
            '    End Try
            '    '---------------------------------
            'End If
        Else



            sqlConnectionKlienciOsoby = New SqlClient.SqlConnection(sConnectionOsoby)
            sqlCommandSelectKlienciOsoby = New SqlClient.SqlCommand(sSelectSQLOsoby, sqlConnectionKlienciOsoby)
            sqlCommandDeleteKlienciOsoby = New SqlClient.SqlCommand(sDeleteSQLOsoby, sqlConnectionKlienciOsoby)
            sqlDataAdapterKlienciOsoby = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienciOsoby)
            '---------------------------------------------
            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliOsoby As System.Data.Common.DataTableMapping
            MapowanieTabeliOsoby = sqlDataAdapterKlienciOsoby.TableMappings.Add("Table", "TABELA_Osoby")

            'komendy DataAdapter
            Dim objParamOsoby As SqlClient.SqlParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            objParamOsoby = sqlCommandDeleteKlienciOsoby.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 6, "IDENTKLIENTA")
            objParamOsoby.SourceVersion = DataRowVersion.Original
            objParamOsoby = sqlCommandDeleteKlienciOsoby.Parameters.Add("@orygOsoba", SqlDbType.VarChar, 50, "OSOBA")
            objParamOsoby.SourceVersion = DataRowVersion.Original
            objParamOsoby = sqlCommandDeleteKlienciOsoby.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParamOsoby.SourceVersion = DataRowVersion.Original
            sqlDataAdapterKlienciOsoby.DeleteCommand = sqlCommandDeleteKlienciOsoby

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienciOsoby.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienciOsoby.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            If sqlConnectionKlienciOsoby.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterKlienciOsoby.FillSchema(DataSetOsoby, SchemaType.Mapped, "TABELA_Osoby")
                    sqlDataAdapterKlienciOsoby.Fill(DataSetOsoby)
                    sqlConnectionKlienciOsoby.Close()

                    'zdefiniuj unikalny klucz wyszukiwania w tabeli
                    DataSetOsoby.Tables(0).PrimaryKey = New DataColumn() {DataSetOsoby.Tables(0).Columns("IDENTKLIENTA"), DataSetOsoby.Tables(0).Columns("OSOBA")}

                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                End Try
                '---------------------------------
            End If
        End If

        Dim dr As DataRow = Nothing
        Dim Id As String = ""
        Dim i As Integer = 0
        For i = DataSetOsoby.Tables(0).Rows.Count - 1 To 0 Step -1
            dr = DataSetOsoby.Tables(0).Rows(i)
            Id = dr.Item("IDENTKLIENTA")
            dr.Delete()
        Next
        AktualizujOsoby()
    End Sub

    Private Sub AktualizujOsoby()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionOsoby.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnectionOsoby.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapterOsoby.Update(DataSetOsoby, "TABELA_Osoby")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapterOsoby.Fill(DataSetOsoby)
            '    End If
            '    dbConnectionOsoby.Close()
            'End If
        Else
            Try
                sqlConnectionKlienciOsoby.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionKlienciOsoby.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterKlienciOsoby.Update(DataSetOsoby, "TABELA_Osoby")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienciOsoby.Fill(DataSetOsoby)
                End If
                sqlConnectionKlienciOsoby.Close()
            End If
        End If
    End Sub











    '*************************************************
    '** usuwanie zapisow z katalogu Kontakty
    '*************************************************

    Dim sSelectSQLKontakty As String = "SELECT * FROM KontaktyHandlowe WHERE IDENTKLIENTA='" & oIdentKli & "'"

    Dim dbConnectionKontakty As OleDb.OleDbConnection
    Dim dbCommandSelectKontakty As OleDb.OleDbCommand
    Dim dbCommandDeleteKontakty As OleDb.OleDbCommand
    Dim dbDataAdapterKontakty As OleDb.OleDbDataAdapter

    Dim sqlConnectionKlienciKontakty As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienciKontakty As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienciKontakty As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienciKontakty As SqlClient.SqlDataAdapter

    Dim DataSetKontakty As New DataSet
    Dim DataViewKontakty As New DataView

    Private Sub UsunKontakty(ByVal IdKl As String)
        sSelectSQLKontakty = "SELECT * FROM KontaktyHandlowe WHERE IDENTKLIENTA='" & IdKl & "'"

        DataSetKontakty = New DataSet
        DataSetKontakty.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            ''OleDbConnection-polaczenie
            'dbConnectionKontakty = New OleDb.OleDbConnection(sConnectionKontakty)
            'dbCommandSelectKontakty = New OleDb.OleDbCommand(sSelectSQLKontakty, dbConnectionKontakty)
            'dbCommandDeleteKontakty = New OleDb.OleDbCommand(sDeleteSQLKontakty, dbConnectionKontakty)
            'dbDataAdapterKontakty = New OleDb.OleDbDataAdapter(dbCommandSelectKontakty)
            ''---------------------------------------------
            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeliKontakty As System.Data.Common.DataTableMapping
            'MapowanieTabeliKontakty = dbDataAdapterKontakty.TableMappings.Add("Table", "TABELA_Kontakty")

            'DBDeleteKontaktyHandlowe(dbCommandDeleteKontakty, dbDataAdapterKontakty)
            ''DBUpdateKontaktyHandlowe(dbCommandUpdate, dbDataAdapter)
            ''DBInsertKontaktyHandlowe(dbCommandInsert, dbDataAdapter)

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterKontakty.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionKontakty.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'End Try

            'If dbConnectionKontakty.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
            '        dbDataAdapterKontakty.Fill(DataSetKontakty)
            '        dbConnectionKontakty.Close()

            '        'zdefiniuj unikalny klucz wyszukiwania w tabeli
            '        DataSetKontakty.Tables(0).PrimaryKey = New DataColumn() {DataSetKontakty.Tables(0).Columns("UNIQID")}

            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '        '    System.Windows.Forms.MessageBoxButtons.OK, _
            '        '    MessageBoxIcon.Information, _
            '        '    MessageBoxDefaultButton.Button1)
            '    End Try
            '    '---------------------------------
            'End If
        Else
            sqlConnectionKlienciKontakty = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectKlienciKontakty = New SqlClient.SqlCommand(sSelectSQLKontakty, sqlConnectionKlienciKontakty)
            sqlCommandDeleteKlienciKontakty = New SqlClient.SqlCommand(sDeleteSQLKontakty, sqlConnectionKlienciKontakty)
            sqlDataAdapterKlienciKontakty = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienciKontakty)
            '---------------------------------------------
            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliKontakty As System.Data.Common.DataTableMapping
            MapowanieTabeliKontakty = sqlDataAdapterKlienciKontakty.TableMappings.Add("Table", "TABELA_Kontakty")

            SQLDeleteKontaktyHandlowe(sqlCommandDeleteKlienciKontakty, sqlDataAdapterKlienciKontakty)
            'SQLUpdateKontaktyHandlowe(sqlCommandUpdateKlienci, sqlDataAdapterKlienci)
            'SQLInsertKontaktyHandlowe(sqlCommandInsertKlienci, sqlDataAdapterKlienci)

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienciKontakty.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienciKontakty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            If sqlConnectionKlienciKontakty.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterKlienciKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
                    sqlDataAdapterKlienciKontakty.Fill(DataSetKontakty)
                    sqlConnectionKlienciKontakty.Close()

                    'zdefiniuj unikalny klucz wyszukiwania w tabeli
                    DataSetKontakty.Tables(0).PrimaryKey = New DataColumn() {DataSetKontakty.Tables(0).Columns("UNIQID")}

                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                End Try
                '---------------------------------
            End If
        End If

        Dim dr As DataRow = Nothing
        Dim Id As String = ""
        Dim i As Integer = 0
        For i = DataSetKontakty.Tables(0).Rows.Count - 1 To 0 Step -1
            dr = DataSetKontakty.Tables(0).Rows(i)
            Id = dr.Item("IDENTKLIENTA")
            dr.Delete()
        Next
        AktualizujKontakty()
    End Sub

    Private Sub AktualizujKontakty()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionKontakty.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnectionKontakty.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapterKontakty.Update(DataSetKontakty, "TABELA_Kontakty")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapterKontakty.Fill(DataSetKontakty)
            '    End If
            '    dbConnectionKontakty.Close()
            'End If
        Else
            Try
                sqlConnectionKlienciKontakty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionKlienciKontakty.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterKlienciKontakty.Update(DataSetKontakty, "TABELA_Kontakty")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienciKontakty.Fill(DataSetKontakty)
                End If
                sqlConnectionKlienciKontakty.Close()
            End If
        End If
    End Sub







    '*************************************************
    '** usuwanie zapisow z katalogu CoDalej
    '*************************************************

    Dim sSelectSQLCoDalej As String = "SELECT * FROM CoDalejHandlowe WHERE IDENTKLIENTA='" & oIdentKli & "'"

    Dim dbConnectionCoDalej As OleDb.OleDbConnection
    Dim dbCommandSelectCoDalej As OleDb.OleDbCommand
    Dim dbCommandDeleteCoDalej As OleDb.OleDbCommand
    Dim dbDataAdapterCoDalej As OleDb.OleDbDataAdapter

    Dim sqlConnectionKlienciCoDalej As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienciCoDalej As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienciCoDalej As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienciCoDalej As SqlClient.SqlDataAdapter

    Dim DataSetCoDalej As New DataSet
    Dim DataViewCoDalej As New DataView

    'Public cdUNIQID As String      'UNIQID Text, 40
    'Public cdIDENTKLIENTA As String      'IDENTKLIENTA Text, 40
    'Public cdIDENT As String             'IDENT        Text, 40
    'Public cdOPIS As String              'OPIS         memo
    'Public cdWersja As Integer           'WERSJA       Integer

    Private Sub UsunCoDalej(ByVal IdKl As String)
        sSelectSQLCoDalej = "SELECT * FROM CoDalejPlan WHERE IDENTKLIENTA='" & IdKl & "'"

        DataSetCoDalej = New DataSet
        DataSetCoDalej.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            ''OleDbConnection-polaczenie
            'dbConnectionCoDalej = New OleDb.OleDbConnection(sConnectionCoDalej)
            'dbCommandSelectCoDalej = New OleDb.OleDbCommand(sSelectSQLCoDalej, dbConnectionCoDalej)
            'dbCommandDeleteCoDalej = New OleDb.OleDbCommand(sDeleteSQLCoDalej, dbConnectionCoDalej)
            'dbDataAdapterCoDalej = New OleDb.OleDbDataAdapter(dbCommandSelectCoDalej)
            ''---------------------------------------------
            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeliCoDalej As System.Data.Common.DataTableMapping
            'MapowanieTabeliCoDalej = dbDataAdapterCoDalej.TableMappings.Add("Table", "TABELA_CoDalej")

            'DBDeleteCoDalej(dbCommandDeleteCoDalej, dbDataAdapterCoDalej)
            ''DBUpdateCoDalej(dbCommandUpdate, dbDataAdapter)
            ''DBInsertCoDalej(dbCommandInsert, dbDataAdapter)

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterCoDalej.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionCoDalej.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'End Try

            'If dbConnectionCoDalej.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterCoDalej.FillSchema(DataSetCoDalej, SchemaType.Mapped, "TABELA_CoDalej")
            '        dbDataAdapterCoDalej.Fill(DataSetCoDalej)
            '        dbConnectionCoDalej.Close()

            '        'zdefiniuj unikalny klucz wyszukiwania w tabeli
            '        DataSetCoDalej.Tables(0).PrimaryKey = New DataColumn() {DataSetCoDalej.Tables(0).Columns("UNIQID"), DataSetCoDalej.Tables(0).Columns("IDENTKLIENTA")}

            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '        '    System.Windows.Forms.MessageBoxButtons.OK, _
            '        '    MessageBoxIcon.Information, _
            '        '    MessageBoxDefaultButton.Button1)
            '    End Try
            '    '---------------------------------
            'End If
        Else
            sqlConnectionKlienciCoDalej = New SqlClient.SqlConnection(sConnectionCoDalej)
            sqlCommandSelectKlienciCoDalej = New SqlClient.SqlCommand(sSelectSQLCoDalej, sqlConnectionKlienciCoDalej)
            sqlCommandDeleteKlienciCoDalej = New SqlClient.SqlCommand(sDeleteSQLCoDalej, sqlConnectionKlienciCoDalej)
            sqlDataAdapterKlienciCoDalej = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienciCoDalej)
            '---------------------------------------------
            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliCoDalej As System.Data.Common.DataTableMapping
            MapowanieTabeliCoDalej = sqlDataAdapterKlienciCoDalej.TableMappings.Add("Table", "TABELA_CoDalej")

            SQLDeleteCoDalej(sqlCommandDeleteKlienciCoDalej, sqlDataAdapterKlienciCoDalej)
            'SQLUpdateCoDalej(sqlCommandUpdateKlienci, sqlDataAdapterKlienci)
            'SQLInsertCoDalej(sqlCommandInsertKlienci, sqlDataAdapterKlienci)

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienciCoDalej.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienciCoDalej.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            If sqlConnectionKlienciCoDalej.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterKlienciCoDalej.FillSchema(DataSetCoDalej, SchemaType.Mapped, "TABELA_CoDalej")
                    sqlDataAdapterKlienciCoDalej.Fill(DataSetCoDalej)
                    sqlConnectionKlienciCoDalej.Close()

                    'zdefiniuj unikalny klucz wyszukiwania w tabeli
                    DataSetCoDalej.Tables(0).PrimaryKey = New DataColumn() {DataSetCoDalej.Tables(0).Columns("UNIQID"), DataSetCoDalej.Tables(0).Columns("IDENTKLIENTA")}

                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                End Try
                '---------------------------------
            End If
        End If

        Dim dr As DataRow = Nothing
        Dim Id As String = ""
        Dim i As Integer = 0
        For i = DataSetCoDalej.Tables(0).Rows.Count - 1 To 0 Step -1
            dr = DataSetCoDalej.Tables(0).Rows(i)
            Id = dr.Item("IDENTKLIENTA")
            dr.Delete()
        Next
        AktualizujCoDalej()
    End Sub

    Private Sub AktualizujCoDalej()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionCoDalej.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnectionCoDalej.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapterCoDalej.Update(DataSetCoDalej, "TABELA_CoDalej")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapterCoDalej.Fill(DataSetCoDalej)
            '    End If
            '    dbConnectionCoDalej.Close()
            'End If
        Else
            Try
                sqlConnectionKlienciCoDalej.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionKlienciCoDalej.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterKlienciCoDalej.Update(DataSetCoDalej, "TABELA_CoDalej")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienciCoDalej.Fill(DataSetCoDalej)
                End If
                sqlConnectionKlienciCoDalej.Close()
            End If
        End If
    End Sub






    '*************************************************
    '** usuwanie zapisow z katalogu Obroty
    '*************************************************

    Dim sSelectSQLObroty As String = "SELECT * FROM Obroty WHERE IDENTKLIENTA='" & oIdentKli & "'"

    Dim dbConnectionObroty As OleDb.OleDbConnection
    Dim dbCommandSelectObroty As OleDb.OleDbCommand
    Dim dbCommandDeleteObroty As OleDb.OleDbCommand
    Dim dbDataAdapterObroty As OleDb.OleDbDataAdapter

    Dim sqlConnectionKlienciObroty As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienciObroty As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienciObroty As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienciObroty As SqlClient.SqlDataAdapter

    Dim DataSetObroty As New DataSet
    Dim DataViewObroty As New DataView

    Private Sub UsunObroty(ByVal IdKl As String)
        sSelectSQLObroty = "SELECT * FROM Obroty WHERE IDENTKLIENTA='" & IdKl & "'"

        DataSetObroty = New DataSet
        DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            ''OleDbConnection-polaczenie
            'dbConnectionObroty = New OleDb.OleDbConnection(sConnectionObroty)
            'dbCommandSelectObroty = New OleDb.OleDbCommand(sSelectSQLObroty, dbConnectionObroty)
            'dbCommandDeleteObroty = New OleDb.OleDbCommand(sDeleteSQLObroty, dbConnectionObroty)
            'dbDataAdapterObroty = New OleDb.OleDbDataAdapter(dbCommandSelectObroty)
            ''---------------------------------------------
            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeliObroty As System.Data.Common.DataTableMapping
            'MapowanieTabeliObroty = dbDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")

            ''komendy DataAdapter
            'Dim objParamObroty As OleDb.OleDbParameter
            ''------------------------------------------
            ''komenda DELETE
            ''parametry filtrowania
            'objParamObroty = dbCommandDeleteObroty.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
            'objParamObroty.SourceVersion = DataRowVersion.Original
            'objParamObroty = dbCommandDeleteObroty.Parameters.Add("@orygData", OleDb.OleDbType.VarChar, 10, "DATA")
            'objParamObroty.SourceVersion = DataRowVersion.Original
            'objParamObroty = dbCommandDeleteObroty.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            'objParamObroty.SourceVersion = DataRowVersion.Original
            'dbDataAdapterObroty.DeleteCommand = dbCommandDeleteObroty

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterObroty.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionObroty.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'End Try

            'If dbConnectionObroty.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
            '        dbDataAdapterObroty.Fill(DataSetObroty)
            '        dbConnectionObroty.Close()

            '        'zdefiniuj unikalny klucz wyszukiwania w tabeli
            '        DataSetObroty.Tables(0).PrimaryKey = New DataColumn() {DataSetObroty.Tables(0).Columns("IDENTKLIENTA"), DataSetObroty.Tables(0).Columns("DATA")}

            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '        '    System.Windows.Forms.MessageBoxButtons.OK, _
            '        '    MessageBoxIcon.Information, _
            '        '    MessageBoxDefaultButton.Button1)
            '    End Try
            '    '---------------------------------
            'End If
        Else
            sqlConnectionKlienciObroty = New SqlClient.SqlConnection(sConnectionObroty)
            sqlCommandSelectKlienciObroty = New SqlClient.SqlCommand(sSelectSQLObroty, sqlConnectionKlienciObroty)
            sqlCommandDeleteKlienciObroty = New SqlClient.SqlCommand(sDeleteSQLObroty, sqlConnectionKlienciObroty)
            sqlDataAdapterKlienciObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienciObroty)
            '---------------------------------------------
            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliObroty As System.Data.Common.DataTableMapping
            MapowanieTabeliObroty = sqlDataAdapterKlienciObroty.TableMappings.Add("Table", "TABELA_Obroty")

            'komendy DataAdapter
            Dim objParamObroty As SqlClient.SqlParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            objParamObroty = sqlCommandDeleteKlienciObroty.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 6, "IDENTKLIENTA")
            objParamObroty.SourceVersion = DataRowVersion.Original
            objParamObroty = sqlCommandDeleteKlienciObroty.Parameters.Add("@orygData", SqlDbType.VarChar, 10, "DATA")
            objParamObroty.SourceVersion = DataRowVersion.Original
            objParamObroty = sqlCommandDeleteKlienciObroty.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParamObroty.SourceVersion = DataRowVersion.Original
            sqlDataAdapterKlienciObroty.DeleteCommand = sqlCommandDeleteKlienciObroty

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienciObroty.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienciObroty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            If sqlConnectionKlienciObroty.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterKlienciObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    sqlDataAdapterKlienciObroty.Fill(DataSetObroty)
                    sqlConnectionKlienciObroty.Close()

                    'zdefiniuj unikalny klucz wyszukiwania w tabeli
                    DataSetObroty.Tables(0).PrimaryKey = New DataColumn() {DataSetObroty.Tables(0).Columns("IDENTKLIENTA"), DataSetObroty.Tables(0).Columns("DATA")}

                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                End Try
                '---------------------------------
            End If
        End If

        Dim dr As DataRow = Nothing
        Dim Id As String = ""
        Dim i As Integer = 0
        For i = DataSetObroty.Tables(0).Rows.Count - 1 To 0 Step -1
            dr = DataSetObroty.Tables(0).Rows(i)
            Id = dr.Item("IDENTKLIENTA")
            dr.Delete()
        Next
        AktualizujObroty()
    End Sub

    Private Sub AktualizujObroty()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionObroty.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnectionObroty.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapterObroty.Update(DataSetObroty, "TABELA_Obroty")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapterObroty.Fill(DataSetObroty)
            '    End If
            '    dbConnectionObroty.Close()
            'End If
        Else
            Try
                sqlConnectionKlienciObroty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionKlienciObroty.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterKlienciObroty.Update(DataSetObroty, "TABELA_Obroty")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienciObroty.Fill(DataSetObroty)
                End If
                sqlConnectionKlienciObroty.Close()
            End If
        End If
    End Sub










    '*************************************************
    '** usuwanie zapisow z katalogu ObrotyMies
    '*************************************************

    Dim sSelectSQLObrotyMies As String = "SELECT * FROM ObrotyMies WHERE IDENTKLIENTA='" & oIdentKli & "'"

    Dim dbConnectionObrotyMies As OleDb.OleDbConnection
    Dim dbCommandSelectObrotyMies As OleDb.OleDbCommand
    Dim dbCommandDeleteObrotyMies As OleDb.OleDbCommand
    Dim dbDataAdapterObrotyMies As OleDb.OleDbDataAdapter

    Dim sqlConnectionKlienciObrotyMies As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienciObrotyMies As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienciObrotyMies As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienciObrotyMies As SqlClient.SqlDataAdapter

    Dim DataSetObrotyMies As New DataSet
    Dim DataViewObrotyMies As New DataView

    Private Sub UsunObrotyMies(ByVal IdKl As String)
        sSelectSQLObrotyMies = "SELECT * FROM ObrotyMies WHERE IDENTKLIENTA='" & IdKl & "'"

        DataSetObrotyMies = New DataSet
        DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            ''OleDbConnection-polaczenie
            'dbConnectionObrotyMies = New OleDb.OleDbConnection(sConnectionObrotyMies)
            'dbCommandSelectObrotyMies = New OleDb.OleDbCommand(sSelectSQLObrotyMies, dbConnectionObrotyMies)
            'dbCommandDeleteObrotyMies = New OleDb.OleDbCommand(sDeleteSQLObrotyMies, dbConnectionObrotyMies)
            'dbDataAdapterObrotyMies = New OleDb.OleDbDataAdapter(dbCommandSelectObrotyMies)
            ''---------------------------------------------
            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            'MapowanieTabeliObrotyMies = dbDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")

            ''komendy DataAdapter
            'Dim objParamObrotyMies As OleDb.OleDbParameter
            ''------------------------------------------
            ''komenda DELETE
            ''parametry filtrowania
            'objParamObrotyMies = dbCommandDeleteObrotyMies.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 6, "IDENTKLIENTA")
            'objParamObrotyMies.SourceVersion = DataRowVersion.Original
            'objParamObrotyMies = dbCommandDeleteObrotyMies.Parameters.Add("@orygMies", OleDb.OleDbType.VarChar, 7, "MIESIAC")
            'objParamObrotyMies.SourceVersion = DataRowVersion.Original
            'objParamObrotyMies = dbCommandDeleteObrotyMies.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            'objParamObrotyMies.SourceVersion = DataRowVersion.Original
            'dbDataAdapterObrotyMies.DeleteCommand = dbCommandDeleteObrotyMies

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterObrotyMies.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionObrotyMies.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'End Try

            'If dbConnectionObrotyMies.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
            '        dbDataAdapterObrotyMies.Fill(DataSetObrotyMies)
            '        dbConnectionObrotyMies.Close()

            '        'zdefiniuj unikalny klucz wyszukiwania w tabeli
            '        DataSetObrotyMies.Tables(0).PrimaryKey = New DataColumn() {DataSetObrotyMies.Tables(0).Columns("IDENTKLIENTA"), DataSetObrotyMies.Tables(0).Columns("MIESIAC")}

            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '        '    System.Windows.Forms.MessageBoxButtons.OK, _
            '        '    MessageBoxIcon.Information, _
            '        '    MessageBoxDefaultButton.Button1)
            '    End Try
            '    '---------------------------------
            'End If
        Else
            sqlConnectionKlienciObrotyMies = New SqlClient.SqlConnection(sConnectionObrotyMies)
            sqlCommandSelectKlienciObrotyMies = New SqlClient.SqlCommand(sSelectSQLObrotyMies, sqlConnectionKlienciObrotyMies)
            sqlCommandDeleteKlienciObrotyMies = New SqlClient.SqlCommand(sDeleteSQLObrotyMies, sqlConnectionKlienciObrotyMies)
            sqlDataAdapterKlienciObrotyMies = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienciObrotyMies)
            '---------------------------------------------
            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            MapowanieTabeliObrotyMies = sqlDataAdapterKlienciObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")

            'komendy DataAdapter
            Dim objParamObrotyMies As SqlClient.SqlParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            objParamObrotyMies = sqlCommandDeleteKlienciObrotyMies.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 6, "IDENTKLIENTA")
            objParamObrotyMies.SourceVersion = DataRowVersion.Original
            objParamObrotyMies = sqlCommandDeleteKlienciObrotyMies.Parameters.Add("@orygMies", SqlDbType.VarChar, 7, "MIESIAC")
            objParamObrotyMies.SourceVersion = DataRowVersion.Original
            objParamObrotyMies = sqlCommandDeleteKlienciObrotyMies.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParamObrotyMies.SourceVersion = DataRowVersion.Original
            sqlDataAdapterKlienciObrotyMies.DeleteCommand = sqlCommandDeleteKlienciObrotyMies

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienciObrotyMies.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienciObrotyMies.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            If sqlConnectionKlienciObrotyMies.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterKlienciObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    sqlDataAdapterKlienciObrotyMies.Fill(DataSetObrotyMies)
                    sqlConnectionKlienciObrotyMies.Close()

                    'zdefiniuj unikalny klucz wyszukiwania w tabeli
                    DataSetObrotyMies.Tables(0).PrimaryKey = New DataColumn() {DataSetObrotyMies.Tables(0).Columns("IDENTKLIENTA"), DataSetObrotyMies.Tables(0).Columns("MIESIAC")}

                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                End Try
                '---------------------------------
            End If
        End If

        Dim dr As DataRow = Nothing
        Dim Id As String = ""
        Dim i As Integer = 0
        For i = DataSetObrotyMies.Tables(0).Rows.Count - 1 To 0 Step -1
            dr = DataSetObrotyMies.Tables(0).Rows(i)
            Id = dr.Item("IDENTKLIENTA")
            dr.Delete()
        Next
        AktualizujObrotyMies()
    End Sub

    Private Sub AktualizujObrotyMies()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionObrotyMies.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnectionObrotyMies.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapterObrotyMies.Update(DataSetObrotyMies, "TABELA_ObrotyMies")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapterObrotyMies.Fill(DataSetObrotyMies)
            '    End If
            '    dbConnectionObrotyMies.Close()
            'End If
        Else
            Try
                sqlConnectionKlienciObrotyMies.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionKlienciObrotyMies.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterKlienciObrotyMies.Update(DataSetObrotyMies, "TABELA_ObrotyMies")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienciObrotyMies.Fill(DataSetObrotyMies)
                End If
                sqlConnectionKlienciObrotyMies.Close()
            End If
        End If
    End Sub






    '*************************************************
    '** usuwanie zapisow z katalogu AkcjeSpec
    '*************************************************

    Dim dbSelectSQLAkcjeSpec As String = "SELECT * FROM AkcjeSpec WHERE IDENTKLIENTA='" & oIdentKli & "'"

    Dim dbConnectionAkcjeSpec As OleDb.OleDbConnection
    Dim dbCommandSelectAkcjeSpec As OleDb.OleDbCommand
    Dim dbCommandDeleteAkcjeSpec As OleDb.OleDbCommand
    Dim dbDataAdapterAkcjeSpec As OleDb.OleDbDataAdapter

    Dim sqlConnectionKlienciAkcjeSpec As SqlClient.SqlConnection
    Dim sqlCommandSelectKlienciAkcjeSpec As SqlClient.SqlCommand
    Dim sqlCommandDeleteKlienciAkcjeSpec As SqlClient.SqlCommand
    Dim sqlDataAdapterKlienciAkcjeSpec As SqlClient.SqlDataAdapter

    Dim DataSetAkcjeSpec As New DataSet
    Dim DataViewAkcjeSpec As New DataView

    'Public asIdentAkcji As String            'IDENTAKCJI     Text 20
    'Public asIdentKlienta As String          'IDENTKLI       Text 6
    'Public asWersjaAkcji As Integer           'WERSJA    Integer

    Private Sub UsunAkcjeSpec(ByVal IdKl As String)
        dbSelectSQLAkcjeSpec = "SELECT * FROM AkcjeSpec WHERE IDENTKLI='" & IdKl & "'"

        DataSetAkcjeSpec = New DataSet
        DataSetAkcjeSpec.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            ''OleDbConnection-polaczenie
            'dbConnectionAkcjeSpec = New OleDb.OleDbConnection(sConnectionAkcjeSpec)
            'dbCommandSelectAkcjeSpec = New OleDb.OleDbCommand(dbSelectSQLAkcjeSpec, dbConnectionAkcjeSpec)
            'dbCommandDeleteAkcjeSpec = New OleDb.OleDbCommand(sDeleteSQLAkcjeSpec, dbConnectionAkcjeSpec)
            'dbDataAdapterAkcjeSpec = New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeSpec)
            ''---------------------------------------------
            ''mapujemy domyslna nazwe tabeli
            'Dim MapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
            'MapowanieTabeliAkcjeSpec = dbDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_AkcjeSpec")

            ''komendy DataAdapter
            'Dim objParamAkcjeSpec As OleDb.OleDbParameter
            ''------------------------------------------
            ''komenda DELETE
            'objParamAkcjeSpec = dbCommandDeleteAkcjeSpec.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 20, "IDENTAKCJI")
            'objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
            'objParamAkcjeSpec = dbCommandDeleteAkcjeSpec.Parameters.Add("@orygKlient", OleDb.OleDbType.VarChar, 6, "IDENTKLI")
            'objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
            'objParamAkcjeSpec = dbCommandDeleteAkcjeSpec.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            'objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
            'dbDataAdapterAkcjeSpec.DeleteCommand = dbCommandDeleteAkcjeSpec

            '' Add the RowUpdated event handler.
            'AddHandler dbDataAdapterAkcjeSpec.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            ''------------------------------------------
            ''wypelnij DATASET
            'Try
            '    dbConnectionAkcjeSpec.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '    '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    '    MessageBoxIcon.Information, _
            '    '    MessageBoxDefaultButton.Button1)
            'End Try

            'If dbConnectionAkcjeSpec.State = ConnectionState.Open Then
            '    Try
            '        dbDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_AkcjeSpec")
            '        dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
            '        dbConnectionAkcjeSpec.Close()

            '        'zdefiniuj unikalny klucz wyszukiwania w tabeli
            '        DataSetAkcjeSpec.Tables(0).PrimaryKey = New DataColumn() {DataSetAkcjeSpec.Tables(0).Columns("IDENTAKCJI"), DataSetAkcjeSpec.Tables(0).Columns("IDENTKLI")}

            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
            '        '    System.Windows.Forms.MessageBoxButtons.OK, _
            '        '    MessageBoxIcon.Information, _
            '        '    MessageBoxDefaultButton.Button1)
            '    End Try
            '    '---------------------------------
            'End If
        Else
            sqlConnectionKlienciAkcjeSpec = New SqlClient.SqlConnection(sConnectionAkcjeSpec)
            sqlCommandSelectKlienciAkcjeSpec = New SqlClient.SqlCommand(dbSelectSQLAkcjeSpec, sqlConnectionKlienciAkcjeSpec)
            sqlCommandDeleteKlienciAkcjeSpec = New SqlClient.SqlCommand(sDeleteSQLAkcjeSpec, sqlConnectionKlienciAkcjeSpec)
            sqlDataAdapterKlienciAkcjeSpec = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienciAkcjeSpec)
            '---------------------------------------------
            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
            MapowanieTabeliAkcjeSpec = sqlDataAdapterKlienciAkcjeSpec.TableMappings.Add("Table", "TABELA_AkcjeSpec")

            'komendy DataAdapter
            Dim objParamAkcjeSpec As SqlClient.SqlParameter
            '------------------------------------------
            'komenda DELETE
            objParamAkcjeSpec = sqlCommandDeleteKlienciAkcjeSpec.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 20, "IDENTAKCJI")
            objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
            objParamAkcjeSpec = sqlCommandDeleteKlienciAkcjeSpec.Parameters.Add("@orygKlient", SqlDbType.VarChar, 6, "IDENTKLI")
            objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
            objParamAkcjeSpec = sqlCommandDeleteKlienciAkcjeSpec.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParamAkcjeSpec.SourceVersion = DataRowVersion.Original
            sqlDataAdapterKlienciAkcjeSpec.DeleteCommand = sqlCommandDeleteKlienciAkcjeSpec

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterKlienciAkcjeSpec.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionKlienciAkcjeSpec.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            If sqlConnectionKlienciAkcjeSpec.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterKlienciAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_AkcjeSpec")
                    sqlDataAdapterKlienciAkcjeSpec.Fill(DataSetAkcjeSpec)
                    sqlConnectionKlienciAkcjeSpec.Close()

                    'zdefiniuj unikalny klucz wyszukiwania w tabeli
                    DataSetAkcjeSpec.Tables(0).PrimaryKey = New DataColumn() {DataSetAkcjeSpec.Tables(0).Columns("IDENTAKCJI"), DataSetAkcjeSpec.Tables(0).Columns("IDENTKLI")}

                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    MessageBoxIcon.Information, _
                    '    MessageBoxDefaultButton.Button1)
                End Try
                '---------------------------------
            End If
        End If

        Dim dr As DataRow = Nothing
        Dim Id As String = ""
        Dim i As Integer = 0
        For i = DataSetAkcjeSpec.Tables(0).Rows.Count - 1 To 0 Step -1
            dr = DataSetAkcjeSpec.Tables(0).Rows(i)
            Id = dr.Item("IDENTKLI")
            dr.Delete()
        Next
        AktualizujAkcjeSpec()
    End Sub

    Private Sub AktualizujAkcjeSpec()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionAkcjeSpec.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnectionAkcjeSpec.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapterAkcjeSpec.Update(DataSetAkcjeSpec, "TABELA_AkcjeSpec")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
            '    End If
            '    dbConnectionAkcjeSpec.Close()
            'End If
        Else
            Try
                sqlConnectionKlienciAkcjeSpec.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionKlienciAkcjeSpec.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterKlienciAkcjeSpec.Update(DataSetAkcjeSpec, "TABELA_AkcjeSpec")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterKlienciAkcjeSpec.Fill(DataSetAkcjeSpec)
                End If
                sqlConnectionKlienciAkcjeSpec.Close()
            End If
        End If
    End Sub





































    '===================================
    ' przepisujemy dane adresowe do EXCELa
    '===================================

    '-------------odbc
    'Private EXCELConnection As String = "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & oPlikNaDysku & ";"
    Private EXCELConnectionODBC As String =
                "Driver={Microsoft Excel Driver (*.xls)};" &
                "DSN=Pliki programu Excel;" &
                "DBQ=" & oPlikNaDysku & ";" &
                "DriverId=790;" &
                "MaxBufferSize=2048;" &
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

    Private Sub cmdAdrExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdrExport.Click
        Dim Text As String = ""
        Dim Text2 As String = ""
        Dim Num As Integer = 0

        Dim N As String = ""
        Dim T As String = ""
        Dim R As Integer = 0
        Dim H As String = ""
        Dim W As String = ""
        Dim I As Integer = 0
        '---------------------------------
        If OSArchitectureIs64bit() Then
            ''ACE
            'If you are an application developer using OLEDB, set the Provider argument 
            'of the ConnectionString property to “Microsoft.ACE.OLEDB.12.0”. 
            EXCELConnectionOLeDB = "Provider=""Microsoft.ACE.OLEDB.12.0"";" &
            "Data Source=" & oPlikNaDysku & "; " &
            "Extended Properties='Excel 8.0; HDR=YES; IMEX=1';"

            'If you are an application developer using ODBC to connect to Microsoft Office Access data, 
            'set the Connection String to “Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=path to mdb/accdb file”
            'If you are an application developer using ODBC to connect to Microsoft Office Excel data, 
            'set the Connection String to “Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=path to xls/xlsx/xlsm/xlsb file”
            'Return ("Driver={Microsoft Access Driver (*.mdb, *.accdb)};" & _
            '        "DBQ=""" & ParL_KatZbiorow & "\" & oPlikKartoteki & """ ")
        Else
            'Jet 4,0
            EXCELConnectionOLeDB = "Provider=Microsoft.Jet.OLEDB.4.0; " &
            "Data Source=" & oPlikNaDysku & "; " &
            "Extended Properties='Excel 8.0; HDR=YES; IMEX=1';"
        End If

        Dim app As Microsoft.Office.Interop.Excel.Application
        app = New Microsoft.Office.Interop.Excel.Application()

        If app Is Nothing Then
            MessageBox.Show("Nie mogê uruchomiæ program EXCEL", "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            app.Visible = False             'EXCEL WIDOCZNY NA EKRANIE
            Cursor = Cursors.WaitCursor

            Dim workbooks As Workbooks      'Getting the workbooks collection
            workbooks = app.Workbooks

            Dim workbook As _Workbook       'adding a new workbook
            workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)

            Dim sheets As Sheets            'Getting the worksheets collection
            sheets = workbook.Worksheets

            Dim worksheet As _Worksheet     'adding a new worksheet
            worksheet = sheets.Item(1)
            If worksheet Is Nothing Then
                MessageBox.Show("Nie mogê dodaæ nowego arkusza EXCEL", "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End If
            worksheet.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnRight
            worksheet.Columns.ColumnWidth = 100
            '--------------------------
            'dbCommandUpdate.Parameters.Add("@oNAZWA1Kli", OleDb.OleDbType.varchar, 100, "NAZWAKLIENTA")
            'dbCommandUpdate.Parameters.Add("@oMiejscKli", OleDb.OleDbType.varchar, 50, "MIEJSCOWOSC")
            'dbCommandUpdate.Parameters.Add("@oKodPoczKli", OleDb.OleDbType.varchar, 10, "KODPOCZTOWY")
            'dbCommandUpdate.Parameters.Add("@oUlicaKli", OleDb.OleDbType.varchar, 50, "ULICA")
            'dbCommandUpdate.Parameters.Add("@oNrDomuKli", OleDb.OleDbType.varchar, 10, "NRDOMU")
            'dbCommandUpdate.Parameters.Add("@oOsobaKontaktowaKli", OleDb.OleDbType.varchar, 50, "OSOBAKONTAKTOWA")
            '--------------------------
            If DataViewKlienci.Count > 0 Then
                Dim ir As Integer = 0
                Dim dr As DataRowView
                For ir = 0 To DataViewKlienci.Count - 1
                    dr = DataViewKlienci.Item(ir)

                    N = "NAZWAKLIENTA"
                    R = 60
                    W = HorizontalAlignment.Left
                    Text = Trim(dr.Item(N))
                    WpiszDoArkusza(worksheet, "A" & Trim(Str(ir + 1)), Text, R, W)

                    N = "KODPOCZTOWY"
                    R = 10
                    W = HorizontalAlignment.Left
                    Text = Trim(dr.Item(N))
                    WpiszDoArkusza(worksheet, "B" & Trim(Str(ir + 1)), Text, R, W)

                    N = "MIEJSCOWOSC"
                    R = 30
                    W = HorizontalAlignment.Left
                    Text = Trim(dr.Item(N))
                    WpiszDoArkusza(worksheet, "C" & Trim(Str(ir + 1)), Text, R, W)

                    N = "ULICA"
                    R = 30
                    W = HorizontalAlignment.Left
                    Text = Trim(dr.Item(N))
                    WpiszDoArkusza(worksheet, "D" & Trim(Str(ir + 1)), Text, R, W)

                    N = "NRDOMU"
                    R = 10
                    W = HorizontalAlignment.Left
                    Num = dr.Item("NUMNRDOMU")
                    Text2 = Trim(dr.Item("IDDOMU"))

                    If Num <> 0 Then    'czy jest numer domu
                        Text = "'" & Trim(Str(Num))
                    End If
                    If Len(Text2) > 0 Then  'czy jest ident
                        Text &= " " & Text2
                    End If
                    WpiszDoArkusza(worksheet, "E" & Trim(Str(ir + 1)), Text, R, W)

                    N = "OSOBAKONTAKTOWA"
                    R = 30
                    W = HorizontalAlignment.Left
                    'Text = Trim(dr.Item(N))
                    Text = Trim(OsobaKontaktowa(dr.Item("NRKARTY")))
                    WpiszDoArkusza(worksheet, "F" & Trim(Str(ir + 1)), Text, R, W)

                    N = "EMAIL"
                    R = 30
                    W = HorizontalAlignment.Left
                    Text = Trim(dr.Item(N))
                    WpiszDoArkusza(worksheet, "G" & Trim(Str(ir + 1)), Text, R, W)

                Next
            End If

            'wybierz góre arkusza...
            Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
            range = worksheet.Range("A1", Missing.Value)
            range.Select()

            Cursor = Cursors.Default
            app.Visible = True      'EXCEL WIDOCZNY NA EKRANIE
            'Try
            '    ' If user interacted with Excel it will not close when the app object is destroyed, so we close it explicitely
            '    workbook.Saved = True
            '    app.UserControl = False
            '    app.Quit()
            'Catch Outer As COMException
            '    'Console.WriteLine("User closed Excel manually, so we don't have to do that")
            'End Try
        End If
    End Sub





    Private Sub WpiszDoArkusza(ByVal Arkusz As _Worksheet,
                               ByVal Adres As String,
                               ByVal Zawartosc As String,
                               ByVal Szerokosc As Integer,
                               ByVal Wyrownanie As String)
        Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
        Dim ColRange As Microsoft.Office.Interop.Excel.Range
        'Dim RowRange As Microsoft.Office.Interop.Excel.Range

        range = Arkusz.Range(Adres, Missing.Value)
        If range Is Nothing Then
            MessageBox.Show("Nie mogê zdefiniowaæ zakresu arkusza EXCEL", "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            range.Value2 = Zawartosc
            range.ColumnWidth = Szerokosc
            range.RowHeight = 15
            range.WrapText = True
            'range.Columns.HorizontalAlignment = HorizontalAlignment.Left

            'dotyczy calej kolumny
            ColRange = range.Columns
            'Select Case Wyrownanie
            '    Case "Do lewej"
            '        ColRange.HorizontalAlignment = HorizontalAlignment.Left
            '    Case "Do prawej"
            '        ColRange.HorizontalAlignment = HorizontalAlignment.Right
            '    Case Else
            '        ColRange.HorizontalAlignment = HorizontalAlignment.Center
            'End Select

            'RowRange = range.Rows
            'RowRange.VerticalAlignment

        End If

    End Sub













    '---------------------------------
    ' Zmiana ustawienia Markera
    '--------------------------------

    Private Sub ZmianaWskAkt()
        If DataSetKlienci.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzKlienci()

            If oAktywnyKli Then
                If MessageBox.Show("Oznaczenie klienta " & oIdentKli & " jako NIEAKTYWNY spowoduje" &
                                "¿e kartoteka tego klienta NIE BÊDZIE BRANA do analiz i sumowañ" &
                                "realizowanych przez program..." & vbCrLf & vbCrLf &
                                "Czy na pewno mam oznaczyæ tego klienta jako NIEAKTYWNY ?",
                                "Uwaga :",
                                System.Windows.Forms.MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question,
                                MessageBoxDefaultButton.Button1) = DialogResult.Yes Then

                    oAktywnyKli = False
                Else
                    'oznakowanie bez zmian
                End If
            Else
                'lest nieaktywny - przywróæ
                oAktywnyKli = True
            End If

            Dim foundRow As DataRow
            ' Create an array for the key values to find.
            Dim findTheseVals(0) As Object
            findTheseVals(0) = oIdentKli
            foundRow = DataSetKlienci.Tables(0).Rows.Find(findTheseVals)
            ' sprawdz czy jest...
            If Not (foundRow Is Nothing) Then
                foundRow("AKTYWNY") = oAktywnyKli
                foundRow("WERSJA") = ZmienWersje(oWersjaKli) 'Wersja         Integer, 2
                AktualizujBaze()
                dagKlienci.Update()

                OdswiezBaze()
            End If
        End If
    End Sub



    '---------------------------------
    ' Zmiana ustawienia Markera
    '--------------------------------

    Private Sub ZmianaUstawieniaMarkera(ByVal UstawienieMarkera As String)      'TAK / NIE / ""=zmien
        If DataSetKlienci.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzKlienci()
            Select Case UstawienieMarkera
                Case "TAK"
                    oMarkerKli = True
                Case "NIE"
                    oMarkerKli = False
                Case Else
                    oMarkerKli = Not oMarkerKli
            End Select

            Dim foundRow As DataRow
            ' Create an array for the key values to find.
            Dim findTheseVals(0) As Object
            findTheseVals(0) = oIdentKli
            foundRow = DataSetKlienci.Tables(0).Rows.Find(findTheseVals)
            ' sprawdz czy jest...
            If Not (foundRow Is Nothing) Then
                foundRow("MARKER") = oMarkerKli
                foundRow("WERSJA") = ZmienWersje(oWersjaKli) 'Wersja         Integer, 2
                AktualizujBaze()
                dagKlienci.Update()
                'OdswiezBaze()
            End If
        End If
    End Sub




    Private Sub ZmianaUstawieniaMarkeraP(ByVal UstawienieMarkera As String)      'TAK / NIE / ""=zmien
        If DataSetKlienci.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzKlienci()
            Select Case UstawienieMarkera
                Case "TAK"
                    oMarkerPKli = True
                Case "NIE"
                    oMarkerPKli = False
                Case Else
                    oMarkerPKli = Not oMarkerPKli
            End Select

            Dim foundRow As DataRow
            ' Create an array for the key values to find.
            Dim findTheseVals(0) As Object
            findTheseVals(0) = oIdentKli
            foundRow = DataSetKlienci.Tables(0).Rows.Find(findTheseVals)
            ' sprawdz czy jest...
            If Not (foundRow Is Nothing) Then
                foundRow("MARKERP") = oMarkerPKli
                foundRow("WERSJA") = ZmienWersje(oWersjaKli) 'Wersja         Integer, 2
                AktualizujBaze()
                dagKlienci.Update()
                'OdswiezBaze()
            End If
        End If
    End Sub



    Private Sub cmdHarmonogramWizyt_Click(sender As Object, e As EventArgs) Handles cmdHarmonogramWizyt.Click
        Dim Form As New KatalogHarmonogramWizyt
        Form.Show()
    End Sub



    Private Sub cmdDoExcela_Click(sender As Object, e As EventArgs) Handles cmdDoExcela.Click
        'Me.cmdDoExcela.Image = My.Resources.Excel_16
        Me.Cursor = Cursors.WaitCursor
        ExportDataGrid2Excel(dagKlienci, DataViewKlienci, Me.Text)
        Me.Cursor = Cursors.Default
    End Sub


    Private Sub cmdOdswiez_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOdswiez.Click
        OdswiezBaze()
        dagKlienci.Update()
    End Sub

























    '*************************************************
    '** obs³ug amenu kontekstowego
    '*************************************************

    Private Sub menuEKlienci_Opening(sender As Object, e As CancelEventArgs) Handles menuEKlienci.Opening
        If DataSetKlienci.Tables(0).Rows.Count <= 0 Then
            menuZmienWskAkt.Visible = False
        Else
            menuZmienWskAkt.Visible = True
            oOperacja = "EDYTUJ"
            PobierzKlienci()
            If oAktywnyKli Then
                menuZmienWskAkt.Text = "Oznacz klienta " & oIdentKli & " jako NIEAKTYWNY"
            Else
                menuZmienWskAkt.Text = "Oznacz klienta " & oIdentKli & " jako AKTYWNY"
            End If
        End If
    End Sub



    Private Sub menuEEdytuj_Click(sender As Object, e As EventArgs) Handles menuEEdytuj.Click
        cmdEdytuj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuEDodaj_Click(sender As Object, e As EventArgs) Handles menuEDodaj.Click
        cmdDodaj_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuEUsun_Click(sender As Object, e As EventArgs) Handles menuEUsun.Click
        cmdUsuñ_Click(Me, System.EventArgs.Empty)
    End Sub

    Private Sub menuESingleL_Click(sender As Object, e As EventArgs) Handles menuESingleL.Click
        dagKlienci.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagKlienci.Refresh()
    End Sub

    Private Sub menuEMultiL_Click(sender As Object, e As EventArgs) Handles menuEMultiL.Click
        dagKlienci.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagKlienci.Refresh()
    End Sub

    Private Sub menuEOdswiez_Click(sender As Object, e As EventArgs) Handles menuEOdswiez.Click
        OdswiezBaze()
        dagKlienci.Update()
    End Sub





    Private Sub menuENowaA_Click(sender As Object, e As EventArgs) Handles menuENowaA.Click
        Dim Form As New AkcjaMarketingowaNowa(DataViewKlienci)
        Form.ShowDialog()
    End Sub

    Private Sub menuEWybierzA_Click(sender As Object, e As EventArgs) Handles menuEWybierzA.Click
        oWybranoAkcjeMarketingowa = False
        'definiuj delegat
        Dim deleg As delegateAktualizuj = AddressOf AktualizujBaze
        Dim Form As New AkcjaMarketingowaWybierz(DataSetKlienci, deleg)
        Form.ShowDialog()

        OdswiezBaze()
        If oWybranoAkcjeMarketingowa Then
            oLokalnyFiltr = ""
            oNazwaSchematu = ""
            oSchematFiltrowania = ""
            stbKlienci.Panels(4).Text = "Szablon : "
            'dagKlienci.Update()

            'aktualizuj Katalog klientow
            CzyscFiltryDGV(Me.Panel4,
                        dagKlienci,
                        DataViewKlienci.Table.Columns.Count,
                        DataViewKlienci,
                        stbKlienci)

            'CzyscFiltry(Me.Panel4,
            '            dagKlienci,
            '            DataViewKlienci.Table.Columns.Count,
            '            DataViewKlienci,
            '            stbKlienci,
            '            Trim(oSchematFiltrowania))

            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Marker  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Marker P.  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "Aktywny  ", 50, True)

            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Panel4.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi052" Or ctrl.Name = "txtFi053" Then
                        ctrl.Text = "TRUE"
                        LiMark += 1
                        If LiMark = 2 Then
                            Exit For
                        End If
                    End If
                End If
            Next
            dagKlienci.Update()
        End If
    End Sub

    Private Sub menuEWybierzKilkaA_Click(sender As Object, e As EventArgs) Handles menuEWybierzKilkaA.Click
        oWybranoAkcjeMarketingowa = False
        'definiuj delegat
        Dim deleg As delegateAktualizuj = AddressOf AktualizujBaze
        Dim Form As New AkcjaMarketingowaWybierzKilka(DataSetKlienci, deleg)
        Form.ShowDialog()

        OdswiezBaze()
        If oWybranoAkcjeMarketingowa Then
            oLokalnyFiltr = ""
            oNazwaSchematu = ""
            oSchematFiltrowania = ""
            stbKlienci.Panels(4).Text = "Szablon : "
            'dagKlienci.Update()

            'aktualizuj Katalog klientow
            CzyscFiltryDGV(Me.Panel4,
                        dagKlienci,
                        DataViewKlienci.Table.Columns.Count,
                        DataViewKlienci,
                        stbKlienci)

            'CzyscFiltry(Me.Panel4,
            '            dagKlienci,
            '            DataViewKlienci.Table.Columns.Count,
            '            DataViewKlienci,
            '            stbKlienci,
            '            Trim(oSchematFiltrowania))

            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Marker  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Marker P.  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "Aktywny  ", 50, True)

            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Panel4.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi052" Or ctrl.Name = "txtFi053" Then
                        ctrl.Text = "TRUE"
                        LiMark += 1
                        If LiMark = 2 Then
                            Exit For
                        End If
                    End If
                End If
            Next
            dagKlienci.Update()
        End If
    End Sub

    Private Sub menuEDodajDoA_Click(sender As Object, e As EventArgs) Handles menuEDodajDoA.Click
        Dim Form As New AkcjaMarketingowaDodajDo(DataViewKlienci)
        Form.ShowDialog()
    End Sub

    Private Sub menuEUsunZA_Click(sender As Object, e As EventArgs) Handles menuEUsunZA.Click
        Dim Form As New AkcjaMarketingowaUsunZ(DataViewKlienci)
        Form.ShowDialog()
    End Sub

    Private Sub menuEWybierzWgSprzed_Click(sender As Object, e As EventArgs) Handles menuEWybierzWgSprzed.Click
        oWybranoAkcjeMarketingowa = False
        'definiuj delegat
        Dim deleg As delegateAktualizuj = AddressOf AktualizujBaze
        Dim Form As New AkcjaMarketingowaWybierzWgSprzedazy(DataSetKlienci, deleg)
        Form.ShowDialog()

        If oWybranoAkcjeMarketingowa Then
            Me.Cursor = Cursors.WaitCursor
            oLokalnyFiltr = ""
            oNazwaSchematu = ""
            oSchematFiltrowania = ""
            stbKlienci.Panels(4).Text = "Szablon : "
            OdswiezBaze()
            'dagKlienci.Update()

            'aktualizuj Katalog klientow
            CzyscFiltryDGV(Me.Panel4,
                        dagKlienci,
                        DataViewKlienci.Table.Columns.Count,
                        DataViewKlienci,
                        stbKlienci)

            'CzyscFiltry(Me.Panel4,
            '            dagKlienci,
            '            DataViewKlienci.Table.Columns.Count,
            '            DataViewKlienci,
            '            stbKlienci,
            '            Trim(oSchematFiltrowania))

            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Marker  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Marker P.  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "Aktywny  ", 50, True)

            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Panel4.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi052" Or ctrl.Name = "txtFi053" Then
                        ctrl.Text = "TRUE"
                        LiMark += 1
                        If LiMark = 2 Then
                            Exit For
                        End If
                    End If
                End If
            Next
            dagKlienci.Update()
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub menuEWybierzWgOsoby_Click(sender As Object, e As EventArgs) Handles menuEWybierzWgOsoby.Click
        oWybranoAkcjeMarketingowa = False
        'definiuj delegat
        Dim deleg As delegateAktualizuj = AddressOf AktualizujBaze
        Dim Form As New AkcjaMarketingowaWybierzWgOsobKontaktowych(DataSetKlienci, deleg)
        Form.ShowDialog()

        If oWybranoAkcjeMarketingowa Then
            Me.Cursor = Cursors.WaitCursor
            oLokalnyFiltr = ""
            oNazwaSchematu = ""
            oSchematFiltrowania = ""
            stbKlienci.Panels(4).Text = "Szablon : "
            OdswiezBaze()
            'dagKlienci.Update()

            'aktualizuj Katalog klientow
            CzyscFiltryDGV(Me.Panel4,
                        dagKlienci,
                        DataViewKlienci.Table.Columns.Count,
                        DataViewKlienci,
                        stbKlienci)

            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Marker  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Marker P.  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "Aktywny  ", 50, True)

            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Panel4.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi052" Or ctrl.Name = "txtFi053" Then
                        ctrl.Text = "TRUE"
                        LiMark += 1
                        If LiMark = 2 Then
                            Exit For
                        End If
                    End If
                End If
            Next
            dagKlienci.Update()
            Me.Cursor = Cursors.Default
        End If
    End Sub



    Private Sub menuEWybierzWgRZU_Click(sender As Object, e As EventArgs) Handles menuEWybierzWgRZU.Click
        oWybranoAkcjeMarketingowa = False
        'definiuj delegat
        Dim deleg As delegateAktualizuj = AddressOf AktualizujBaze
        Dim Form As New AkcjaMarketingowaWybierzWgUslug(DataSetKlienci, deleg)
        Form.ShowDialog()

        If oWybranoAkcjeMarketingowa Then
            Me.Cursor = Cursors.WaitCursor
            oLokalnyFiltr = ""
            oNazwaSchematu = ""
            oSchematFiltrowania = ""
            stbKlienci.Panels(4).Text = "Szablon : "
            OdswiezBaze()
            'dagKlienci.Update()

            'aktualizuj Katalog klientow
            CzyscFiltryDGV(Me.Panel4,
                        dagKlienci,
                        DataViewKlienci.Table.Columns.Count,
                        DataViewKlienci,
                        stbKlienci)

            'CzyscFiltry(Me.Panel4,
            '            dagKlienci,
            '            DataViewKlienci.Table.Columns.Count,
            '            DataViewKlienci,
            '            stbKlienci,
            '            Trim(oSchematFiltrowania))

            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Marker  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Marker P.  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "Aktywny  ", 50, True)

            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Panel4.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi052" Or ctrl.Name = "txtFi053" Then
                        ctrl.Text = "TRUE"
                        LiMark += 1
                        If LiMark = 2 Then
                            Exit For
                        End If
                    End If
                End If
            Next
            dagKlienci.Update()
            dagKlienci_CurrentCellChanged(sender, e)
            Me.Cursor = Cursors.Default
        End If
    End Sub





    Private Sub menuEWybierzWgWzrostu_Click(sender As Object, e As EventArgs) Handles menuEWybierzWgWzrostu.Click
        oWybranoAkcjeMarketingowa = False
        'definiuj delegat
        Dim deleg As delegateAktualizuj = AddressOf AktualizujBaze


        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne")
        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Iloœci")
        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne")
        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Wartoœci")
        'cbxCoAnalizowac.Items.Add("Analizuj wartoœæ Mar¿y")
        'cbxCoAnalizowac.Items.Add("Analizuj procent Mar¿y")

        Dim Form As New AkcjaMarketingowaWybierzWgObrotow(DataSetKlienci, "Analizuj Wartoœciowo-Sumuj Wartoœci", deleg)
        Form.ShowDialog()

        If oWybranoAkcjeMarketingowa Then
            Me.Cursor = Cursors.WaitCursor
            oLokalnyFiltr = ""
            oNazwaSchematu = ""
            oSchematFiltrowania = ""
            stbKlienci.Panels(4).Text = "Szablon : "
            OdswiezBaze()
            'dagKlienci.Update()

            'aktualizuj Katalog klientow
            CzyscFiltryDGV(Me.Panel4,
                        dagKlienci,
                        DataViewKlienci.Table.Columns.Count,
                        DataViewKlienci,
                        stbKlienci)

            'CzyscFiltry(Me.Panel4,
            '            dagKlienci,
            '            DataViewKlienci.Table.Columns.Count,
            '            DataViewKlienci,
            '            stbKlienci,
            '            Trim(oSchematFiltrowania))

            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, "Marker  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, "Marker P.  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, "Aktywny  ", 50, True)

            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Panel4.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi052" Or ctrl.Name = "txtFi053" Then
                        ctrl.Text = "TRUE"
                        LiMark += 1
                        If LiMark = 2 Then
                            Exit For
                        End If
                    End If
                End If
            Next
            dagKlienci.Update()
            dagKlienci_CurrentCellChanged(sender, e)
            Me.Cursor = Cursors.Default
        End If
    End Sub



    Private Sub menuZmienWskAkt_Click(sender As Object, e As EventArgs) Handles menuZmienWskAkt.Click
        ZmianaWskAkt()
    End Sub


    Private Sub menuEUstawM_Click(sender As Object, e As EventArgs) Handles menuEUstawM.Click
        ZmianaUstawieniaMarkera("TAK")
    End Sub

    Private Sub menuEUsunM_Click(sender As Object, e As EventArgs) Handles menuEUsunM.Click
        ZmianaUstawieniaMarkera("NIE")
    End Sub

    Private Sub menuEZmienM_Click(sender As Object, e As EventArgs) Handles menuEZmienM.Click
        ZmianaUstawieniaMarkera("")
    End Sub

    Private Sub menuEUstawMP_Click(sender As Object, e As EventArgs) Handles menuEUstawMP.Click
        ZmianaUstawieniaMarkeraP("TAK")
    End Sub

    Private Sub menuEUsunMP_Click(sender As Object, e As EventArgs) Handles menuEUsunMP.Click
        ZmianaUstawieniaMarkeraP("NIE")
    End Sub

    Private Sub menuEZmienMP_Click(sender As Object, e As EventArgs) Handles menuEZmienMP.Click
        ZmianaUstawieniaMarkeraP("")
    End Sub

















    '**************************************************************
    '** Zapamietaj szerokosci kolumn
    '**************************************************************

    Public Sub ZapamietajSzerokosciKolumn()
        Dim ii As Integer = 0
        Dim Txt As String = ""
        Dim FileNum As Integer

        For ii = 0 To DataViewKlienci.Table.Columns.Count - 1
            Txt &= Trim(Str(dagKlienci.Columns.Item(ii).Width)) & "|"
        Next

        If IO.File.Exists(oKatParametry & "\" & oPlikSzerokosciKolumnKartoteka) Then
            IO.File.Delete(oKatParametry & "\" & oPlikSzerokosciKolumnKartoteka)
        End If

        'ZAPISZ DO PLIKU...
        FileNum = FreeFile()    'kolejny nr otwartego zbioru
        FileOpen(FileNum, oKatParametry & "\" & oPlikSzerokosciKolumnKartoteka, OpenMode.Output)
        PrintLine(FileNum, "KartotekaKlientowPryzmat" & oAktualnaWersjaBazyDanych.ToString)
        PrintLine(FileNum, Txt)
        FileClose(FileNum)
    End Sub

    Public Sub UstalSzerokosciKolumn()
        Dim ii As Integer = 0
        Dim Txt As String = ""
        Dim FileNum As Integer
        Dim pos As Integer = 0
        Dim NrWersjiPliku As Integer = 0
        Dim NowyTXT As String = ""

        If IO.File.Exists(oKatParametry & "\" & oPlikSzerokosciKolumnKartoteka) Then
            FileNum = FreeFile()    'kolejny nr otwartego zbioru
            FileOpen(FileNum, oKatParametry & "\" & oPlikSzerokosciKolumnKartoteka, OpenMode.Input)
            If Not EOF(FileNum) Then
                Txt = LineInput(FileNum)

                '-------------obsluga zmiany pliku
                If InStr(Txt, "KartotekaKlientowPryzmat") = 1 Then
                    NrWersjiPliku = CInt(Mid(Txt, 25, 3))
                    Txt = LineInput(FileNum)
                Else
                    'ident wersji w pliku od wersji 2.40
                    NrWersjiPliku = 230
                End If

                'sprawdzenie koniecznosci korekty
                If NrWersjiPliku <> oAktualnaWersjaBazyDanych Then
                    'pozostawiamy domyœlne...
                Else
                    'pobieramy z pliku
                    For ii = 0 To DataViewKlienci.Table.Columns.Count - 1
                        pos = InStr(Txt, "|")
                        If pos > 0 Then
                            dagKlienci.Columns.Item(ii).Width = CInt(Mid(Txt, 1, pos - 1))
                            Txt = Mid(Txt, pos + 1)
                        End If
                    Next
                End If
                '---------------------------------
            End If
            FileClose(FileNum)
        End If
    End Sub


    Private Sub clrClearColWidth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearColWidth.Click
        Dim ii As Integer = 0
        StartFormy = True
        Me.Cursor = Cursors.WaitCursor
        For ii = 0 To DataViewKlienci.Table.Columns.Count - 1
            'dagKlienci.Columns.Item(ii).Width = GetIntFromText(dagKlienci.Columns.Item(ii).Tag)
            If IsNumeric(dagKlienci.Columns.Item(ii).Tag) Then
                dagKlienci.Columns.Item(ii).Width = dagKlienci.Columns.Item(ii).Tag
            Else
                dagKlienci.Columns.Item(ii).Width = 0
            End If
        Next
        StartFormy = False
        dagKlienci_CurrentCellChanged(sender, e)
        InicjujFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        Me.Cursor = Cursors.Default
        ''--------------------------------
        ZapamietajSzerokosciKolumn()
    End Sub

End Class
