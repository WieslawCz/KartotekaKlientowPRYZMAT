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
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel

Public Class KatalogKlientowAnaliza
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
    Friend WithEvents lblNrDomu As System.Windows.Forms.Label
    Friend WithEvents lblTelefon As System.Windows.Forms.Label
    Friend WithEvents lblUlica As System.Windows.Forms.Label
    Friend WithEvents lblNazwaHandlowa As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdWybierzSchemat As System.Windows.Forms.Button
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents lblRodzaj As System.Windows.Forms.Label
    Friend WithEvents cmdWysylaj As System.Windows.Forms.Button
    Friend WithEvents lblTransakcje As System.Windows.Forms.Label
    Friend WithEvents lblNastKontakt As System.Windows.Forms.Label
    Friend WithEvents lblOstKontakt As System.Windows.Forms.Label
    Friend WithEvents lblOpiekun As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents lblPotencjal As System.Windows.Forms.Label
    Friend WithEvents stbFiltry As System.Windows.Forms.StatusBar
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmdFi As System.Windows.Forms.Button
    Friend WithEvents txtTOFI As System.Windows.Forms.TextBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents cmdOdswiez As System.Windows.Forms.Button
    Friend WithEvents lblIdDomu As System.Windows.Forms.Label
    Friend WithEvents cmdClearColWidth As System.Windows.Forms.Button
    Friend WithEvents HorizScrollTimer As System.Windows.Forms.Timer
    Friend WithEvents cbxCoAnalizowac As System.Windows.Forms.ComboBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Chk212 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk211 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk210 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk209 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk208 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk207 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk206 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk205 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk204 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk203 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk202 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk201 As System.Windows.Forms.CheckBox
    Friend WithEvents lblOkres2 As System.Windows.Forms.Label
    Friend WithEvents Chk112 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk111 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk110 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk109 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk108 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk107 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk106 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk105 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk104 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk103 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk102 As System.Windows.Forms.CheckBox
    Friend WithEvents Chk101 As System.Windows.Forms.CheckBox
    Friend WithEvents lblOkres1 As System.Windows.Forms.Label
    Friend WithEvents cmdPrzeliczSprzedaz As System.Windows.Forms.Button
    Friend WithEvents cmdHarmonogramWizyt As System.Windows.Forms.Button
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
    Friend WithEvents TabSumy As TabControl
    Friend WithEvents TabSprzedaz As TabPage
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents valM31 As System.Windows.Forms.Label
    Friend WithEvents valM25 As System.Windows.Forms.Label
    Friend WithEvents lblM25 As System.Windows.Forms.Label
    Friend WithEvents lblM31 As System.Windows.Forms.Label
    Friend WithEvents valM24 As System.Windows.Forms.Label
    Friend WithEvents lblM24 As System.Windows.Forms.Label
    Friend WithEvents valM30 As System.Windows.Forms.Label
    Friend WithEvents valM23 As System.Windows.Forms.Label
    Friend WithEvents lblM23 As System.Windows.Forms.Label
    Friend WithEvents lblM30 As System.Windows.Forms.Label
    Friend WithEvents valM22 As System.Windows.Forms.Label
    Friend WithEvents lblM22 As System.Windows.Forms.Label
    Friend WithEvents valM29 As System.Windows.Forms.Label
    Friend WithEvents valM21 As System.Windows.Forms.Label
    Friend WithEvents lblM21 As System.Windows.Forms.Label
    Friend WithEvents lblM29 As System.Windows.Forms.Label
    Friend WithEvents valM20 As System.Windows.Forms.Label
    Friend WithEvents lblM20 As System.Windows.Forms.Label
    Friend WithEvents valM28 As System.Windows.Forms.Label
    Friend WithEvents valM11 As System.Windows.Forms.Label
    Friend WithEvents lblM11 As System.Windows.Forms.Label
    Friend WithEvents lblM28 As System.Windows.Forms.Label
    Friend WithEvents valM10 As System.Windows.Forms.Label
    Friend WithEvents lblM10 As System.Windows.Forms.Label
    Friend WithEvents valM27 As System.Windows.Forms.Label
    Friend WithEvents valM09 As System.Windows.Forms.Label
    Friend WithEvents lblM09 As System.Windows.Forms.Label
    Friend WithEvents lblM27 As System.Windows.Forms.Label
    Friend WithEvents valM08 As System.Windows.Forms.Label
    Friend WithEvents lblM08 As System.Windows.Forms.Label
    Friend WithEvents valM26 As System.Windows.Forms.Label
    Friend WithEvents lblM26 As System.Windows.Forms.Label
    Friend WithEvents valM07 As System.Windows.Forms.Label
    Friend WithEvents lblM07 As System.Windows.Forms.Label
    Friend WithEvents valM06 As System.Windows.Forms.Label
    Friend WithEvents lblM06 As System.Windows.Forms.Label
    Friend WithEvents valM05 As System.Windows.Forms.Label
    Friend WithEvents lblM05 As System.Windows.Forms.Label
    Friend WithEvents valM04 As System.Windows.Forms.Label
    Friend WithEvents lblM04 As System.Windows.Forms.Label
    Friend WithEvents valM03 As System.Windows.Forms.Label
    Friend WithEvents lblM03 As System.Windows.Forms.Label
    Friend WithEvents valM02 As System.Windows.Forms.Label
    Friend WithEvents lblM02 As System.Windows.Forms.Label
    Friend WithEvents valM01 As System.Windows.Forms.Label
    Friend WithEvents lblM01 As System.Windows.Forms.Label
    Friend WithEvents valM00 As System.Windows.Forms.Label
    Friend WithEvents lblM00 As System.Windows.Forms.Label
    Friend WithEvents val1223 As System.Windows.Forms.Label
    Friend WithEvents lbl1223 As System.Windows.Forms.Label
    Friend WithEvents val0011 As System.Windows.Forms.Label
    Friend WithEvents lbl0011 As System.Windows.Forms.Label
    Friend WithEvents cmdOdznacz3 As System.Windows.Forms.Button
    Friend WithEvents cmdZaznacz3 As System.Windows.Forms.Button
    Friend WithEvents cmdOdznacz2 As System.Windows.Forms.Button
    Friend WithEvents cmdZaznacz2 As System.Windows.Forms.Button
    Friend WithEvents Panel4 As Panel
    Friend WithEvents lblPlanD As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents lblPlanK As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents ToolStripSeparator7 As ToolStripSeparator
    Friend WithEvents menuZmienWskAkt As ToolStripMenuItem
    Friend WithEvents menuEWybierzWgOsoby As ToolStripMenuItem
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents chkOrygTonery As System.Windows.Forms.CheckBox
    Friend WithEvents chkOrygAtrament As System.Windows.Forms.CheckBox
    Friend WithEvents chkPryzmatTonery As System.Windows.Forms.CheckBox
    Friend WithEvents chkPryzmatAtrament As System.Windows.Forms.CheckBox
    Friend WithEvents chkStrony As System.Windows.Forms.CheckBox
    Friend WithEvents chkNajem As System.Windows.Forms.CheckBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents chkSkup As System.Windows.Forms.CheckBox
    Friend WithEvents chkDrukarki As System.Windows.Forms.CheckBox
    Friend WithEvents cmdWszystko As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(KatalogKlientowAnaliza))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.stbKlienci = New System.Windows.Forms.StatusBar()
        Me.cmdEdytuj = New System.Windows.Forms.Button()
        Me.cmdUsuñ = New System.Windows.Forms.Button()
        Me.cmdDodaj = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TabSumy = New System.Windows.Forms.TabControl()
        Me.TabSprzedaz = New System.Windows.Forms.TabPage()
        Me.lblPlanD = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.lblPlanK = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.valM31 = New System.Windows.Forms.Label()
        Me.valM25 = New System.Windows.Forms.Label()
        Me.lblM25 = New System.Windows.Forms.Label()
        Me.lblM31 = New System.Windows.Forms.Label()
        Me.valM24 = New System.Windows.Forms.Label()
        Me.lblM24 = New System.Windows.Forms.Label()
        Me.valM30 = New System.Windows.Forms.Label()
        Me.valM23 = New System.Windows.Forms.Label()
        Me.lblM23 = New System.Windows.Forms.Label()
        Me.lblM30 = New System.Windows.Forms.Label()
        Me.valM22 = New System.Windows.Forms.Label()
        Me.lblM22 = New System.Windows.Forms.Label()
        Me.valM29 = New System.Windows.Forms.Label()
        Me.valM21 = New System.Windows.Forms.Label()
        Me.lblM21 = New System.Windows.Forms.Label()
        Me.lblM29 = New System.Windows.Forms.Label()
        Me.valM20 = New System.Windows.Forms.Label()
        Me.lblM20 = New System.Windows.Forms.Label()
        Me.valM28 = New System.Windows.Forms.Label()
        Me.valM11 = New System.Windows.Forms.Label()
        Me.lblM11 = New System.Windows.Forms.Label()
        Me.lblM28 = New System.Windows.Forms.Label()
        Me.valM10 = New System.Windows.Forms.Label()
        Me.lblM10 = New System.Windows.Forms.Label()
        Me.valM27 = New System.Windows.Forms.Label()
        Me.valM09 = New System.Windows.Forms.Label()
        Me.lblM09 = New System.Windows.Forms.Label()
        Me.lblM27 = New System.Windows.Forms.Label()
        Me.valM08 = New System.Windows.Forms.Label()
        Me.lblM08 = New System.Windows.Forms.Label()
        Me.valM26 = New System.Windows.Forms.Label()
        Me.lblM26 = New System.Windows.Forms.Label()
        Me.valM07 = New System.Windows.Forms.Label()
        Me.lblM07 = New System.Windows.Forms.Label()
        Me.valM06 = New System.Windows.Forms.Label()
        Me.lblM06 = New System.Windows.Forms.Label()
        Me.valM05 = New System.Windows.Forms.Label()
        Me.lblM05 = New System.Windows.Forms.Label()
        Me.valM04 = New System.Windows.Forms.Label()
        Me.lblM04 = New System.Windows.Forms.Label()
        Me.valM03 = New System.Windows.Forms.Label()
        Me.lblM03 = New System.Windows.Forms.Label()
        Me.valM02 = New System.Windows.Forms.Label()
        Me.lblM02 = New System.Windows.Forms.Label()
        Me.valM01 = New System.Windows.Forms.Label()
        Me.lblM01 = New System.Windows.Forms.Label()
        Me.valM00 = New System.Windows.Forms.Label()
        Me.lblM00 = New System.Windows.Forms.Label()
        Me.val1223 = New System.Windows.Forms.Label()
        Me.lbl1223 = New System.Windows.Forms.Label()
        Me.val0011 = New System.Windows.Forms.Label()
        Me.lbl0011 = New System.Windows.Forms.Label()
        Me.lbleMail = New System.Windows.Forms.Label()
        Me.lblIdDomu = New System.Windows.Forms.Label()
        Me.txtTOFI = New System.Windows.Forms.TextBox()
        Me.lblTelefon = New System.Windows.Forms.Label()
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
        Me.cmdWybierzSchemat = New System.Windows.Forms.Button()
        Me.cmdWysylaj = New System.Windows.Forms.Button()
        Me.stbFiltry = New System.Windows.Forms.StatusBar()
        Me.cmdWszystko = New System.Windows.Forms.Button()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdOdswiez = New System.Windows.Forms.Button()
        Me.cmdClearColWidth = New System.Windows.Forms.Button()
        Me.cmdPrzeliczSprzedaz = New System.Windows.Forms.Button()
        Me.cmdDoExcela = New System.Windows.Forms.Button()
        Me.cmdOdznacz2 = New System.Windows.Forms.Button()
        Me.cmdZaznacz2 = New System.Windows.Forms.Button()
        Me.cmdOdznacz3 = New System.Windows.Forms.Button()
        Me.cmdZaznacz3 = New System.Windows.Forms.Button()
        Me.HorizScrollTimer = New System.Windows.Forms.Timer(Me.components)
        Me.cbxCoAnalizowac = New System.Windows.Forms.ComboBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.chkSkup = New System.Windows.Forms.CheckBox()
        Me.chkDrukarki = New System.Windows.Forms.CheckBox()
        Me.chkStrony = New System.Windows.Forms.CheckBox()
        Me.chkNajem = New System.Windows.Forms.CheckBox()
        Me.chkOrygTonery = New System.Windows.Forms.CheckBox()
        Me.chkOrygAtrament = New System.Windows.Forms.CheckBox()
        Me.chkPryzmatTonery = New System.Windows.Forms.CheckBox()
        Me.chkPryzmatAtrament = New System.Windows.Forms.CheckBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Chk212 = New System.Windows.Forms.CheckBox()
        Me.Chk211 = New System.Windows.Forms.CheckBox()
        Me.Chk210 = New System.Windows.Forms.CheckBox()
        Me.Chk209 = New System.Windows.Forms.CheckBox()
        Me.Chk208 = New System.Windows.Forms.CheckBox()
        Me.Chk207 = New System.Windows.Forms.CheckBox()
        Me.Chk206 = New System.Windows.Forms.CheckBox()
        Me.Chk205 = New System.Windows.Forms.CheckBox()
        Me.Chk204 = New System.Windows.Forms.CheckBox()
        Me.Chk203 = New System.Windows.Forms.CheckBox()
        Me.Chk202 = New System.Windows.Forms.CheckBox()
        Me.Chk201 = New System.Windows.Forms.CheckBox()
        Me.lblOkres2 = New System.Windows.Forms.Label()
        Me.Chk112 = New System.Windows.Forms.CheckBox()
        Me.Chk111 = New System.Windows.Forms.CheckBox()
        Me.Chk110 = New System.Windows.Forms.CheckBox()
        Me.Chk109 = New System.Windows.Forms.CheckBox()
        Me.Chk108 = New System.Windows.Forms.CheckBox()
        Me.Chk107 = New System.Windows.Forms.CheckBox()
        Me.Chk106 = New System.Windows.Forms.CheckBox()
        Me.Chk105 = New System.Windows.Forms.CheckBox()
        Me.Chk104 = New System.Windows.Forms.CheckBox()
        Me.Chk103 = New System.Windows.Forms.CheckBox()
        Me.Chk102 = New System.Windows.Forms.CheckBox()
        Me.Chk101 = New System.Windows.Forms.CheckBox()
        Me.lblOkres1 = New System.Windows.Forms.Label()
        Me.cmdHarmonogramWizyt = New System.Windows.Forms.Button()
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
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.TabSumy.SuspendLayout()
        Me.TabSprzedaz.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.dagKlienci, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuEKlienci.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.ForeColor = System.Drawing.Color.Black
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(1055, 479)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(80, 24)
        Me.cmdStop.TabIndex = 38
        Me.cmdStop.Text = "Powrót"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdStop, "Koniec dzia³ania funkcji." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Powrót do ekranu z którego wywo³ano tê funkcjê.")
        '
        'stbKlienci
        '
        Me.stbKlienci.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbKlienci.Dock = System.Windows.Forms.DockStyle.None
        Me.stbKlienci.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbKlienci.Location = New System.Drawing.Point(0, 240)
        Me.stbKlienci.Name = "stbKlienci"
        Me.stbKlienci.ShowPanels = True
        Me.stbKlienci.Size = New System.Drawing.Size(1127, 22)
        Me.stbKlienci.TabIndex = 43
        Me.stbKlienci.Text = "StatusBar1"
        '
        'cmdEdytuj
        '
        Me.cmdEdytuj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytuj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytuj.ForeColor = System.Drawing.Color.Black
        Me.cmdEdytuj.Image = CType(resources.GetObject("cmdEdytuj.Image"), System.Drawing.Image)
        Me.cmdEdytuj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytuj.Location = New System.Drawing.Point(785, 479)
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
        Me.cmdUsuñ.Location = New System.Drawing.Point(700, 479)
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
        Me.cmdDodaj.Location = New System.Drawing.Point(613, 479)
        Me.cmdDodaj.Name = "cmdDodaj"
        Me.cmdDodaj.Size = New System.Drawing.Size(80, 24)
        Me.cmdDodaj.TabIndex = 40
        Me.cmdDodaj.Text = "Dodaj"
        Me.cmdDodaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdDodaj, "Dopisanie (dodanie) nowego klienta do Kartoteki.")
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.Location = New System.Drawing.Point(10, 478)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(104, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 39
        Me.PictureBox1.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.AutoScroll = True
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.TabSumy)
        Me.Panel1.Controls.Add(Me.lbleMail)
        Me.Panel1.Controls.Add(Me.lblIdDomu)
        Me.Panel1.Controls.Add(Me.txtTOFI)
        Me.Panel1.Controls.Add(Me.lblTelefon)
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
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1127, 130)
        Me.Panel1.TabIndex = 46
        '
        'TabSumy
        '
        Me.TabSumy.Alignment = System.Windows.Forms.TabAlignment.Left
        Me.TabSumy.Controls.Add(Me.TabSprzedaz)
        Me.TabSumy.ItemSize = New System.Drawing.Size(32, 20)
        Me.TabSumy.Location = New System.Drawing.Point(640, -2)
        Me.TabSumy.Multiline = True
        Me.TabSumy.Name = "TabSumy"
        Me.TabSumy.Padding = New System.Drawing.Point(3, 3)
        Me.TabSumy.SelectedIndex = 0
        Me.TabSumy.Size = New System.Drawing.Size(747, 110)
        Me.TabSumy.TabIndex = 93
        '
        'TabSprzedaz
        '
        Me.TabSprzedaz.BackColor = System.Drawing.SystemColors.Control
        Me.TabSprzedaz.Controls.Add(Me.lblPlanD)
        Me.TabSprzedaz.Controls.Add(Me.Label9)
        Me.TabSprzedaz.Controls.Add(Me.lblPlanK)
        Me.TabSprzedaz.Controls.Add(Me.Label17)
        Me.TabSprzedaz.Controls.Add(Me.Label18)
        Me.TabSprzedaz.Controls.Add(Me.valM31)
        Me.TabSprzedaz.Controls.Add(Me.valM25)
        Me.TabSprzedaz.Controls.Add(Me.lblM25)
        Me.TabSprzedaz.Controls.Add(Me.lblM31)
        Me.TabSprzedaz.Controls.Add(Me.valM24)
        Me.TabSprzedaz.Controls.Add(Me.lblM24)
        Me.TabSprzedaz.Controls.Add(Me.valM30)
        Me.TabSprzedaz.Controls.Add(Me.valM23)
        Me.TabSprzedaz.Controls.Add(Me.lblM23)
        Me.TabSprzedaz.Controls.Add(Me.lblM30)
        Me.TabSprzedaz.Controls.Add(Me.valM22)
        Me.TabSprzedaz.Controls.Add(Me.lblM22)
        Me.TabSprzedaz.Controls.Add(Me.valM29)
        Me.TabSprzedaz.Controls.Add(Me.valM21)
        Me.TabSprzedaz.Controls.Add(Me.lblM21)
        Me.TabSprzedaz.Controls.Add(Me.lblM29)
        Me.TabSprzedaz.Controls.Add(Me.valM20)
        Me.TabSprzedaz.Controls.Add(Me.lblM20)
        Me.TabSprzedaz.Controls.Add(Me.valM28)
        Me.TabSprzedaz.Controls.Add(Me.valM11)
        Me.TabSprzedaz.Controls.Add(Me.lblM11)
        Me.TabSprzedaz.Controls.Add(Me.lblM28)
        Me.TabSprzedaz.Controls.Add(Me.valM10)
        Me.TabSprzedaz.Controls.Add(Me.lblM10)
        Me.TabSprzedaz.Controls.Add(Me.valM27)
        Me.TabSprzedaz.Controls.Add(Me.valM09)
        Me.TabSprzedaz.Controls.Add(Me.lblM09)
        Me.TabSprzedaz.Controls.Add(Me.lblM27)
        Me.TabSprzedaz.Controls.Add(Me.valM08)
        Me.TabSprzedaz.Controls.Add(Me.lblM08)
        Me.TabSprzedaz.Controls.Add(Me.valM26)
        Me.TabSprzedaz.Controls.Add(Me.lblM26)
        Me.TabSprzedaz.Controls.Add(Me.valM07)
        Me.TabSprzedaz.Controls.Add(Me.lblM07)
        Me.TabSprzedaz.Controls.Add(Me.valM06)
        Me.TabSprzedaz.Controls.Add(Me.lblM06)
        Me.TabSprzedaz.Controls.Add(Me.valM05)
        Me.TabSprzedaz.Controls.Add(Me.lblM05)
        Me.TabSprzedaz.Controls.Add(Me.valM04)
        Me.TabSprzedaz.Controls.Add(Me.lblM04)
        Me.TabSprzedaz.Controls.Add(Me.valM03)
        Me.TabSprzedaz.Controls.Add(Me.lblM03)
        Me.TabSprzedaz.Controls.Add(Me.valM02)
        Me.TabSprzedaz.Controls.Add(Me.lblM02)
        Me.TabSprzedaz.Controls.Add(Me.valM01)
        Me.TabSprzedaz.Controls.Add(Me.lblM01)
        Me.TabSprzedaz.Controls.Add(Me.valM00)
        Me.TabSprzedaz.Controls.Add(Me.lblM00)
        Me.TabSprzedaz.Controls.Add(Me.val1223)
        Me.TabSprzedaz.Controls.Add(Me.lbl1223)
        Me.TabSprzedaz.Controls.Add(Me.val0011)
        Me.TabSprzedaz.Controls.Add(Me.lbl0011)
        Me.TabSprzedaz.Location = New System.Drawing.Point(24, 4)
        Me.TabSprzedaz.Name = "TabSprzedaz"
        Me.TabSprzedaz.Padding = New System.Windows.Forms.Padding(3)
        Me.TabSprzedaz.Size = New System.Drawing.Size(719, 102)
        Me.TabSprzedaz.TabIndex = 0
        Me.TabSprzedaz.Text = "Sprzeda¿"
        '
        'lblPlanD
        '
        Me.lblPlanD.BackColor = System.Drawing.Color.GreenYellow
        Me.lblPlanD.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPlanD.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblPlanD.ForeColor = System.Drawing.Color.Purple
        Me.lblPlanD.Location = New System.Drawing.Point(80, 84)
        Me.lblPlanD.Name = "lblPlanD"
        Me.lblPlanD.Size = New System.Drawing.Size(70, 16)
        Me.lblPlanD.TabIndex = 149
        Me.lblPlanD.Text = "0.00"
        Me.lblPlanD.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(13, 85)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(79, 16)
        Me.Label9.TabIndex = 148
        Me.Label9.Text = "Plan d³. . . ."
        '
        'lblPlanK
        '
        Me.lblPlanK.BackColor = System.Drawing.Color.GreenYellow
        Me.lblPlanK.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPlanK.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblPlanK.ForeColor = System.Drawing.Color.Purple
        Me.lblPlanK.Location = New System.Drawing.Point(80, 68)
        Me.lblPlanK.Name = "lblPlanK"
        Me.lblPlanK.Size = New System.Drawing.Size(70, 16)
        Me.lblPlanK.TabIndex = 147
        Me.lblPlanK.Text = "0.00"
        Me.lblPlanK.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Navy
        Me.Label17.Location = New System.Drawing.Point(13, 69)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(79, 16)
        Me.Label17.TabIndex = 146
        Me.Label17.Text = "Plan kr. . . . ."
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Navy
        Me.Label18.Location = New System.Drawing.Point(13, 4)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(142, 16)
        Me.Label18.TabIndex = 145
        Me.Label18.Text = "Sum.SPRZED.Mar¿a"
        '
        'valM31
        '
        Me.valM31.BackColor = System.Drawing.Color.Azure
        Me.valM31.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM31.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM31.ForeColor = System.Drawing.Color.Purple
        Me.valM31.Location = New System.Drawing.Point(607, 83)
        Me.valM31.Name = "valM31"
        Me.valM31.Size = New System.Drawing.Size(70, 16)
        Me.valM31.TabIndex = 144
        Me.valM31.Text = "0.00"
        Me.valM31.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'valM25
        '
        Me.valM25.BackColor = System.Drawing.Color.Azure
        Me.valM25.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM25.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM25.ForeColor = System.Drawing.Color.Purple
        Me.valM25.Location = New System.Drawing.Point(471, 83)
        Me.valM25.Name = "valM25"
        Me.valM25.Size = New System.Drawing.Size(70, 16)
        Me.valM25.TabIndex = 132
        Me.valM25.Text = "0.00"
        Me.valM25.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM25
        '
        Me.lblM25.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM25.ForeColor = System.Drawing.Color.Navy
        Me.lblM25.Location = New System.Drawing.Point(415, 84)
        Me.lblM25.Name = "lblM25"
        Me.lblM25.Size = New System.Drawing.Size(79, 16)
        Me.lblM25.TabIndex = 131
        Me.lblM25.Text = "99-9999...."
        '
        'lblM31
        '
        Me.lblM31.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM31.ForeColor = System.Drawing.Color.Navy
        Me.lblM31.Location = New System.Drawing.Point(548, 84)
        Me.lblM31.Name = "lblM31"
        Me.lblM31.Size = New System.Drawing.Size(79, 16)
        Me.lblM31.TabIndex = 143
        Me.lblM31.Text = "99-9999...."
        '
        'valM24
        '
        Me.valM24.BackColor = System.Drawing.Color.Azure
        Me.valM24.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM24.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM24.ForeColor = System.Drawing.Color.Purple
        Me.valM24.Location = New System.Drawing.Point(471, 67)
        Me.valM24.Name = "valM24"
        Me.valM24.Size = New System.Drawing.Size(70, 16)
        Me.valM24.TabIndex = 130
        Me.valM24.Text = "0.00"
        Me.valM24.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM24
        '
        Me.lblM24.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM24.ForeColor = System.Drawing.Color.Navy
        Me.lblM24.Location = New System.Drawing.Point(415, 68)
        Me.lblM24.Name = "lblM24"
        Me.lblM24.Size = New System.Drawing.Size(79, 16)
        Me.lblM24.TabIndex = 129
        Me.lblM24.Text = "99-9999...."
        '
        'valM30
        '
        Me.valM30.BackColor = System.Drawing.Color.Azure
        Me.valM30.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM30.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM30.ForeColor = System.Drawing.Color.Purple
        Me.valM30.Location = New System.Drawing.Point(607, 67)
        Me.valM30.Name = "valM30"
        Me.valM30.Size = New System.Drawing.Size(70, 16)
        Me.valM30.TabIndex = 142
        Me.valM30.Text = "0.00"
        Me.valM30.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'valM23
        '
        Me.valM23.BackColor = System.Drawing.Color.Azure
        Me.valM23.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM23.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM23.ForeColor = System.Drawing.Color.Purple
        Me.valM23.Location = New System.Drawing.Point(471, 51)
        Me.valM23.Name = "valM23"
        Me.valM23.Size = New System.Drawing.Size(70, 16)
        Me.valM23.TabIndex = 128
        Me.valM23.Text = "0.00"
        Me.valM23.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM23
        '
        Me.lblM23.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM23.ForeColor = System.Drawing.Color.Navy
        Me.lblM23.Location = New System.Drawing.Point(415, 52)
        Me.lblM23.Name = "lblM23"
        Me.lblM23.Size = New System.Drawing.Size(79, 16)
        Me.lblM23.TabIndex = 127
        Me.lblM23.Text = "99-9999...."
        '
        'lblM30
        '
        Me.lblM30.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM30.ForeColor = System.Drawing.Color.Navy
        Me.lblM30.Location = New System.Drawing.Point(548, 68)
        Me.lblM30.Name = "lblM30"
        Me.lblM30.Size = New System.Drawing.Size(79, 16)
        Me.lblM30.TabIndex = 141
        Me.lblM30.Text = "99-9999...."
        '
        'valM22
        '
        Me.valM22.BackColor = System.Drawing.Color.Azure
        Me.valM22.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM22.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM22.ForeColor = System.Drawing.Color.Purple
        Me.valM22.Location = New System.Drawing.Point(471, 35)
        Me.valM22.Name = "valM22"
        Me.valM22.Size = New System.Drawing.Size(70, 16)
        Me.valM22.TabIndex = 126
        Me.valM22.Text = "0.00"
        Me.valM22.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM22
        '
        Me.lblM22.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM22.ForeColor = System.Drawing.Color.Navy
        Me.lblM22.Location = New System.Drawing.Point(415, 36)
        Me.lblM22.Name = "lblM22"
        Me.lblM22.Size = New System.Drawing.Size(79, 16)
        Me.lblM22.TabIndex = 125
        Me.lblM22.Text = "99-9999...."
        '
        'valM29
        '
        Me.valM29.BackColor = System.Drawing.Color.Azure
        Me.valM29.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM29.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM29.ForeColor = System.Drawing.Color.Purple
        Me.valM29.Location = New System.Drawing.Point(607, 51)
        Me.valM29.Name = "valM29"
        Me.valM29.Size = New System.Drawing.Size(70, 16)
        Me.valM29.TabIndex = 140
        Me.valM29.Text = "0.00"
        Me.valM29.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'valM21
        '
        Me.valM21.BackColor = System.Drawing.Color.Azure
        Me.valM21.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM21.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM21.ForeColor = System.Drawing.Color.Purple
        Me.valM21.Location = New System.Drawing.Point(471, 19)
        Me.valM21.Name = "valM21"
        Me.valM21.Size = New System.Drawing.Size(70, 16)
        Me.valM21.TabIndex = 124
        Me.valM21.Text = "0.00"
        Me.valM21.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM21
        '
        Me.lblM21.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM21.ForeColor = System.Drawing.Color.Navy
        Me.lblM21.Location = New System.Drawing.Point(415, 20)
        Me.lblM21.Name = "lblM21"
        Me.lblM21.Size = New System.Drawing.Size(79, 16)
        Me.lblM21.TabIndex = 123
        Me.lblM21.Text = "99-9999...."
        '
        'lblM29
        '
        Me.lblM29.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM29.ForeColor = System.Drawing.Color.Navy
        Me.lblM29.Location = New System.Drawing.Point(548, 52)
        Me.lblM29.Name = "lblM29"
        Me.lblM29.Size = New System.Drawing.Size(79, 16)
        Me.lblM29.TabIndex = 139
        Me.lblM29.Text = "99-9999...."
        '
        'valM20
        '
        Me.valM20.BackColor = System.Drawing.Color.Azure
        Me.valM20.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM20.ForeColor = System.Drawing.Color.Purple
        Me.valM20.Location = New System.Drawing.Point(471, 3)
        Me.valM20.Name = "valM20"
        Me.valM20.Size = New System.Drawing.Size(70, 16)
        Me.valM20.TabIndex = 122
        Me.valM20.Text = "0.00"
        Me.valM20.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM20
        '
        Me.lblM20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM20.ForeColor = System.Drawing.Color.Navy
        Me.lblM20.Location = New System.Drawing.Point(415, 4)
        Me.lblM20.Name = "lblM20"
        Me.lblM20.Size = New System.Drawing.Size(79, 16)
        Me.lblM20.TabIndex = 121
        Me.lblM20.Text = "99-9999...."
        '
        'valM28
        '
        Me.valM28.BackColor = System.Drawing.Color.Azure
        Me.valM28.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM28.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM28.ForeColor = System.Drawing.Color.Purple
        Me.valM28.Location = New System.Drawing.Point(607, 35)
        Me.valM28.Name = "valM28"
        Me.valM28.Size = New System.Drawing.Size(70, 16)
        Me.valM28.TabIndex = 138
        Me.valM28.Text = "0.00"
        Me.valM28.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'valM11
        '
        Me.valM11.BackColor = System.Drawing.Color.LightCyan
        Me.valM11.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM11.ForeColor = System.Drawing.Color.Purple
        Me.valM11.Location = New System.Drawing.Point(342, 83)
        Me.valM11.Name = "valM11"
        Me.valM11.Size = New System.Drawing.Size(70, 16)
        Me.valM11.TabIndex = 120
        Me.valM11.Text = "0.00"
        Me.valM11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM11
        '
        Me.lblM11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM11.ForeColor = System.Drawing.Color.Navy
        Me.lblM11.Location = New System.Drawing.Point(286, 83)
        Me.lblM11.Name = "lblM11"
        Me.lblM11.Size = New System.Drawing.Size(79, 16)
        Me.lblM11.TabIndex = 119
        Me.lblM11.Text = "99-9999...."
        '
        'lblM28
        '
        Me.lblM28.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM28.ForeColor = System.Drawing.Color.Navy
        Me.lblM28.Location = New System.Drawing.Point(548, 36)
        Me.lblM28.Name = "lblM28"
        Me.lblM28.Size = New System.Drawing.Size(79, 16)
        Me.lblM28.TabIndex = 137
        Me.lblM28.Text = "99-9999...."
        '
        'valM10
        '
        Me.valM10.BackColor = System.Drawing.Color.LightCyan
        Me.valM10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM10.ForeColor = System.Drawing.Color.Purple
        Me.valM10.Location = New System.Drawing.Point(342, 67)
        Me.valM10.Name = "valM10"
        Me.valM10.Size = New System.Drawing.Size(70, 16)
        Me.valM10.TabIndex = 118
        Me.valM10.Text = "0.00"
        Me.valM10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM10
        '
        Me.lblM10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM10.ForeColor = System.Drawing.Color.Navy
        Me.lblM10.Location = New System.Drawing.Point(286, 67)
        Me.lblM10.Name = "lblM10"
        Me.lblM10.Size = New System.Drawing.Size(79, 16)
        Me.lblM10.TabIndex = 117
        Me.lblM10.Text = "99-9999...."
        '
        'valM27
        '
        Me.valM27.BackColor = System.Drawing.Color.Azure
        Me.valM27.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM27.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM27.ForeColor = System.Drawing.Color.Purple
        Me.valM27.Location = New System.Drawing.Point(607, 19)
        Me.valM27.Name = "valM27"
        Me.valM27.Size = New System.Drawing.Size(70, 16)
        Me.valM27.TabIndex = 136
        Me.valM27.Text = "0.00"
        Me.valM27.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'valM09
        '
        Me.valM09.BackColor = System.Drawing.Color.LightCyan
        Me.valM09.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM09.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM09.ForeColor = System.Drawing.Color.Purple
        Me.valM09.Location = New System.Drawing.Point(342, 51)
        Me.valM09.Name = "valM09"
        Me.valM09.Size = New System.Drawing.Size(70, 16)
        Me.valM09.TabIndex = 116
        Me.valM09.Text = "0.00"
        Me.valM09.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM09
        '
        Me.lblM09.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM09.ForeColor = System.Drawing.Color.Navy
        Me.lblM09.Location = New System.Drawing.Point(286, 51)
        Me.lblM09.Name = "lblM09"
        Me.lblM09.Size = New System.Drawing.Size(79, 16)
        Me.lblM09.TabIndex = 115
        Me.lblM09.Text = "99-9999...."
        '
        'lblM27
        '
        Me.lblM27.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM27.ForeColor = System.Drawing.Color.Navy
        Me.lblM27.Location = New System.Drawing.Point(548, 20)
        Me.lblM27.Name = "lblM27"
        Me.lblM27.Size = New System.Drawing.Size(79, 16)
        Me.lblM27.TabIndex = 135
        Me.lblM27.Text = "99-9999...."
        '
        'valM08
        '
        Me.valM08.BackColor = System.Drawing.Color.LightCyan
        Me.valM08.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM08.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM08.ForeColor = System.Drawing.Color.Purple
        Me.valM08.Location = New System.Drawing.Point(342, 35)
        Me.valM08.Name = "valM08"
        Me.valM08.Size = New System.Drawing.Size(70, 16)
        Me.valM08.TabIndex = 114
        Me.valM08.Text = "0.00"
        Me.valM08.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM08
        '
        Me.lblM08.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM08.ForeColor = System.Drawing.Color.Navy
        Me.lblM08.Location = New System.Drawing.Point(286, 35)
        Me.lblM08.Name = "lblM08"
        Me.lblM08.Size = New System.Drawing.Size(79, 16)
        Me.lblM08.TabIndex = 113
        Me.lblM08.Text = "99-9999...."
        '
        'valM26
        '
        Me.valM26.BackColor = System.Drawing.Color.Azure
        Me.valM26.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM26.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM26.ForeColor = System.Drawing.Color.Purple
        Me.valM26.Location = New System.Drawing.Point(607, 3)
        Me.valM26.Name = "valM26"
        Me.valM26.Size = New System.Drawing.Size(70, 16)
        Me.valM26.TabIndex = 134
        Me.valM26.Text = "0.00"
        Me.valM26.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM26
        '
        Me.lblM26.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM26.ForeColor = System.Drawing.Color.Navy
        Me.lblM26.Location = New System.Drawing.Point(548, 4)
        Me.lblM26.Name = "lblM26"
        Me.lblM26.Size = New System.Drawing.Size(79, 16)
        Me.lblM26.TabIndex = 133
        Me.lblM26.Text = "99-9999...."
        '
        'valM07
        '
        Me.valM07.BackColor = System.Drawing.Color.LightCyan
        Me.valM07.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM07.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM07.ForeColor = System.Drawing.Color.Purple
        Me.valM07.Location = New System.Drawing.Point(342, 19)
        Me.valM07.Name = "valM07"
        Me.valM07.Size = New System.Drawing.Size(70, 16)
        Me.valM07.TabIndex = 112
        Me.valM07.Text = "0.00"
        Me.valM07.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM07
        '
        Me.lblM07.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM07.ForeColor = System.Drawing.Color.Navy
        Me.lblM07.Location = New System.Drawing.Point(286, 19)
        Me.lblM07.Name = "lblM07"
        Me.lblM07.Size = New System.Drawing.Size(79, 16)
        Me.lblM07.TabIndex = 111
        Me.lblM07.Text = "99-9999...."
        '
        'valM06
        '
        Me.valM06.BackColor = System.Drawing.Color.LightCyan
        Me.valM06.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM06.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM06.ForeColor = System.Drawing.Color.Purple
        Me.valM06.Location = New System.Drawing.Point(342, 3)
        Me.valM06.Name = "valM06"
        Me.valM06.Size = New System.Drawing.Size(70, 16)
        Me.valM06.TabIndex = 110
        Me.valM06.Text = "0.00"
        Me.valM06.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM06
        '
        Me.lblM06.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM06.ForeColor = System.Drawing.Color.Navy
        Me.lblM06.Location = New System.Drawing.Point(286, 3)
        Me.lblM06.Name = "lblM06"
        Me.lblM06.Size = New System.Drawing.Size(79, 16)
        Me.lblM06.TabIndex = 109
        Me.lblM06.Text = "99-9999...."
        '
        'valM05
        '
        Me.valM05.BackColor = System.Drawing.Color.LightCyan
        Me.valM05.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM05.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM05.ForeColor = System.Drawing.Color.Purple
        Me.valM05.Location = New System.Drawing.Point(211, 84)
        Me.valM05.Name = "valM05"
        Me.valM05.Size = New System.Drawing.Size(70, 16)
        Me.valM05.TabIndex = 108
        Me.valM05.Text = "0.00"
        Me.valM05.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM05
        '
        Me.lblM05.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM05.ForeColor = System.Drawing.Color.Navy
        Me.lblM05.Location = New System.Drawing.Point(155, 84)
        Me.lblM05.Name = "lblM05"
        Me.lblM05.Size = New System.Drawing.Size(79, 16)
        Me.lblM05.TabIndex = 107
        Me.lblM05.Text = "99-9999...."
        '
        'valM04
        '
        Me.valM04.BackColor = System.Drawing.Color.LightCyan
        Me.valM04.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM04.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM04.ForeColor = System.Drawing.Color.Purple
        Me.valM04.Location = New System.Drawing.Point(211, 68)
        Me.valM04.Name = "valM04"
        Me.valM04.Size = New System.Drawing.Size(70, 16)
        Me.valM04.TabIndex = 106
        Me.valM04.Text = "0.00"
        Me.valM04.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM04
        '
        Me.lblM04.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM04.ForeColor = System.Drawing.Color.Navy
        Me.lblM04.Location = New System.Drawing.Point(155, 68)
        Me.lblM04.Name = "lblM04"
        Me.lblM04.Size = New System.Drawing.Size(79, 16)
        Me.lblM04.TabIndex = 105
        Me.lblM04.Text = "99-9999...."
        '
        'valM03
        '
        Me.valM03.BackColor = System.Drawing.Color.LightCyan
        Me.valM03.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM03.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM03.ForeColor = System.Drawing.Color.Purple
        Me.valM03.Location = New System.Drawing.Point(211, 52)
        Me.valM03.Name = "valM03"
        Me.valM03.Size = New System.Drawing.Size(70, 16)
        Me.valM03.TabIndex = 104
        Me.valM03.Text = "0.00"
        Me.valM03.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM03
        '
        Me.lblM03.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM03.ForeColor = System.Drawing.Color.Navy
        Me.lblM03.Location = New System.Drawing.Point(155, 52)
        Me.lblM03.Name = "lblM03"
        Me.lblM03.Size = New System.Drawing.Size(79, 16)
        Me.lblM03.TabIndex = 103
        Me.lblM03.Text = "99-9999...."
        '
        'valM02
        '
        Me.valM02.BackColor = System.Drawing.Color.LightCyan
        Me.valM02.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM02.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM02.ForeColor = System.Drawing.Color.Purple
        Me.valM02.Location = New System.Drawing.Point(211, 36)
        Me.valM02.Name = "valM02"
        Me.valM02.Size = New System.Drawing.Size(70, 16)
        Me.valM02.TabIndex = 102
        Me.valM02.Text = "0.00"
        Me.valM02.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM02
        '
        Me.lblM02.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM02.ForeColor = System.Drawing.Color.Navy
        Me.lblM02.Location = New System.Drawing.Point(155, 36)
        Me.lblM02.Name = "lblM02"
        Me.lblM02.Size = New System.Drawing.Size(79, 16)
        Me.lblM02.TabIndex = 101
        Me.lblM02.Text = "99-9999...."
        '
        'valM01
        '
        Me.valM01.BackColor = System.Drawing.Color.LightCyan
        Me.valM01.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM01.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM01.ForeColor = System.Drawing.Color.Purple
        Me.valM01.Location = New System.Drawing.Point(211, 20)
        Me.valM01.Name = "valM01"
        Me.valM01.Size = New System.Drawing.Size(70, 16)
        Me.valM01.TabIndex = 100
        Me.valM01.Text = "0.00"
        Me.valM01.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM01
        '
        Me.lblM01.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM01.ForeColor = System.Drawing.Color.Navy
        Me.lblM01.Location = New System.Drawing.Point(155, 20)
        Me.lblM01.Name = "lblM01"
        Me.lblM01.Size = New System.Drawing.Size(79, 16)
        Me.lblM01.TabIndex = 99
        Me.lblM01.Text = "99-9999...."
        '
        'valM00
        '
        Me.valM00.BackColor = System.Drawing.Color.LightCyan
        Me.valM00.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.valM00.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.valM00.ForeColor = System.Drawing.Color.Purple
        Me.valM00.Location = New System.Drawing.Point(211, 4)
        Me.valM00.Name = "valM00"
        Me.valM00.Size = New System.Drawing.Size(70, 16)
        Me.valM00.TabIndex = 98
        Me.valM00.Text = "0.00"
        Me.valM00.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblM00
        '
        Me.lblM00.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblM00.ForeColor = System.Drawing.Color.Navy
        Me.lblM00.Location = New System.Drawing.Point(155, 4)
        Me.lblM00.Name = "lblM00"
        Me.lblM00.Size = New System.Drawing.Size(79, 16)
        Me.lblM00.TabIndex = 97
        Me.lblM00.Text = "99-9999...."
        '
        'val1223
        '
        Me.val1223.BackColor = System.Drawing.Color.PaleTurquoise
        Me.val1223.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.val1223.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.val1223.ForeColor = System.Drawing.Color.Purple
        Me.val1223.Location = New System.Drawing.Point(80, 36)
        Me.val1223.Name = "val1223"
        Me.val1223.Size = New System.Drawing.Size(70, 16)
        Me.val1223.TabIndex = 96
        Me.val1223.Text = "0.00"
        Me.val1223.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbl1223
        '
        Me.lbl1223.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lbl1223.ForeColor = System.Drawing.Color.Navy
        Me.lbl1223.Location = New System.Drawing.Point(13, 37)
        Me.lbl1223.Name = "lbl1223"
        Me.lbl1223.Size = New System.Drawing.Size(79, 16)
        Me.lbl1223.TabIndex = 95
        Me.lbl1223.Text = "Okres II. . . ."
        '
        'val0011
        '
        Me.val0011.BackColor = System.Drawing.Color.PaleTurquoise
        Me.val0011.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.val0011.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.val0011.ForeColor = System.Drawing.Color.Purple
        Me.val0011.Location = New System.Drawing.Point(80, 20)
        Me.val0011.Name = "val0011"
        Me.val0011.Size = New System.Drawing.Size(70, 16)
        Me.val0011.TabIndex = 94
        Me.val0011.Text = "0.00"
        Me.val0011.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbl0011
        '
        Me.lbl0011.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lbl0011.ForeColor = System.Drawing.Color.Navy
        Me.lbl0011.Location = New System.Drawing.Point(13, 21)
        Me.lbl0011.Name = "lbl0011"
        Me.lbl0011.Size = New System.Drawing.Size(79, 16)
        Me.lbl0011.TabIndex = 93
        Me.lbl0011.Text = "Okres I. . . . ."
        '
        'lbleMail
        '
        Me.lbleMail.BackColor = System.Drawing.SystemColors.Control
        Me.lbleMail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbleMail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lbleMail.ForeColor = System.Drawing.Color.Purple
        Me.lbleMail.Location = New System.Drawing.Point(168, 72)
        Me.lbleMail.Name = "lbleMail"
        Me.lbleMail.Size = New System.Drawing.Size(232, 16)
        Me.lbleMail.TabIndex = 20
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
        'lblTelefon
        '
        Me.lblTelefon.BackColor = System.Drawing.SystemColors.Control
        Me.lblTelefon.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefon.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTelefon.ForeColor = System.Drawing.Color.Purple
        Me.lblTelefon.Location = New System.Drawing.Point(168, 56)
        Me.lblTelefon.Name = "lblTelefon"
        Me.lblTelefon.Size = New System.Drawing.Size(232, 16)
        Me.lblTelefon.TabIndex = 9
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
        Me.lblPotencjal.Size = New System.Drawing.Size(120, 16)
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
        Me.Label8.Text = "Potencja³ . . . . . . . ."
        '
        'lblTransakcje
        '
        Me.lblTransakcje.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTransakcje.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTransakcje.ForeColor = System.Drawing.Color.Purple
        Me.lblTransakcje.Location = New System.Drawing.Point(504, 56)
        Me.lblTransakcje.Name = "lblTransakcje"
        Me.lblTransakcje.Size = New System.Drawing.Size(120, 16)
        Me.lblTransakcje.TabIndex = 19
        '
        'lblNastKontakt
        '
        Me.lblNastKontakt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNastKontakt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNastKontakt.ForeColor = System.Drawing.Color.Purple
        Me.lblNastKontakt.Location = New System.Drawing.Point(504, 40)
        Me.lblNastKontakt.Name = "lblNastKontakt"
        Me.lblNastKontakt.Size = New System.Drawing.Size(120, 16)
        Me.lblNastKontakt.TabIndex = 18
        '
        'lblOstKontakt
        '
        Me.lblOstKontakt.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOstKontakt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOstKontakt.ForeColor = System.Drawing.Color.Purple
        Me.lblOstKontakt.Location = New System.Drawing.Point(504, 24)
        Me.lblOstKontakt.Name = "lblOstKontakt"
        Me.lblOstKontakt.Size = New System.Drawing.Size(120, 16)
        Me.lblOstKontakt.TabIndex = 17
        '
        'lblOpiekun
        '
        Me.lblOpiekun.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOpiekun.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOpiekun.ForeColor = System.Drawing.Color.Purple
        Me.lblOpiekun.Location = New System.Drawing.Point(504, 8)
        Me.lblOpiekun.Name = "lblOpiekun"
        Me.lblOpiekun.Size = New System.Drawing.Size(120, 16)
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
        Me.Label10.Text = "Opiekun . . . . . . . ."
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
        'cmdWybierzSchemat
        '
        Me.cmdWybierzSchemat.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzSchemat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWybierzSchemat.ForeColor = System.Drawing.Color.Black
        Me.cmdWybierzSchemat.Location = New System.Drawing.Point(183, 238)
        Me.cmdWybierzSchemat.Name = "cmdWybierzSchemat"
        Me.cmdWybierzSchemat.Size = New System.Drawing.Size(56, 23)
        Me.cmdWybierzSchemat.TabIndex = 57
        Me.cmdWybierzSchemat.Text = "Wybierz"
        '
        'cmdWysylaj
        '
        Me.cmdWysylaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWysylaj.ForeColor = System.Drawing.Color.Black
        Me.cmdWysylaj.Image = CType(resources.GetObject("cmdWysylaj.Image"), System.Drawing.Image)
        Me.cmdWysylaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWysylaj.Location = New System.Drawing.Point(871, 478)
        Me.cmdWysylaj.Name = "cmdWysylaj"
        Me.cmdWysylaj.Size = New System.Drawing.Size(80, 24)
        Me.cmdWysylaj.TabIndex = 16
        Me.cmdWysylaj.Text = "Wyœlij"
        Me.cmdWysylaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdWysylaj, "Wysy³anie listu eMail do wszystkich klientów" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "widocznych na ekranie (przefiltrowa" &
        "nych).")
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
        Me.stbFiltry.Size = New System.Drawing.Size(1127, 22)
        Me.stbFiltry.TabIndex = 67
        Me.stbFiltry.Text = "StatusBar1"
        '
        'cmdWszystko
        '
        Me.cmdWszystko.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdWszystko.ForeColor = System.Drawing.Color.Black
        Me.cmdWszystko.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystko.Location = New System.Drawing.Point(724, 0)
        Me.cmdWszystko.Name = "cmdWszystko"
        Me.cmdWszystko.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystko.TabIndex = 115
        Me.ToolTip1.SetToolTip(Me.cmdWszystko, "Wyczyœæ wszystkie filtry.")
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(702, 0)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 171
        Me.ToolTip1.SetToolTip(Me.cmdFi, "Unikalne wartoœci pól tej kolumny.")
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'ToolTip1
        '
        Me.ToolTip1.AutoPopDelay = 5000
        Me.ToolTip1.InitialDelay = 100
        Me.ToolTip1.IsBalloon = True
        Me.ToolTip1.ReshowDelay = 100
        '
        'cmdOdswiez
        '
        Me.cmdOdswiez.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOdswiez.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdOdswiez.ForeColor = System.Drawing.Color.Black
        Me.cmdOdswiez.Image = CType(resources.GetObject("cmdOdswiez.Image"), System.Drawing.Image)
        Me.cmdOdswiez.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOdswiez.Location = New System.Drawing.Point(957, 479)
        Me.cmdOdswiez.Name = "cmdOdswiez"
        Me.cmdOdswiez.Size = New System.Drawing.Size(92, 24)
        Me.cmdOdswiez.TabIndex = 173
        Me.cmdOdswiez.Text = "Odœwie¿"
        Me.cmdOdswiez.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdOdswiez, "Odœwie¿enie informacji." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Pobranie Kartoteki Klientów z serwera aby pokazaæ wszyst" &
        "kie" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "równolegle wykonywane modyfikacje." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10))
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
        'cmdPrzeliczSprzedaz
        '
        Me.cmdPrzeliczSprzedaz.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPrzeliczSprzedaz.ForeColor = System.Drawing.Color.Black
        Me.cmdPrzeliczSprzedaz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPrzeliczSprzedaz.Location = New System.Drawing.Point(973, 1)
        Me.cmdPrzeliczSprzedaz.Name = "cmdPrzeliczSprzedaz"
        Me.cmdPrzeliczSprzedaz.Size = New System.Drawing.Size(90, 55)
        Me.cmdPrzeliczSprzedaz.TabIndex = 217
        Me.cmdPrzeliczSprzedaz.Text = "Przelicz Tabelê"
        Me.ToolTip1.SetToolTip(Me.cmdPrzeliczSprzedaz, "Dostosuj informacje wyœwietlane w Tabeli do zdefiniowanego zakresu.")
        '
        'cmdDoExcela
        '
        Me.cmdDoExcela.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdDoExcela.ForeColor = System.Drawing.Color.Black
        Me.cmdDoExcela.Image = CType(resources.GetObject("cmdDoExcela.Image"), System.Drawing.Image)
        Me.cmdDoExcela.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDoExcela.Location = New System.Drawing.Point(149, 478)
        Me.cmdDoExcela.Name = "cmdDoExcela"
        Me.cmdDoExcela.Size = New System.Drawing.Size(24, 23)
        Me.cmdDoExcela.TabIndex = 205
        Me.cmdDoExcela.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdDoExcela, "Eksport do EXCELa informacji o klientach" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "widocznych na ekranie (przefiltrowanych" &
        ").")
        '
        'cmdOdznacz2
        '
        Me.cmdOdznacz2.Image = CType(resources.GetObject("cmdOdznacz2.Image"), System.Drawing.Image)
        Me.cmdOdznacz2.Location = New System.Drawing.Point(450, 0)
        Me.cmdOdznacz2.Name = "cmdOdznacz2"
        Me.cmdOdznacz2.Size = New System.Drawing.Size(24, 21)
        Me.cmdOdznacz2.TabIndex = 241
        Me.cmdOdznacz2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdOdznacz2, "Odznacz wszystko")
        '
        'cmdZaznacz2
        '
        Me.cmdZaznacz2.Image = CType(resources.GetObject("cmdZaznacz2.Image"), System.Drawing.Image)
        Me.cmdZaznacz2.Location = New System.Drawing.Point(420, 0)
        Me.cmdZaznacz2.Name = "cmdZaznacz2"
        Me.cmdZaznacz2.Size = New System.Drawing.Size(24, 21)
        Me.cmdZaznacz2.TabIndex = 240
        Me.cmdZaznacz2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdZaznacz2, "Zaznacz wszystko")
        '
        'cmdOdznacz3
        '
        Me.cmdOdznacz3.Image = CType(resources.GetObject("cmdOdznacz3.Image"), System.Drawing.Image)
        Me.cmdOdznacz3.Location = New System.Drawing.Point(938, 0)
        Me.cmdOdznacz3.Name = "cmdOdznacz3"
        Me.cmdOdznacz3.Size = New System.Drawing.Size(24, 21)
        Me.cmdOdznacz3.TabIndex = 243
        Me.cmdOdznacz3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdOdznacz3, "Odznacz wszystko")
        '
        'cmdZaznacz3
        '
        Me.cmdZaznacz3.Image = CType(resources.GetObject("cmdZaznacz3.Image"), System.Drawing.Image)
        Me.cmdZaznacz3.Location = New System.Drawing.Point(908, 0)
        Me.cmdZaznacz3.Name = "cmdZaznacz3"
        Me.cmdZaznacz3.Size = New System.Drawing.Size(24, 21)
        Me.cmdZaznacz3.TabIndex = 242
        Me.cmdZaznacz3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdZaznacz3, "Zaznacz wszystko")
        '
        'HorizScrollTimer
        '
        '
        'cbxCoAnalizowac
        '
        Me.cbxCoAnalizowac.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cbxCoAnalizowac.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxCoAnalizowac.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxCoAnalizowac.ForeColor = System.Drawing.Color.Purple
        Me.cbxCoAnalizowac.Location = New System.Drawing.Point(678, 20)
        Me.cbxCoAnalizowac.Name = "cbxCoAnalizowac"
        Me.cbxCoAnalizowac.Size = New System.Drawing.Size(216, 22)
        Me.cbxCoAnalizowac.TabIndex = 175
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label16)
        Me.Panel2.Controls.Add(Me.cbxCoAnalizowac)
        Me.Panel2.Controls.Add(Me.chkSkup)
        Me.Panel2.Controls.Add(Me.chkDrukarki)
        Me.Panel2.Controls.Add(Me.chkStrony)
        Me.Panel2.Controls.Add(Me.chkNajem)
        Me.Panel2.Controls.Add(Me.chkOrygTonery)
        Me.Panel2.Controls.Add(Me.chkOrygAtrament)
        Me.Panel2.Controls.Add(Me.chkPryzmatTonery)
        Me.Panel2.Controls.Add(Me.chkPryzmatAtrament)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.cmdOdznacz3)
        Me.Panel2.Controls.Add(Me.cmdZaznacz3)
        Me.Panel2.Controls.Add(Me.cmdOdznacz2)
        Me.Panel2.Controls.Add(Me.cmdZaznacz2)
        Me.Panel2.Controls.Add(Me.cmdPrzeliczSprzedaz)
        Me.Panel2.Controls.Add(Me.Chk212)
        Me.Panel2.Controls.Add(Me.Chk211)
        Me.Panel2.Controls.Add(Me.Chk210)
        Me.Panel2.Controls.Add(Me.Chk209)
        Me.Panel2.Controls.Add(Me.Chk208)
        Me.Panel2.Controls.Add(Me.Chk207)
        Me.Panel2.Controls.Add(Me.Chk206)
        Me.Panel2.Controls.Add(Me.Chk205)
        Me.Panel2.Controls.Add(Me.Chk204)
        Me.Panel2.Controls.Add(Me.Chk203)
        Me.Panel2.Controls.Add(Me.Chk202)
        Me.Panel2.Controls.Add(Me.Chk201)
        Me.Panel2.Controls.Add(Me.lblOkres2)
        Me.Panel2.Controls.Add(Me.Chk112)
        Me.Panel2.Controls.Add(Me.Chk111)
        Me.Panel2.Controls.Add(Me.Chk110)
        Me.Panel2.Controls.Add(Me.Chk109)
        Me.Panel2.Controls.Add(Me.Chk108)
        Me.Panel2.Controls.Add(Me.Chk107)
        Me.Panel2.Controls.Add(Me.Chk106)
        Me.Panel2.Controls.Add(Me.Chk105)
        Me.Panel2.Controls.Add(Me.Chk104)
        Me.Panel2.Controls.Add(Me.Chk103)
        Me.Panel2.Controls.Add(Me.Chk102)
        Me.Panel2.Controls.Add(Me.Chk101)
        Me.Panel2.Controls.Add(Me.lblOkres1)
        Me.Panel2.Location = New System.Drawing.Point(8, 410)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1127, 60)
        Me.Panel2.TabIndex = 203
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Navy
        Me.Label16.Location = New System.Drawing.Point(560, 23)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(113, 16)
        Me.Label16.TabIndex = 253
        Me.Label16.Text = "DANE DO ANALIZY :"
        '
        'chkSkup
        '
        Me.chkSkup.ForeColor = System.Drawing.Color.Navy
        Me.chkSkup.Location = New System.Drawing.Point(480, 39)
        Me.chkSkup.Name = "chkSkup"
        Me.chkSkup.Size = New System.Drawing.Size(160, 17)
        Me.chkSkup.TabIndex = 252
        Me.chkSkup.Text = "Skup"
        Me.chkSkup.UseVisualStyleBackColor = True
        '
        'chkDrukarki
        '
        Me.chkDrukarki.ForeColor = System.Drawing.Color.Navy
        Me.chkDrukarki.Location = New System.Drawing.Point(402, 39)
        Me.chkDrukarki.Name = "chkDrukarki"
        Me.chkDrukarki.Size = New System.Drawing.Size(160, 17)
        Me.chkDrukarki.TabIndex = 251
        Me.chkDrukarki.Text = "Drukarki"
        Me.chkDrukarki.UseVisualStyleBackColor = True
        '
        'chkStrony
        '
        Me.chkStrony.ForeColor = System.Drawing.Color.Navy
        Me.chkStrony.Location = New System.Drawing.Point(480, 22)
        Me.chkStrony.Name = "chkStrony"
        Me.chkStrony.Size = New System.Drawing.Size(160, 17)
        Me.chkStrony.TabIndex = 250
        Me.chkStrony.Text = "Strony"
        Me.chkStrony.UseVisualStyleBackColor = True
        '
        'chkNajem
        '
        Me.chkNajem.ForeColor = System.Drawing.Color.Navy
        Me.chkNajem.Location = New System.Drawing.Point(402, 22)
        Me.chkNajem.Name = "chkNajem"
        Me.chkNajem.Size = New System.Drawing.Size(160, 17)
        Me.chkNajem.TabIndex = 249
        Me.chkNajem.Text = "Najem"
        Me.chkNajem.UseVisualStyleBackColor = True
        '
        'chkOrygTonery
        '
        Me.chkOrygTonery.ForeColor = System.Drawing.Color.Navy
        Me.chkOrygTonery.Location = New System.Drawing.Point(256, 39)
        Me.chkOrygTonery.Name = "chkOrygTonery"
        Me.chkOrygTonery.Size = New System.Drawing.Size(160, 17)
        Me.chkOrygTonery.TabIndex = 248
        Me.chkOrygTonery.Text = "Oryginalne TONERY"
        Me.chkOrygTonery.UseVisualStyleBackColor = True
        '
        'chkOrygAtrament
        '
        Me.chkOrygAtrament.ForeColor = System.Drawing.Color.Navy
        Me.chkOrygAtrament.Location = New System.Drawing.Point(110, 39)
        Me.chkOrygAtrament.Name = "chkOrygAtrament"
        Me.chkOrygAtrament.Size = New System.Drawing.Size(160, 17)
        Me.chkOrygAtrament.TabIndex = 247
        Me.chkOrygAtrament.Text = "Oryginalne ATRAMENT"
        Me.chkOrygAtrament.UseVisualStyleBackColor = True
        '
        'chkPryzmatTonery
        '
        Me.chkPryzmatTonery.ForeColor = System.Drawing.Color.Navy
        Me.chkPryzmatTonery.Location = New System.Drawing.Point(256, 22)
        Me.chkPryzmatTonery.Name = "chkPryzmatTonery"
        Me.chkPryzmatTonery.Size = New System.Drawing.Size(160, 17)
        Me.chkPryzmatTonery.TabIndex = 246
        Me.chkPryzmatTonery.Text = "Pryzmat TONERY"
        Me.chkPryzmatTonery.UseVisualStyleBackColor = True
        '
        'chkPryzmatAtrament
        '
        Me.chkPryzmatAtrament.ForeColor = System.Drawing.Color.Navy
        Me.chkPryzmatAtrament.Location = New System.Drawing.Point(110, 22)
        Me.chkPryzmatAtrament.Name = "chkPryzmatAtrament"
        Me.chkPryzmatAtrament.Size = New System.Drawing.Size(160, 17)
        Me.chkPryzmatAtrament.TabIndex = 245
        Me.chkPryzmatAtrament.Text = "Pryzmat ATRAMENT"
        Me.chkPryzmatAtrament.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(20, 23)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(173, 16)
        Me.Label6.TabIndex = 244
        Me.Label6.Text = "ASORTYMENT :"
        '
        'Chk212
        '
        Me.Chk212.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk212.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk212.ForeColor = System.Drawing.Color.Navy
        Me.Chk212.Location = New System.Drawing.Point(876, 2)
        Me.Chk212.Name = "Chk212"
        Me.Chk212.Size = New System.Drawing.Size(18, 18)
        Me.Chk212.TabIndex = 215
        Me.Chk212.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk212.UseVisualStyleBackColor = True
        '
        'Chk211
        '
        Me.Chk211.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk211.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk211.ForeColor = System.Drawing.Color.Navy
        Me.Chk211.Location = New System.Drawing.Point(858, 2)
        Me.Chk211.Name = "Chk211"
        Me.Chk211.Size = New System.Drawing.Size(18, 18)
        Me.Chk211.TabIndex = 214
        Me.Chk211.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk211.UseVisualStyleBackColor = True
        '
        'Chk210
        '
        Me.Chk210.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk210.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk210.ForeColor = System.Drawing.Color.Navy
        Me.Chk210.Location = New System.Drawing.Point(840, 2)
        Me.Chk210.Name = "Chk210"
        Me.Chk210.Size = New System.Drawing.Size(18, 18)
        Me.Chk210.TabIndex = 213
        Me.Chk210.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk210.UseVisualStyleBackColor = True
        '
        'Chk209
        '
        Me.Chk209.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk209.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk209.ForeColor = System.Drawing.Color.Navy
        Me.Chk209.Location = New System.Drawing.Point(822, 2)
        Me.Chk209.Name = "Chk209"
        Me.Chk209.Size = New System.Drawing.Size(18, 18)
        Me.Chk209.TabIndex = 212
        Me.Chk209.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk209.UseVisualStyleBackColor = True
        '
        'Chk208
        '
        Me.Chk208.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk208.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk208.ForeColor = System.Drawing.Color.Navy
        Me.Chk208.Location = New System.Drawing.Point(804, 2)
        Me.Chk208.Name = "Chk208"
        Me.Chk208.Size = New System.Drawing.Size(18, 18)
        Me.Chk208.TabIndex = 211
        Me.Chk208.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk208.UseVisualStyleBackColor = True
        '
        'Chk207
        '
        Me.Chk207.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk207.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk207.ForeColor = System.Drawing.Color.Navy
        Me.Chk207.Location = New System.Drawing.Point(786, 2)
        Me.Chk207.Name = "Chk207"
        Me.Chk207.Size = New System.Drawing.Size(18, 18)
        Me.Chk207.TabIndex = 210
        Me.Chk207.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk207.UseVisualStyleBackColor = True
        '
        'Chk206
        '
        Me.Chk206.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk206.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk206.ForeColor = System.Drawing.Color.Navy
        Me.Chk206.Location = New System.Drawing.Point(768, 2)
        Me.Chk206.Name = "Chk206"
        Me.Chk206.Size = New System.Drawing.Size(18, 18)
        Me.Chk206.TabIndex = 209
        Me.Chk206.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk206.UseVisualStyleBackColor = True
        '
        'Chk205
        '
        Me.Chk205.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk205.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk205.ForeColor = System.Drawing.Color.Navy
        Me.Chk205.Location = New System.Drawing.Point(750, 2)
        Me.Chk205.Name = "Chk205"
        Me.Chk205.Size = New System.Drawing.Size(18, 18)
        Me.Chk205.TabIndex = 208
        Me.Chk205.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk205.UseVisualStyleBackColor = True
        '
        'Chk204
        '
        Me.Chk204.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk204.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk204.ForeColor = System.Drawing.Color.Navy
        Me.Chk204.Location = New System.Drawing.Point(732, 2)
        Me.Chk204.Name = "Chk204"
        Me.Chk204.Size = New System.Drawing.Size(18, 18)
        Me.Chk204.TabIndex = 207
        Me.Chk204.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk204.UseVisualStyleBackColor = True
        '
        'Chk203
        '
        Me.Chk203.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk203.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk203.ForeColor = System.Drawing.Color.Navy
        Me.Chk203.Location = New System.Drawing.Point(714, 2)
        Me.Chk203.Name = "Chk203"
        Me.Chk203.Size = New System.Drawing.Size(18, 18)
        Me.Chk203.TabIndex = 206
        Me.Chk203.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk203.UseVisualStyleBackColor = True
        '
        'Chk202
        '
        Me.Chk202.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk202.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk202.ForeColor = System.Drawing.Color.Navy
        Me.Chk202.Location = New System.Drawing.Point(696, 2)
        Me.Chk202.Name = "Chk202"
        Me.Chk202.Size = New System.Drawing.Size(18, 18)
        Me.Chk202.TabIndex = 205
        Me.Chk202.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk202.UseVisualStyleBackColor = True
        '
        'Chk201
        '
        Me.Chk201.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk201.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk201.ForeColor = System.Drawing.Color.Navy
        Me.Chk201.Location = New System.Drawing.Point(678, 2)
        Me.Chk201.Name = "Chk201"
        Me.Chk201.Size = New System.Drawing.Size(18, 18)
        Me.Chk201.TabIndex = 204
        Me.Chk201.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk201.UseVisualStyleBackColor = True
        '
        'lblOkres2
        '
        Me.lblOkres2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOkres2.ForeColor = System.Drawing.Color.Navy
        Me.lblOkres2.Location = New System.Drawing.Point(560, 5)
        Me.lblOkres2.Name = "lblOkres2"
        Me.lblOkres2.Size = New System.Drawing.Size(113, 16)
        Me.lblOkres2.TabIndex = 203
        Me.lblOkres2.Text = "RRRR.MM : 0..-11 . . . ."
        '
        'Chk112
        '
        Me.Chk112.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk112.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk112.ForeColor = System.Drawing.Color.Navy
        Me.Chk112.Location = New System.Drawing.Point(391, 2)
        Me.Chk112.Name = "Chk112"
        Me.Chk112.Size = New System.Drawing.Size(18, 18)
        Me.Chk112.TabIndex = 202
        Me.Chk112.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk112.UseVisualStyleBackColor = True
        '
        'Chk111
        '
        Me.Chk111.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk111.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk111.ForeColor = System.Drawing.Color.Navy
        Me.Chk111.Location = New System.Drawing.Point(373, 2)
        Me.Chk111.Name = "Chk111"
        Me.Chk111.Size = New System.Drawing.Size(18, 18)
        Me.Chk111.TabIndex = 201
        Me.Chk111.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk111.UseVisualStyleBackColor = True
        '
        'Chk110
        '
        Me.Chk110.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk110.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk110.ForeColor = System.Drawing.Color.Navy
        Me.Chk110.Location = New System.Drawing.Point(355, 2)
        Me.Chk110.Name = "Chk110"
        Me.Chk110.Size = New System.Drawing.Size(18, 18)
        Me.Chk110.TabIndex = 200
        Me.Chk110.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk110.UseVisualStyleBackColor = True
        '
        'Chk109
        '
        Me.Chk109.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk109.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk109.ForeColor = System.Drawing.Color.Navy
        Me.Chk109.Location = New System.Drawing.Point(337, 2)
        Me.Chk109.Name = "Chk109"
        Me.Chk109.Size = New System.Drawing.Size(18, 18)
        Me.Chk109.TabIndex = 199
        Me.Chk109.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk109.UseVisualStyleBackColor = True
        '
        'Chk108
        '
        Me.Chk108.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk108.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk108.ForeColor = System.Drawing.Color.Navy
        Me.Chk108.Location = New System.Drawing.Point(319, 2)
        Me.Chk108.Name = "Chk108"
        Me.Chk108.Size = New System.Drawing.Size(18, 18)
        Me.Chk108.TabIndex = 198
        Me.Chk108.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk108.UseVisualStyleBackColor = True
        '
        'Chk107
        '
        Me.Chk107.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk107.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk107.ForeColor = System.Drawing.Color.Navy
        Me.Chk107.Location = New System.Drawing.Point(301, 2)
        Me.Chk107.Name = "Chk107"
        Me.Chk107.Size = New System.Drawing.Size(18, 18)
        Me.Chk107.TabIndex = 197
        Me.Chk107.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk107.UseVisualStyleBackColor = True
        '
        'Chk106
        '
        Me.Chk106.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk106.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk106.ForeColor = System.Drawing.Color.Navy
        Me.Chk106.Location = New System.Drawing.Point(283, 2)
        Me.Chk106.Name = "Chk106"
        Me.Chk106.Size = New System.Drawing.Size(18, 18)
        Me.Chk106.TabIndex = 196
        Me.Chk106.Text = " "
        Me.Chk106.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk106.UseVisualStyleBackColor = True
        '
        'Chk105
        '
        Me.Chk105.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk105.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk105.ForeColor = System.Drawing.Color.Navy
        Me.Chk105.Location = New System.Drawing.Point(265, 2)
        Me.Chk105.Name = "Chk105"
        Me.Chk105.Size = New System.Drawing.Size(18, 18)
        Me.Chk105.TabIndex = 195
        Me.Chk105.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk105.UseVisualStyleBackColor = True
        '
        'Chk104
        '
        Me.Chk104.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk104.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk104.ForeColor = System.Drawing.Color.Navy
        Me.Chk104.Location = New System.Drawing.Point(247, 2)
        Me.Chk104.Name = "Chk104"
        Me.Chk104.Size = New System.Drawing.Size(18, 18)
        Me.Chk104.TabIndex = 194
        Me.Chk104.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk104.UseVisualStyleBackColor = True
        '
        'Chk103
        '
        Me.Chk103.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk103.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk103.ForeColor = System.Drawing.Color.Navy
        Me.Chk103.Location = New System.Drawing.Point(229, 2)
        Me.Chk103.Name = "Chk103"
        Me.Chk103.Size = New System.Drawing.Size(18, 18)
        Me.Chk103.TabIndex = 193
        Me.Chk103.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk103.UseVisualStyleBackColor = True
        '
        'Chk102
        '
        Me.Chk102.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk102.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk102.ForeColor = System.Drawing.Color.Navy
        Me.Chk102.Location = New System.Drawing.Point(211, 2)
        Me.Chk102.Name = "Chk102"
        Me.Chk102.Size = New System.Drawing.Size(18, 18)
        Me.Chk102.TabIndex = 192
        Me.Chk102.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk102.UseVisualStyleBackColor = True
        '
        'Chk101
        '
        Me.Chk101.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Chk101.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Chk101.ForeColor = System.Drawing.Color.Navy
        Me.Chk101.Location = New System.Drawing.Point(193, 2)
        Me.Chk101.Name = "Chk101"
        Me.Chk101.Size = New System.Drawing.Size(18, 18)
        Me.Chk101.TabIndex = 191
        Me.Chk101.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Chk101.UseVisualStyleBackColor = True
        '
        'lblOkres1
        '
        Me.lblOkres1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOkres1.ForeColor = System.Drawing.Color.Navy
        Me.lblOkres1.Location = New System.Drawing.Point(20, 4)
        Me.lblOkres1.Name = "lblOkres1"
        Me.lblOkres1.Size = New System.Drawing.Size(173, 16)
        Me.lblOkres1.TabIndex = 190
        Me.lblOkres1.Text = "SPRZEDA¯ : RRRR.MM : 0..-11 . . . ."
        '
        'cmdHarmonogramWizyt
        '
        Me.cmdHarmonogramWizyt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdHarmonogramWizyt.ForeColor = System.Drawing.Color.Black
        Me.cmdHarmonogramWizyt.Image = CType(resources.GetObject("cmdHarmonogramWizyt.Image"), System.Drawing.Image)
        Me.cmdHarmonogramWizyt.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdHarmonogramWizyt.Location = New System.Drawing.Point(120, 478)
        Me.cmdHarmonogramWizyt.Name = "cmdHarmonogramWizyt"
        Me.cmdHarmonogramWizyt.Size = New System.Drawing.Size(24, 23)
        Me.cmdHarmonogramWizyt.TabIndex = 204
        Me.cmdHarmonogramWizyt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
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
        Me.dagKlienci.Size = New System.Drawing.Size(1127, 215)
        Me.dagKlienci.TabIndex = 206
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
        Me.Panel4.Location = New System.Drawing.Point(8, 141)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(1127, 265)
        Me.Panel4.TabIndex = 208
        '
        'KatalogKlientowAnaliza
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1142, 508)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.cmdDoExcela)
        Me.Controls.Add(Me.cmdHarmonogramWizyt)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.cmdOdswiez)
        Me.Controls.Add(Me.cmdWysylaj)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.cmdEdytuj)
        Me.Controls.Add(Me.cmdUsuñ)
        Me.Controls.Add(Me.cmdDodaj)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PictureBox1)
        Me.ForeColor = System.Drawing.Color.Azure
        Me.Location = New System.Drawing.Point(8, 8)
        Me.Name = "KatalogKlientowAnaliza"
        Me.Text = "Analiza Zakupów Klientów firmy PRYZMAT"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.TabSumy.ResumeLayout(False)
        Me.TabSprzedaz.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
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
    Dim dbSelectKlienciSorted As String = "SELECT * FROM Klienci ORDER BY IDENTKLIENTA"

    Dim dbUpdateKlienci As String = ""      'sUpdateSQLKlienci

    'Dim dbConnection As OleDb.OleDbConnection
    'Dim dbCommandSelect As OleDb.OleDbCommand
    'Dim dbCommandDelete As OleDb.OleDbCommand
    'Dim dbCommandUpdate As OleDb.OleDbCommand
    'Dim dbCommandInsert As OleDb.OleDbCommand
    'Dim dbDataAdapterKlienci As OleDb.OleDbDataAdapter

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
    Dim Okres1 As String = ""
    Dim Okres2 As String = ""




    '====================================================
    'UWAGA - aby to zadzialalo trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    '====================================================
    Private Sub KatalogKlientowAnaliza_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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



    Private Sub KatalogKlientowAnaliza_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        Me.Icon = My.Resources.MrJones
        '-------------------------------
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


        cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne")
        cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Iloœci")
        cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne")
        cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Wartoœci")
        cbxCoAnalizowac.Items.Add("Analizuj wartoœæ Mar¿y")
        cbxCoAnalizowac.Items.Add("Analizuj procent Mar¿y")

        PobierzParametry()

        'cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne"
        '2016.04.20
        'cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Wartoœci"
        '2019.07.14
        cbxCoAnalizowac.Text = "Analizuj wartoœæ Mar¿y"



        'If Mid(Par_DaneDoAnalizy, 7, 1) = "T" Then
        '    cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne"
        '    cbxCoAnalizowac.SelectedIndex = 0
        'End If
        'If Mid(Par_DaneDoAnalizy, 8, 1) = "T" Then
        '    cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne"
        '    cbxCoAnalizowac.SelectedIndex = 2
        'End If
        'If Mid(Par_DaneDoAnalizy, 9, 1) = "T" Then
        '    cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Iloœci"
        '    cbxCoAnalizowac.SelectedIndex = 1
        'End If
        'If Mid(Par_DaneDoAnalizy, 10, 1) = "T" Then
        '    cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Wartoœci"
        '    cbxCoAnalizowac.SelectedIndex = 3
        'End If




        '-------------------------------
        'Par_OstOkresAnalizy = MiesDoAnalizy(-1)
        'Par_DataAktAnalizy = SysData()
        Me.Text = "Analiza zakupów klientów (" &
                  IIf(Mid(Par_DaneDoAnalizy, 1, 1) = "T", "PRYZMAT ", "") &
                  IIf(Mid(Par_DaneDoAnalizy, 2, 1) = "T", "ORYGINALNE ", "") &
                  IIf(Mid(Par_DaneDoAnalizy, 3, 1) = "T", "NAJEM ", "") &
                  IIf(Mid(Par_DaneDoAnalizy, 4, 1) = "T", "STRONY ", "") &
                  IIf(Mid(Par_DaneDoAnalizy, 5, 1) = "T", "DRUKARKI ", "") &
                  IIf(Mid(Par_DaneDoAnalizy, 6, 1) = "T", "SKUP", "") &
                ")"

        'Par_OSTOkresAnalizy="RRRR-MM" & "RRRR-MM"
        'I okres analizy
        If Len(Trim(Mid(Par_OstOkresAnalizy, 1, 7))) = 0 Then
            Okres1 = Mid(SysData, 1, 7)
        Else
            Okres1 = Mid(Par_OstOkresAnalizy, 1, 7)
        End If

        'II okres analizy
        If Len(Trim(Mid(Par_OstOkresAnalizy, 8, 7))) = 0 Then
            Okres2 = MiesDoAnalizy(Okres1, -12)
        Else
            Okres2 = Mid(Par_OstOkresAnalizy, 8, 7)
        End If
        lblOkres1.Text = "SPRZEDA¯ : " & InnyMiesiac(Okres1, 0) & " .. " & InnyMiesiac(Okres1, -11)
        lblOkres2.Text = InnyMiesiac(Okres2, 0) & " .. " & InnyMiesiac(Okres2, -11)

        ToolTip1.SetToolTip(Me.Chk101, InnyMiesiac(Okres1, 0))
        ToolTip1.SetToolTip(Me.Chk102, InnyMiesiac(Okres1, -1))
        ToolTip1.SetToolTip(Me.Chk103, InnyMiesiac(Okres1, -2))
        ToolTip1.SetToolTip(Me.Chk104, InnyMiesiac(Okres1, -3))
        ToolTip1.SetToolTip(Me.Chk105, InnyMiesiac(Okres1, -4))
        ToolTip1.SetToolTip(Me.Chk106, InnyMiesiac(Okres1, -5))
        ToolTip1.SetToolTip(Me.Chk107, InnyMiesiac(Okres1, -6))
        ToolTip1.SetToolTip(Me.Chk108, InnyMiesiac(Okres1, -7))
        ToolTip1.SetToolTip(Me.Chk109, InnyMiesiac(Okres1, -8))
        ToolTip1.SetToolTip(Me.Chk110, InnyMiesiac(Okres1, -9))
        ToolTip1.SetToolTip(Me.Chk111, InnyMiesiac(Okres1, -10))
        ToolTip1.SetToolTip(Me.Chk112, InnyMiesiac(Okres1, -11))

        Chk101.Checked = (Mid(Par_MiesAnalizy1, 1, 1) = "T")
        Chk102.Checked = (Mid(Par_MiesAnalizy1, 2, 1) = "T")
        Chk103.Checked = (Mid(Par_MiesAnalizy1, 3, 1) = "T")
        Chk104.Checked = (Mid(Par_MiesAnalizy1, 4, 1) = "T")
        Chk105.Checked = (Mid(Par_MiesAnalizy1, 5, 1) = "T")
        Chk106.Checked = (Mid(Par_MiesAnalizy1, 6, 1) = "T")
        Chk107.Checked = (Mid(Par_MiesAnalizy1, 7, 1) = "T")
        Chk108.Checked = (Mid(Par_MiesAnalizy1, 8, 1) = "T")
        Chk109.Checked = (Mid(Par_MiesAnalizy1, 9, 1) = "T")
        Chk110.Checked = (Mid(Par_MiesAnalizy1, 10, 1) = "T")
        Chk111.Checked = (Mid(Par_MiesAnalizy1, 11, 1) = "T")
        Chk112.Checked = (Mid(Par_MiesAnalizy1, 12, 1) = "T")

        ToolTip1.SetToolTip(Me.Chk201, InnyMiesiac(Okres2, 0))
        ToolTip1.SetToolTip(Me.Chk202, InnyMiesiac(Okres2, -1))
        ToolTip1.SetToolTip(Me.Chk203, InnyMiesiac(Okres2, -2))
        ToolTip1.SetToolTip(Me.Chk204, InnyMiesiac(Okres2, -3))
        ToolTip1.SetToolTip(Me.Chk205, InnyMiesiac(Okres2, -4))
        ToolTip1.SetToolTip(Me.Chk206, InnyMiesiac(Okres2, -5))
        ToolTip1.SetToolTip(Me.Chk207, InnyMiesiac(Okres2, -6))
        ToolTip1.SetToolTip(Me.Chk208, InnyMiesiac(Okres2, -7))
        ToolTip1.SetToolTip(Me.Chk209, InnyMiesiac(Okres2, -8))
        ToolTip1.SetToolTip(Me.Chk210, InnyMiesiac(Okres2, -9))
        ToolTip1.SetToolTip(Me.Chk211, InnyMiesiac(Okres2, -10))
        ToolTip1.SetToolTip(Me.Chk212, InnyMiesiac(Okres2, -11))

        Chk201.Checked = (Mid(Par_MiesAnalizy2, 1, 1) = "T")
        Chk202.Checked = (Mid(Par_MiesAnalizy2, 2, 1) = "T")
        Chk203.Checked = (Mid(Par_MiesAnalizy2, 3, 1) = "T")
        Chk204.Checked = (Mid(Par_MiesAnalizy2, 4, 1) = "T")
        Chk205.Checked = (Mid(Par_MiesAnalizy2, 5, 1) = "T")
        Chk206.Checked = (Mid(Par_MiesAnalizy2, 6, 1) = "T")
        Chk207.Checked = (Mid(Par_MiesAnalizy2, 7, 1) = "T")
        Chk208.Checked = (Mid(Par_MiesAnalizy2, 8, 1) = "T")
        Chk209.Checked = (Mid(Par_MiesAnalizy2, 9, 1) = "T")
        Chk210.Checked = (Mid(Par_MiesAnalizy2, 10, 1) = "T")
        Chk211.Checked = (Mid(Par_MiesAnalizy2, 11, 1) = "T")
        Chk212.Checked = (Mid(Par_MiesAnalizy2, 12, 1) = "T")


        chkPryzmatAtrament.Checked = Mid(Par_DaneDoAnalizy, 1, 1) = "T" Or Mid(Par_DaneDoAnalizy, 1, 1) = "A"
        chkPryzmatTonery.Checked = Mid(Par_DaneDoAnalizy, 1, 1) = "T" Or Mid(Par_DaneDoAnalizy, 1, 1) = "L"
        chkOrygAtrament.Checked = Mid(Par_DaneDoAnalizy, 2, 1) = "T" Or Mid(Par_DaneDoAnalizy, 2, 1) = "A"
        chkOrygTonery.Checked = Mid(Par_DaneDoAnalizy, 2, 1) = "T" Or Mid(Par_DaneDoAnalizy, 2, 1) = "L"
        chkNajem.Checked = Mid(Par_DaneDoAnalizy, 3, 1) = "T"
        chkStrony.Checked = Mid(Par_DaneDoAnalizy, 3, 1) = "T"
        chkDrukarki.Checked = Mid(Par_DaneDoAnalizy, 3, 1) = "T"
        chkSkup.Checked = Mid(Par_DaneDoAnalizy, 3, 1) = "T"

        'If Mid(Par_DaneDoAnalizy, 7, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne"
        'If Mid(Par_DaneDoAnalizy, 8, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne"
        'If Mid(Par_DaneDoAnalizy, 9, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Iloœci"
        'If Mid(Par_DaneDoAnalizy, 10, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Wartoœci"
        'If Mid(Par_DaneDoAnalizy, 11, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj wartoœæ Mar¿y"
        'If Mid(Par_DaneDoAnalizy, 12, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj procent Mar¿y"

        cbxCoAnalizowac.Text = "Analizuj wartoœæ Mar¿y"


        'Par_DaneDoAnalizy = IIf(chkPryzmat.Checked, "T", IIf(chkPryzmatAtrament.Checked, "A", IIf(chkPryzmatTonery.Checked, "L", "N"))) &
        '                    IIf(chkOryg.Checked, "T", IIf(chkOrygAtrament.Checked, "A", IIf(chkOrygTonery.Checked, "L", "N"))) &
        '                    IIf(chkNajem.Checked, "T", "N") &
        '                    IIf(chkStrony.Checked, "T", "N") &
        '                    IIf(chkDrukarki.Checked, "T", "N") &
        '                    IIf(chkSkup.Checked, "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Iloœci", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Wartoœci", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj wartoœæ Mar¿y", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj procent Mar¿y", "T", "N")



        '-------------------------------
        _CoMamRobic = oCoMamRobic
        StartFormy = True
        PrzygotujAnalize()
        WyliczSumyDlaTabeli()
        StartFormy = False
        'dagKlienci.Refresh()
        '-------------------------------
        'InicjujFiltryColumnDGV(Me.panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        'dagKlienci_CurrentCellChanged(sender, e)
        '-------------------------------
        'RysujGradientWPanelu(Panel2)
        'RysujGradientWPanelu(Panel3)
        'RysujGradientWPanelu(Panel4)
        'RysujGradientWPanelu(Panel5)
        Me.WindowState = FormWindowState.Maximized
        InicjujFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        'RowHeightSetter.MultiHightsView()
        dagKlienci_CurrentCellChanged(sender, e)
        '--------------------------
        'zegar do scrollingu poziomego
        HorizScrollTimer.Interval = 2 * 1000
        HorizScrollTimer.Enabled = True
    End Sub














    Private Sub PrzygotujAnalize()
        Dim SzerKolumn As Integer = 60
        oCzasStart = DateTime.UtcNow

        '---------------------------------------

        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne")
        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Iloœci")
        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne")
        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Wartoœci")
        'cbxCoAnalizowac.Items.Add("Analizuj wartoœæ Mar¿y")
        'cbxCoAnalizowac.Items.Add("Analizuj procent Mar¿y")

        Dim AnalIlosciowo As Boolean = False
        Dim AnalWartMarzy As Boolean = False
        Dim AnalProcMarzy As Boolean = False
        Dim SumujAktywnosci As Boolean = False

        Select Case cbxCoAnalizowac.Text
            Case "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne"
                AnalIlosciowo = True
                SumujAktywnosci = True
                AnalWartMarzy = False
                AnalProcMarzy = False
            Case "Analizuj Iloœciowo-Sumuj Iloœci"
                AnalIlosciowo = True
                SumujAktywnosci = False
                AnalWartMarzy = False
                AnalProcMarzy = False
            Case "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne"
                AnalIlosciowo = False
                SumujAktywnosci = True
                AnalWartMarzy = False
                AnalProcMarzy = False
            Case "Analizuj Wartoœciowo-Sumuj Wartoœci"
                AnalIlosciowo = False
                SumujAktywnosci = False
                AnalWartMarzy = False
                AnalProcMarzy = False
            Case "Analizuj wartoœæ Mar¿y"
                AnalIlosciowo = False
                SumujAktywnosci = False
                AnalWartMarzy = True
                AnalProcMarzy = False
            Case "Analizuj procent Mar¿y"
                AnalIlosciowo = False
                SumujAktywnosci = False
                AnalWartMarzy = False
                AnalProcMarzy = True
        End Select



        dbSelectKlienci = "SELECT " &
                                             "K.IDENTKLIENTA AS NRKARTY, " &
                                             "K.NIP," &
                                             "K.NRTOFITXT," &
                                             "K.KARTAPAYBACK," &
                                             "K.NRYKARTPAYBACK," &
                                             "K.NAZWA1 AS NAZWAKLIENTA," &
                                             "K.MIEJSCOWOSC," &
                                             "K.KODPOCZTOWY," &
                                             "K.ULICA," &
                                             "K.NUMNRDOMU, "

        If ParL_DataBaseType = "MS ACCESS" Then
            dbSelectKlienci &= "IIF(K.NUMNRDOMU MOD 2=0,'True','False') AS PARZYSTE, "
        Else
            dbSelectKlienci &= "CASE WHEN K.NUMNRDOMU%2=0 THEN CAST('True' AS BIT) ELSE CAST('False' AS BIT) END AS PARZYSTE, "
        End If

        dbSelectKlienci &=
                                             "K.NRDOMU," &
                                             "K.IDDOMU," &
                                             "K.REJKLIENTA AS REJONKLIENTA," &
                                             "K.TELEFON," &
                                             "K.FAX," &
                                             "K.EMAIL," &
                                               "K.BRANZA," &
                                               "K.PODBRANZA," &
                                               "K.LICZBAZATRUDNIONYCH," &
                                               "K.LICZBAPRACZDALNIE," &
                                       "K.WYKAZURZADZEN," &
                                         "K.SPOSOBWYBORUDOSTAWCY," &
                                         "K.ZAINSTALOWANOMONITOR," &
                                         "K.LICZBAURZADZEN," &
                                         "K.LICZBAWYDRUKOW," &
                                         "K.POTENCJALDRUKU," &
                                       "K.AKTZAKUPMATEKSP," &
                                       "K.AKTROZLICZASTRONYDRUKU," &
                                       "K.AKTKORZYSTAZNAJMU," &
                                       "K.ZAINTROZLICZANIEMSTRONDRUKU," &
                                       "K.ZAINTKORZYSTANIEMZNAJMU," &
                                         "K.DATAWERYFSEGMENTACJI," &
                                     "K.PLANDLUGOOKRESOWY," &
                                     "K.PLANKROTKOOKRESOWY,"



        If SumujAktywnosci Then
            dbSelectKlienci &=
                                         "A.WARTZA00 AS ILOSCAA," &
                                         "A.ILOSZA00 AS ILOSCBB,"
        Else
            If AnalIlosciowo Then
                dbSelectKlienci &=
                    "A.ILOSZABM + A.ILOSZA01 + A.ILOSZA02 + A.ILOSZA03 + A.ILOSZA04 + A.ILOSZA05 + A.ILOSZA06 + A.ILOSZA07 + A.ILOSZA08 + A.ILOSZA09 + A.ILOSZA10 + A.ILOSZA11 AS ILOSCAA," &
                    "A.ILOSZA12 + A.ILOSZA13 + A.ILOSZA14 + A.ILOSZA15 + A.ILOSZA16 + A.ILOSZA17 + A.ILOSZA18 + A.ILOSZA19 + A.ILOSZA20 + A.ILOSZA21 + A.ILOSZA22 + A.ILOSZA23 AS ILOSCBB,"
            ElseIf AnalWartMarzy Then
                dbSelectKlienci &=
                    "A.MARZABM + A.MARZA01 + A.MARZA02 + A.MARZA03 + A.MARZA04 + A.MARZA05 + A.MARZA06 + A.MARZA07 + A.MARZA08 + A.MARZA09 + A.MARZA10 + A.MARZA11 AS ILOSCAA," &
                    "A.MARZA12 + A.MARZA13 + A.MARZA14 + A.MARZA15 + A.MARZA16 + A.MARZA17 + A.MARZA18 + A.MARZA19 + A.MARZA20 + A.MARZA21 + A.MARZA22 + A.MARZA23 AS ILOSCBB,"
            ElseIf AnalProcMarzy Then
                dbSelectKlienci &=
                    "(A.PROCMBM + A.PROCM01 + A.PROCM02 + A.PROCM03 + A.PROCM04 + A.PROCM05 + A.PROCM06 + A.PROCM07 + A.PROCM08 + A.PROCM09 + A.PROCM10 + A.PROCM11) /12 AS ILOSCAA," &
                    "(A.PROCM12 + A.PROCM13 + A.PROCM14 + A.PROCM15 + A.PROCM16 + A.PROCM17 + A.PROCM18 + A.PROCM19 + A.PROCM20 + A.PROCM21 + A.PROCM22 + A.PROCM23) / 12 AS ILOSCBB,"
            Else    'wartosciowo
                dbSelectKlienci &=
                    "A.WARTZABM + A.WARTZA01 + A.WARTZA02 + A.WARTZA03 + A.WARTZA04 + A.WARTZA05 + A.WARTZA06 + A.WARTZA07 + A.WARTZA08 + A.WARTZA09 + A.WARTZA10 + A.WARTZA11 AS ILOSCAA," &
                    "A.WARTZA12 + A.WARTZA13 + A.WARTZA14 + A.WARTZA15 + A.WARTZA16 + A.WARTZA17 + A.WARTZA18 + A.WARTZA19 + A.WARTZA20 + A.WARTZA21 + A.WARTZA22 + A.WARTZA23 AS ILOSCBB,"
            End If
        End If

        If AnalIlosciowo Then
            'Tylko Ilosc
            SzerKolumn = 60
            dbSelectKlienci &=
                                                 "A.ILOSZABM AS ILOSCBM," &
                                                   "A.ILOSZA01 AS ILZAK01," &
                                                   "A.ILOSZA02 AS ILZAK02," &
                                                   "A.ILOSZA03 AS ILZAK03," &
                                                   "A.ILOSZA04 AS ILZAK04," &
                                                   "A.ILOSZA05 AS ILZAK05," &
                                                   "A.ILOSZA06 AS ILZAK06," &
                                                   "A.ILOSZA07 AS ILZAK07," &
                                                   "A.ILOSZA08 AS ILZAK08," &
                                                   "A.ILOSZA09 AS ILZAK09," &
                                                   "A.ILOSZA10 AS ILZAK10," &
                                                   "A.ILOSZA11 AS ILZAK11," &
                                                   "A.ILOSZA12 AS ILZAK12," &
                                                 "A.ILOSZA13 AS ILZAK13," &
                                                 "A.ILOSZA14 AS ILZAK14," &
                                                 "A.ILOSZA15 AS ILZAK15," &
                                                 "A.ILOSZA16 AS ILZAK16," &
                                                 "A.ILOSZA17 AS ILZAK17," &
                                                 "A.ILOSZA18 AS ILZAK18," &
                                                 "A.ILOSZA19 AS ILZAK19," &
                                                 "A.ILOSZA20 AS ILZAK20," &
                                                 "A.ILOSZA21 AS ILZAK21," &
                                                 "A.ILOSZA22 AS ILZAK22," &
                                                 "A.ILOSZA23 AS ILZAK23," &
                                                 "A.ILOSZA24 AS ILZAK24,"

        ElseIf AnalWartMarzy Then
            'marza
            SzerKolumn = 60
            dbSelectKlienci &=
                                                 "A.MARZABM AS ILOSCBM," &
                                                   "A.MARZA01 AS ILZAK01," &
                                                   "A.MARZA02 AS ILZAK02," &
                                                   "A.MARZA03 AS ILZAK03," &
                                                   "A.MARZA04 AS ILZAK04," &
                                                   "A.MARZA05 AS ILZAK05," &
                                                   "A.MARZA06 AS ILZAK06," &
                                                   "A.MARZA07 AS ILZAK07," &
                                                   "A.MARZA08 AS ILZAK08," &
                                                   "A.MARZA09 AS ILZAK09," &
                                                   "A.MARZA10 AS ILZAK10," &
                                                   "A.MARZA11 AS ILZAK11," &
                                                   "A.MARZA12 AS ILZAK12," &
                                                 "A.MARZA13 AS ILZAK13," &
                                                 "A.MARZA14 AS ILZAK14," &
                                                 "A.MARZA15 AS ILZAK15," &
                                                 "A.MARZA16 AS ILZAK16," &
                                                 "A.MARZA17 AS ILZAK17," &
                                                 "A.MARZA18 AS ILZAK18," &
                                                 "A.MARZA19 AS ILZAK19," &
                                                 "A.MARZA20 AS ILZAK20," &
                                                 "A.MARZA21 AS ILZAK21," &
                                                 "A.MARZA22 AS ILZAK22," &
                                                 "A.MARZA23 AS ILZAK23," &
                                                 "A.MARZA24 AS ILZAK24,"

        ElseIf AnalProcMarzy Then
            'procent marzy
            SzerKolumn = 60
            dbSelectKlienci &=
                                                 "A.PROCMBM AS ILOSCBM," &
                                                   "A.PROCM01 AS ILZAK01," &
                                                   "A.PROCM02 AS ILZAK02," &
                                                   "A.PROCM03 AS ILZAK03," &
                                                   "A.PROCM04 AS ILZAK04," &
                                                   "A.PROCM05 AS ILZAK05," &
                                                   "A.PROCM06 AS ILZAK06," &
                                                   "A.PROCM07 AS ILZAK07," &
                                                   "A.PROCM08 AS ILZAK08," &
                                                   "A.PROCM09 AS ILZAK09," &
                                                   "A.PROCM10 AS ILZAK10," &
                                                   "A.PROCM11 AS ILZAK11," &
                                                   "A.PROCM12 AS ILZAK12," &
                                                 "A.PROCM13 AS ILZAK13," &
                                                 "A.PROCM14 AS ILZAK14," &
                                                 "A.PROCM15 AS ILZAK15," &
                                                 "A.PROCM16 AS ILZAK16," &
                                                 "A.PROCM17 AS ILZAK17," &
                                                 "A.PROCM18 AS ILZAK18," &
                                                 "A.PROCM19 AS ILZAK19," &
                                                 "A.PROCM20 AS ILZAK20," &
                                                 "A.PROCM21 AS ILZAK21," &
                                                 "A.PROCM22 AS ILZAK22," &
                                                 "A.PROCM23 AS ILZAK23," &
                                                 "A.PROCM24 AS ILZAK24,"

        Else    'wartosciowo
            'Tylko Wartosc
            SzerKolumn = 60
            dbSelectKlienci &=
                                                 "A.WARTZABM AS ILOSCBM," &
                                                   "A.WARTZA01 AS ILZAK01," &
                                                   "A.WARTZA02 AS ILZAK02," &
                                                   "A.WARTZA03 AS ILZAK03," &
                                                   "A.WARTZA04 AS ILZAK04," &
                                                   "A.WARTZA05 AS ILZAK05," &
                                                   "A.WARTZA06 AS ILZAK06," &
                                                   "A.WARTZA07 AS ILZAK07," &
                                                   "A.WARTZA08 AS ILZAK08," &
                                                   "A.WARTZA09 AS ILZAK09," &
                                                   "A.WARTZA10 AS ILZAK10," &
                                                   "A.WARTZA11 AS ILZAK11," &
                                                   "A.WARTZA12 AS ILZAK12," &
                                                 "A.WARTZA13 AS ILZAK13," &
                                                 "A.WARTZA14 AS ILZAK14," &
                                                 "A.WARTZA15 AS ILZAK15," &
                                                 "A.WARTZA16 AS ILZAK16," &
                                                 "A.WARTZA17 AS ILZAK17," &
                                                 "A.WARTZA18 AS ILZAK18," &
                                                 "A.WARTZA19 AS ILZAK19," &
                                                 "A.WARTZA20 AS ILZAK20," &
                                                 "A.WARTZA21 AS ILZAK21," &
                                                 "A.WARTZA22 AS ILZAK22," &
                                                 "A.WARTZA23 AS ILZAK23," &
                                                 "A.WARTZA24 AS ILZAK24,"

        End If


        dbSelectKlienci &=
                                               "K.RODZSPRZETUL," &
                                               "K.RODZSPRZETUA," &
                                               "K.RODZSPRZETUT," &
                                               "K.OFERTATZLOZENIA," &
                                               "K.OFERTAPLIK," &
                                               "K.PKONTAKT," &
                                               "K.SKUPUWAGI AS PROMOTORUWAGI," &
                                               "K.SPRZEDOPIEKUN AS OPIEKUN," &
                                               "K.SPRZEDOKONTAKTR AS OPIEKUNOSTKONTAKTROK," &
                                               "K.SPRZEDOKONTAKTT AS OPIEKUNOSTKONTAKTTYDZ," &
                                               "K.SPRZEDOKONTAKTD AS OPIEKUNOSTKONTAKTDZIEN," &
                                               "K.SPRZEDNKONTAKTR AS OPIEKUNKOLKONTAKTROK," &
                                               "K.SPRZEDNKONTAKTT AS OPIEKUNKOLKONTAKTTYDZ," &
                                               "K.SPRZEDNKONTAKTD AS OPIEKUNKOLKONTAKTDZIEN," &
                                               "K.SPRZEDUWAGI AS OPIEKUNUWAGI," &
                                               "K.KTOWPISAL," &
                                               "K.UWAGI," &
                                               "K.MARKER, " &
                                               "K.MARKERP, " &
                                               "K.AKTYWNY, " &
                                           "K.ZAKUPPOPROK, " &
                                           "K.USLUGIDOD, " &
                                           "K.RZURAP01, " &
                                           "K.RZURAP02, " &
                                           "K.RZURAP03, " &
                                           "K.RZURAP04, " &
                                           "K.RZURAP05, " &
                                           "K.RZURAP06, " &
                                           "K.RZURAP07, " &
                                           "K.RZURAP08, " &
                                           "K.RZURAP09, "

        dbSelectKlienci &=
                                       "K.WERSJA " &
                                    "FROM Klienci AS K LEFT JOIN AnalizyZakupu AS A " &
                                        "ON K.IDENTKLIENTA=A.IDENTKLIENTA "

        dbSelectKlienciSorted = dbSelectKlienci & " ORDER BY NRKARTY ASC"




















        DataSetKlienci = New DataSet
        DataViewKlienci = New DataView
        DataSetKlienci.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
        Else
            sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelectKlienci = New SqlClient.SqlCommand(dbSelectKlienciSorted, sqlConnectionKlienci)
            sqlCommandDeleteKlienci = New SqlClient.SqlCommand(sDeleteSQLKlienci, sqlConnectionKlienci)
            sqlCommandUpdateKlienci = New SqlClient.SqlCommand(sUpdateSQLKlienci, sqlConnectionKlienci)
            sqlCommandInsertKlienci = New SqlClient.SqlCommand(sInsertSQLKlienci, sqlConnectionKlienci)
            sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliKlienci As System.Data.Common.DataTableMapping
            MapowanieTabeliKlienci = sqlDataAdapterKlienci.TableMappings.Add("Table", "TABELA_Klienci")

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
                    'dbDataAdapterKlienci.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
                    'dbDataAdapterKlienci.Fill(DataSetKlienci)
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
                'dagKlienci.Dock = DockStyle.Fill
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
                'dagKlienci.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
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
                'dagKlienci.DataSource = DataViewKlienci
                '-----------------------------------

                'If dagKlienci.Columns.Count > 0 Then
                '    Dim ic As Integer = 0
                '    For ic = dagKlienci.Columns.Count - 1 To 0
                '        dagKlienci.Columns.RemoveAt(ic)
                '    Next
                'End If


                dagKlienci.Columns.Clear()
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(0).ColumnName, "Nr karty", 50, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(1).ColumnName, "NIP klienta", 100, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(2).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)

                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(3).ColumnName, "PayBack", 60, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(4).ColumnName, "Nry Kart Payback", 100, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(5).ColumnName, "Nazwa klienta", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(6).ColumnName, "Miejscowoœæ", 150, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(7).ColumnName, "Kod poczt.", 60, DataGridViewContentAlignment.MiddleCenter, True)
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
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(32).ColumnName, "Data weryfikacji segmentacji", 80, DataGridViewContentAlignment.MiddleCenter, True)

                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(33).ColumnName, "Plan d³ugookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(34).ColumnName, "Plan krótkookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)


                '----------------------------------------------------
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(35).ColumnName, "Okres I", 50, DataGridViewContentAlignment.MiddleRight, "F" & IIf(AnalIlosciowo, 0, 2), True)
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(36).ColumnName, "Okres II", 50, DataGridViewContentAlignment.MiddleRight, "F" & IIf(AnalIlosciowo, 0, 2), True)

                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(37).ColumnName, InnyMiesiac(Okres1, 0), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk101.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(38).ColumnName, InnyMiesiac(Okres1, -1), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk102.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(39).ColumnName, InnyMiesiac(Okres1, -2), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk103.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(40).ColumnName, InnyMiesiac(Okres1, -3), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk104.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(41).ColumnName, InnyMiesiac(Okres1, -4), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk105.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(42).ColumnName, InnyMiesiac(Okres1, -5), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk106.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(43).ColumnName, InnyMiesiac(Okres1, -6), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk107.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(44).ColumnName, InnyMiesiac(Okres1, -7), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk108.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(45).ColumnName, InnyMiesiac(Okres1, -8), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk109.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(46).ColumnName, InnyMiesiac(Okres1, -9), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk110.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(47).ColumnName, InnyMiesiac(Okres1, -10), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk111.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(48).ColumnName, InnyMiesiac(Okres1, -11), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk112.Checked, True, False))

                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(49).ColumnName, InnyMiesiac(Okres2, 0), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk201.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(50).ColumnName, InnyMiesiac(Okres2, -1), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk202.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(51).ColumnName, InnyMiesiac(Okres2, -2), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk203.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, InnyMiesiac(Okres2, -3), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk204.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, InnyMiesiac(Okres2, -4), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk205.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, InnyMiesiac(Okres2, -5), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk206.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(55).ColumnName, InnyMiesiac(Okres2, -6), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk207.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(56).ColumnName, InnyMiesiac(Okres2, -7), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk208.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(57).ColumnName, InnyMiesiac(Okres2, -8), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk209.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(58).ColumnName, InnyMiesiac(Okres2, -9), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk210.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(59).ColumnName, InnyMiesiac(Okres2, -10), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk211.Checked, True, False))
                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(60).ColumnName, InnyMiesiac(Okres2, -11), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk212.Checked, True, False))

                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(61).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "F" & IIf(AnalIlosciowo, 0, 2), False)  'mies 24
                '----------------------------------------------------




                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(62).ColumnName, "Rodz.Sprzêtu L", 100, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(63).ColumnName, "Rodz.Sprzêtu A", 100, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(64).ColumnName, "Rodz.Sprzêtu A3", 100, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(65).ColumnName, "Oferta z³o¿enia", 40, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(66).ColumnName, "Plik z ofert¹", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(67).ColumnName, "I-kontakt", 100, DataGridViewContentAlignment.MiddleLeft, True)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(68).ColumnName, "Opis potencja³u", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(69).ColumnName, "Opiekun", 100, DataGridViewContentAlignment.MiddleCenter, True)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(70).ColumnName, "Opiekun Ost.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(71).ColumnName, "Opiekun Ost.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(72).ColumnName, "Opiekun Ost.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(73).ColumnName, "Opiekun Kol.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(74).ColumnName, "Opiekun Kol.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(75).ColumnName, "Opiekun Kol.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(76).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(77).ColumnName, "Kto wpisa³ ?", 100, DataGridViewContentAlignment.MiddleCenter, True)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(78).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleCenter, True)

                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(79).ColumnName, "Marker  ", 50, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(80).ColumnName, "Marker P.  ", 50, True)
                LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(81).ColumnName, "Aktywny  ", 50, True)

                NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(82).ColumnName, "Zakup pop.rok   ", 60, DataGridViewContentAlignment.MiddleRight, "F2", False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(83).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(84).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(85).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(86).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(87).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(88).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(89).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(90).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(91).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(92).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

                TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(93).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)


                Me.Cursor = Cursors.WaitCursor
                'dagKlienci.DataSource = Nothing
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
            stbKlienci.Panels.Clear()

            'dodaj StatusBar na koncu
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
            If dagKlienci.CurrentCell Is Nothing Then
                stbKlienci.Panels(1).Text = "Wi: "
            Else
                stbKlienci.Panels(1).Text = "Wi: " & dagKlienci.CurrentCell.RowIndex.ToString
            End If

            stbKlienci.Panels.Add("stbPanelKolumna")
            stbKlienci.Panels(2).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(2).Width = 80
            stbKlienci.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
            If dagKlienci.CurrentCell Is Nothing Then
                stbKlienci.Panels(2).Text = "Ko: "
            Else
                stbKlienci.Panels(2).Text = "Ko: " & dagKlienci.CurrentCell.ColumnIndex.ToString
            End If


            stbKlienci.Panels.Add("stbPanelSort")
            stbKlienci.Panels(3).AutoSize = StatusBarPanelAutoSize.None
            stbKlienci.Panels(3).Width = 150
            stbKlienci.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
            stbKlienci.Panels(3).Text = "Sort: "

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
            stbFiltry.Panels.Clear()

            'dodaj StatusBar na koncu
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
        ''--------------------------------------------------
        'StartFormy = False
        'dagKlienci_CurrentCellChanged(Me.panel4, Nothing)
        'InicjujFiltryColumnDGV(Me.panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        'PokazSumy()
        ''--------------------------------------------------
        'If stbKlienci.Panels.Count > 6 Then
        oCzasStop = DateTime.UtcNow
        stbKlienci.Panels(5).Alignment = HorizontalAlignment.Right
        stbKlienci.Panels(5).Text = "Czas ini.(milisek) : " & (oCzasStop - oCzasStart).TotalMilliseconds.ToString
        'End If
    End Sub


    Private Sub KatalogKlientowAnaliza_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        'RysujGradient(Me)
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

                Else
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
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(32).ColumnName, "Data weryfikacji segmentacji", 80, DataGridViewContentAlignment.MiddleCenter, True)

                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(33).ColumnName, "Plan d³ugookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(34).ColumnName, "Plan krótkookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)


                    ''----------------------------------------------------
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(35).ColumnName, "Okres I", 50, DataGridViewContentAlignment.MiddleRight, "F" & IIf(AnalIlosciowo, 0, 2), True)
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(36).ColumnName, "Okres II", 50, DataGridViewContentAlignment.MiddleRight, "F" & IIf(AnalIlosciowo, 0, 2), True)

                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(37).ColumnName, InnyMiesiac(Okres1, 0), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk101.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(38).ColumnName, InnyMiesiac(Okres1, -1), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk102.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(39).ColumnName, InnyMiesiac(Okres1, -2), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk103.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(40).ColumnName, InnyMiesiac(Okres1, -3), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk104.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(41).ColumnName, InnyMiesiac(Okres1, -4), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk105.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(42).ColumnName, InnyMiesiac(Okres1, -5), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk106.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(43).ColumnName, InnyMiesiac(Okres1, -6), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk107.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(44).ColumnName, InnyMiesiac(Okres1, -7), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk108.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(45).ColumnName, InnyMiesiac(Okres1, -8), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk109.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(46).ColumnName, InnyMiesiac(Okres1, -9), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk110.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(47).ColumnName, InnyMiesiac(Okres1, -10), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk111.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(48).ColumnName, InnyMiesiac(Okres1, -11), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk112.Checked, True, False))

                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(49).ColumnName, InnyMiesiac(Okres2, 0), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk201.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(50).ColumnName, InnyMiesiac(Okres2, -1), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk202.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(51).ColumnName, InnyMiesiac(Okres2, -2), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk203.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, InnyMiesiac(Okres2, -3), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk204.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, InnyMiesiac(Okres2, -4), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk205.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, InnyMiesiac(Okres2, -5), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk206.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(55).ColumnName, InnyMiesiac(Okres2, -6), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk207.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(56).ColumnName, InnyMiesiac(Okres2, -7), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk208.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(57).ColumnName, InnyMiesiac(Okres2, -8), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk209.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(58).ColumnName, InnyMiesiac(Okres2, -9), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk210.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(59).ColumnName, InnyMiesiac(Okres2, -10), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk211.Checked, True, False))
                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(60).ColumnName, InnyMiesiac(Okres2, -11), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk212.Checked, True, False))

                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(61).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "F" & IIf(AnalIlosciowo, 0, 2), False)  'mies 24
                    ''----------------------------------------------------




                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(62).ColumnName, "Rodz.Sprzêtu L", 100, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(63).ColumnName, "Rodz.Sprzêtu A", 100, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(64).ColumnName, "Rodz.Sprzêtu A3", 100, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(65).ColumnName, "Oferta z³o¿enia", 40, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(66).ColumnName, "Plik z ofert¹", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(67).ColumnName, "I-kontakt", 100, DataGridViewContentAlignment.MiddleLeft, True)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(68).ColumnName, "Opis potencja³u", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(69).ColumnName, "Opiekun", 100, DataGridViewContentAlignment.MiddleCenter, True)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(70).ColumnName, "Opiekun Ost.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(71).ColumnName, "Opiekun Ost.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(72).ColumnName, "Opiekun Ost.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(73).ColumnName, "Opiekun Kol.kontakt rok", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(74).ColumnName, "Opiekun Kol.kontakt tydz.", 150, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(75).ColumnName, "Opiekun Kol.kontakt dz.", 150, DataGridViewContentAlignment.MiddleCenter, True)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(76).ColumnName, "Opiekun Uwagi", 200, DataGridViewContentAlignment.MiddleLeft, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(77).ColumnName, "Kto wpisa³ ?", 100, DataGridViewContentAlignment.MiddleCenter, True)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(78).ColumnName, "Uwagi", 200, DataGridViewContentAlignment.MiddleCenter, True)

                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(79).ColumnName, "Marker  ", 50, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(80).ColumnName, "Marker P.  ", 50, True)
                    'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(81).ColumnName, "Aktywny  ", 50, True)

                    'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(82).ColumnName, "Zakup pop.rok   ", 60, DataGridViewContentAlignment.MiddleRight, "F2", False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(83).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(84).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(85).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(86).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(87).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(88).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(89).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(90).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(91).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)
                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(92).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

                    'TxtColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(93).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)


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

                    oRodzSprzetuLKli = GetLogField(dagKlienci, 62)
                    oRodzSprzetuAKli = GetLogField(dagKlienci, 63)
                    oRodzSprzetuTKli = GetLogField(dagKlienci, 64)
                    lblRodzaj.Text = IIf(oRodzSprzetuLKli, "L", "") & IIf(oRodzSprzetuAKli, "A", "") & IIf(oRodzSprzetuTKli, "A3", "")

                    lblKod.Text = GetTxtField(dagKlienci, 7)
                    lblMiejscowosc.Text = GetTxtField(dagKlienci, 6)
                    lblUlica.Text = GetTxtField(dagKlienci, 8)
                    lblNrDomu.Text = GetIntField(dagKlienci, 9).ToString("F0")
                    lblIdDomu.Text = GetTxtField(dagKlienci, 12)

                    lblTelefon.Text = GetTxtField(dagKlienci, 14)
                    lbleMail.Text = GetTxtField(dagKlienci, 16)     'osoba kontaktowa

                    lblOpiekun.Text = GetTxtField(dagKlienci, 75)
                    lblOstKontakt.Text = "Tydzieñ " & Trim(GetTxtField(dagKlienci, 77)) & " / " & GetTxtField(dagKlienci, 76)
                    lblNastKontakt.Text = "Tydzieñ " & Trim(GetTxtField(dagKlienci, 80)) & " / " & GetTxtField(dagKlienci, 79)
                    lblTransakcje.Text = OstTransakcje(lblIdent.Text, True)
                    lblPotencjal.Text = GetTxtField(dagKlienci, 74)
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
            'PokazSumy(DataViewKlienci)
        End If
    End Sub
    Dim OstatniFiltrKlienci As String = ""
    Private Sub txtFiXX_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
        OstatniFiltrKlienci = sender.name
    End Sub
    Private Sub txtFiXX_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
        If Not OstatniFiltrKlienci = sender.name Then
            PokazSumy(DataViewKlienci)
        End If
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

            ''----------------------------------------------------
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(29).ColumnName, "Plan d³ugookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(30).ColumnName, "Plan krótkookres.", 60, DataGridViewContentAlignment.MiddleRight, "F0", True)


            ''----------------------------------------------------
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(35).ColumnName, "Okres I", 50, DataGridViewContentAlignment.MiddleRight, "F" & IIf(AnalIlosciowo, 0, 2), True)
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(36).ColumnName, "Okres II", 50, DataGridViewContentAlignment.MiddleRight, "F" & IIf(AnalIlosciowo, 0, 2), True)

            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(37).ColumnName, InnyMiesiac(Okres1, 0), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk101.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(38).ColumnName, InnyMiesiac(Okres1, -1), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk102.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(39).ColumnName, InnyMiesiac(Okres1, -2), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk103.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(40).ColumnName, InnyMiesiac(Okres1, -3), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk104.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(41).ColumnName, InnyMiesiac(Okres1, -4), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk105.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(42).ColumnName, InnyMiesiac(Okres1, -5), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk106.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(43).ColumnName, InnyMiesiac(Okres1, -6), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk107.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(44).ColumnName, InnyMiesiac(Okres1, -7), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk108.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(45).ColumnName, InnyMiesiac(Okres1, -8), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk109.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(46).ColumnName, InnyMiesiac(Okres1, -9), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk110.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(47).ColumnName, InnyMiesiac(Okres1, -10), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk111.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(48).ColumnName, InnyMiesiac(Okres1, -11), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk112.Checked, True, False))

            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(49).ColumnName, InnyMiesiac(Okres2, 0), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk201.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(50).ColumnName, InnyMiesiac(Okres2, -1), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk202.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(51).ColumnName, InnyMiesiac(Okres2, -2), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk203.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(52).ColumnName, InnyMiesiac(Okres2, -3), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk204.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(53).ColumnName, InnyMiesiac(Okres2, -4), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk205.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(54).ColumnName, InnyMiesiac(Okres2, -5), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk206.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(55).ColumnName, InnyMiesiac(Okres2, -6), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk207.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(56).ColumnName, InnyMiesiac(Okres2, -7), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk208.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(57).ColumnName, InnyMiesiac(Okres2, -8), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk209.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(58).ColumnName, InnyMiesiac(Okres2, -9), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk210.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(59).ColumnName, InnyMiesiac(Okres2, -10), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk211.Checked, True, False))
            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(60).ColumnName, InnyMiesiac(Okres2, -11), SzerKolumn, DataGridViewContentAlignment.MiddleCenter, "F" & IIf(AnalIlosciowo, 0, 2), IIf(Chk212.Checked, True, False))

            'NumColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(61).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "F" & IIf(AnalIlosciowo, 0, 2), False)  'mies 24
            ''----------------------------------------------------


            If Not Akt Then
                'e.CellStyle.BackColor = System.Drawing.Color.FromArgb(CType(180, Byte), CType(255, Byte), CType(180, Byte))
                'e.CellStyle.BackColor = System.Drawing.Color.FromArgb(CType(200, Byte), CType(255, Byte), CType(200, Byte))
                'e.CellStyle.BackColor = System.Drawing.Color.FromArgb(CType(220, Byte), CType(255, Byte), CType(220, Byte))
                e.CellStyle.BackColor = System.Drawing.Color.FromArgb(CType(251, Byte), CType(196, Byte), CType(200, Byte))
            End If

            Select Case e.ColumnIndex
                Case 44, 34
                    e.CellStyle.BackColor = System.Drawing.Color.GreenYellow
                Case 35, 36
                    e.CellStyle.BackColor = System.Drawing.Color.PaleTurquoise
                Case 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48
                    e.CellStyle.BackColor = System.Drawing.Color.LightCyan
                Case 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60
                    e.CellStyle.BackColor = System.Drawing.Color.Azure

                Case Else
            End Select
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
        '            PokazFiltryColumnDGV(Me.panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me.panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
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
                '    PokazFiltryColumnDGV(Me.panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
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
        oDataWeryfSegmentacjiKli = ""
        oPlanDlugookresowyKli = 0
        oPlanKrotkookresowyKli = 0

        aIlosAA = 0
        aIlosBB = 0

        aIlosBM = 0
        aIlos01 = 0
        aIlos02 = 0
        aIlos03 = 0
        aIlos04 = 0
        aIlos05 = 0
        aIlos06 = 0
        aIlos07 = 0
        aIlos08 = 0
        aIlos09 = 0
        aIlos10 = 0
        aIlos11 = 0

        aIlos12 = 0
        aIlos13 = 0
        aIlos14 = 0
        aIlos15 = 0
        aIlos16 = 0
        aIlos17 = 0
        aIlos18 = 0
        aIlos19 = 0
        aIlos20 = 0
        aIlos21 = 0
        aIlos22 = 0
        aIlos23 = 0

        aIlos24 = 0

        oRodzSprzetuLKli = False
        oRodzSprzetuAKli = False
        oRodzSprzetuTKli = False
        oOfertaTZlozeniaKli = ""
        oOfertaPlikKli = ""
        oPKontaktKli = ""
        oSkupUwagikli = ""
        oSprzedOpiekunkli = ""
        oSprzedOKontaktRKli = Mid(SysData, 1, 4)
        oSprzedOKontaktTKli = ""
        oSprzedOKontaktDKli = ""
        oSprzedNKontaktRKli = Mid(SysData, 1, 4)
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

    Private Sub PobierzKlienci()
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

        aIlosAA = GetIntField(dagKlienci, 35)
        aIlosBB = GetIntField(dagKlienci, 36)

        'aIlosBM = GetIntField(dagKlienci, 37)
        'aIlos01 = GetIntField(dagKlienci, 38)
        'aIlos02 = GetIntField(dagKlienci, 39)
        'aIlos03 = GetIntField(dagKlienci, 40)
        'aIlos04 = GetIntField(dagKlienci, 41)
        'aIlos05 = GetIntField(dagKlienci, 42)
        'aIlos06 = GetIntField(dagKlienci, 43)
        'aIlos07 = GetIntField(dagKlienci, 44)
        'aIlos08 = GetIntField(dagKlienci, 45)
        'aIlos09 = GetIntField(dagKlienci, 46)
        'aIlos10 = GetIntField(dagKlienci, 47)
        'aIlos11 = GetIntField(dagKlienci, 48)

        'aIlos12 = GetIntField(dagKlienci, 49)
        'aIlos13 = GetIntField(dagKlienci, 50)
        'aIlos14 = GetIntField(dagKlienci, 51)
        'aIlos15 = GetIntField(dagKlienci, 52)
        'aIlos16 = GetIntField(dagKlienci, 53)
        'aIlos17 = GetIntField(dagKlienci, 54)
        'aIlos18 = GetIntField(dagKlienci, 55)
        'aIlos19 = GetIntField(dagKlienci, 56)
        'aIlos20 = GetIntField(dagKlienci, 57)
        'aIlos21 = GetIntField(dagKlienci, 58)
        'aIlos22 = GetIntField(dagKlienci, 59)
        'aIlos23 = GetIntField(dagKlienci, 60)

        'aIlos24 = GetIntField(dagKlienci, 61)
        '----------------
        oRodzSprzetuLKli = GetLogField(dagKlienci, 62)
        oRodzSprzetuAKli = GetLogField(dagKlienci, 63)
        oRodzSprzetuTKli = GetLogField(dagKlienci, 64)

        oOfertaTZlozeniaKli = GetTxtField(dagKlienci, 65)
        oOfertaPlikKli = GetTxtField(dagKlienci, 66)
        oPKontaktKli = GetTxtField(dagKlienci, 67)

        oSkupUwagikli = GetTxtField(dagKlienci, 68)
        oSprzedOpiekunkli = GetTxtField(dagKlienci, 69)

        oSprzedOKontaktRKli = GetTxtField(dagKlienci, 70)
        oSprzedOKontaktTKli = GetTxtField(dagKlienci, 71)
        oSprzedOKontaktDKli = GetTxtField(dagKlienci, 72)
        oSprzedNKontaktRKli = GetTxtField(dagKlienci, 73)
        oSprzedNKontaktTKli = GetTxtField(dagKlienci, 74)
        oSprzedNKontaktDKli = GetTxtField(dagKlienci, 75)

        oSprzedUwagiKli = GetTxtField(dagKlienci, 76)
        oWpisalKli = GetTxtField(dagKlienci, 77)
        oUwagikli = GetTxtField(dagKlienci, 78)
        oMarkerKli = GetLogField(dagKlienci, 79)
        oMarkerPKli = GetLogField(dagKlienci, 80)
        oAktywnyKli = GetLogField(dagKlienci, 81)

        oZakupPopRokKli = GetDblField(dagKlienci, 82)
        oUslugiDodKli = GetTxtField(dagKlienci, 83)
        oRZURap01 = GetTxtField(dagKlienci, 84)
        oRZURap02 = GetTxtField(dagKlienci, 85)
        oRZURap03 = GetTxtField(dagKlienci, 86)
        oRZURap04 = GetTxtField(dagKlienci, 87)
        oRZURap05 = GetTxtField(dagKlienci, 88)
        oRZURap06 = GetTxtField(dagKlienci, 89)
        oRZURap07 = GetTxtField(dagKlienci, 90)
        oRZURap08 = GetTxtField(dagKlienci, 91)
        oRZURap09 = GetTxtField(dagKlienci, 92)

        oWersjaKli = GetTxtField(dagKlienci, 93)
    End Sub

    '**************************************************************
    '* obsluga komend
    '**************************************************************

    Private Sub cmdWybierzSchemat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWybierzSchemat.Click
        oCoMamRobic = "WYBIERAJ"
        oLokalnyFiltr = ""
        oNazwaSchematu = ""
        oSchematFiltrowania = ""

        Dim Form As New SchematFiltrowania("AnalizaSprzedazy")
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

            'FiltrujDataView(Me.panel4,
            '                dagKlienci,
            '                DataViewKlienci.Table.Columns.Count,
            '                DataViewKlienci,
            '                stbKlienci,
            '                Trim(oSchematFiltrowania))
            ''--------------------------------------------------------------------
            PokazSumy(DataViewKlienci)

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
        ZapiszParametryAnalizy()
        '------------------------------
        If DataSetKlienci.Tables(0).Rows.Count > 0 Then
            oKlient = GetTxtField(dagKlienci, 0)
        Else
            oKlient = ""        'IDENT       Text, 10
        End If
        Me.Close()
    End Sub



    Private Sub ZapiszParametryAnalizy()
        Par_MiesAnalizy1 = IIf(Chk101.Checked, "T", "N") &
                           IIf(Chk102.Checked, "T", "N") &
                           IIf(Chk103.Checked, "T", "N") &
                           IIf(Chk104.Checked, "T", "N") &
                           IIf(Chk105.Checked, "T", "N") &
                           IIf(Chk106.Checked, "T", "N") &
                           IIf(Chk107.Checked, "T", "N") &
                           IIf(Chk108.Checked, "T", "N") &
                           IIf(Chk109.Checked, "T", "N") &
                           IIf(Chk110.Checked, "T", "N") &
                           IIf(Chk111.Checked, "T", "N") &
                           IIf(Chk112.Checked, "T", "N")
        Par_MiesAnalizy2 = IIf(Chk201.Checked, "T", "N") &
                           IIf(Chk202.Checked, "T", "N") &
                           IIf(Chk203.Checked, "T", "N") &
                           IIf(Chk204.Checked, "T", "N") &
                           IIf(Chk205.Checked, "T", "N") &
                           IIf(Chk206.Checked, "T", "N") &
                           IIf(Chk207.Checked, "T", "N") &
                           IIf(Chk208.Checked, "T", "N") &
                           IIf(Chk209.Checked, "T", "N") &
                           IIf(Chk210.Checked, "T", "N") &
                           IIf(Chk211.Checked, "T", "N") &
                           IIf(Chk212.Checked, "T", "N")


        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne")
        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Iloœci")
        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne")
        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Wartoœci")
        'cbxCoAnalizowac.Items.Add("Analizuj wartoœæ Mar¿y")
        'cbxCoAnalizowac.Items.Add("Analizuj procent Mar¿y")


        Par_DaneDoAnalizy = Mid(Par_DaneDoAnalizy, 1, 6) &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Iloœci", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Wartoœci", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj wartoœæ Mar¿y", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj procent Mar¿y", "T", "N")

        ZapiszParametry(Me)
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
                oCoDalej = ""
                oAktualizuj = False
                Dim Form As New EdytujKatalogKlientow
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
                        '-------------------
                        '-------------------
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
            'PokazSumy()
        End If
    End Sub


    Private Sub cmdUsuñ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñ.Click
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
                    Exit Do
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
                    '-------------------
                    '-------------------
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
                    'aktualizuj DataGrid
                    dagKlienci.Update()
                    'wyswietl ilosc rekordow
                    stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString
                    AktualizujBaze()
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
            '        dbDataAdapterKlienci.Fill(DataSetKlienci)
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






    '*************************************************
    '** usuwanie zapisow z katalogu OSOBY
    '*************************************************

    Dim sSelectSQLOsoby As String = "SELECT * FROM Osoby WHERE IDENTKLIENTA='" & oIdentKli & "'"

    Dim dbConnectionOsoby As OleDb.OleDbConnection
    Dim dbCommandSelectOsoby As OleDb.OleDbCommand
    Dim dbCommandDeleteOsoby As OleDb.OleDbCommand
    Dim dbDataAdapterOsoby As OleDb.OleDbDataAdapter

    Dim sqlConnectionOsoby As SqlClient.SqlConnection
    Dim sqlCommandSelectOsoby As SqlClient.SqlCommand
    Dim sqlCommandDeleteOsoby As SqlClient.SqlCommand
    Dim sqlDataAdapterOsoby As SqlClient.SqlDataAdapter

    Dim DataSetOsoby As New DataSet
    Dim DataViewOsoby As New DataView

    Private Sub UsunOsoby(ByVal IdKl As String)
        sSelectSQLOsoby = "SELECT * FROM Osoby WHERE IDENTKLIENTA='" & IdKl & "'"
        DataSetOsoby.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
        Else


            sqlConnectionOsoby = New SqlClient.SqlConnection(sConnectionOsoby)
            sqlCommandSelectOsoby = New SqlClient.SqlCommand(sSelectSQLOsoby, sqlConnectionOsoby)
            sqlCommandDeleteOsoby = New SqlClient.SqlCommand(sDeleteSQLOsoby, sqlConnectionOsoby)
            sqlDataAdapterOsoby = New SqlClient.SqlDataAdapter(sqlCommandSelectOsoby)
            '---------------------------------------------
            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliOsoby As System.Data.Common.DataTableMapping
            MapowanieTabeliOsoby = sqlDataAdapterOsoby.TableMappings.Add("Table", "TABELA_Osoby")

            'komendy DataAdapter
            Dim objParamOsoby As SqlClient.SqlParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            objParamOsoby = sqlCommandDeleteOsoby.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
            objParamOsoby.SourceVersion = DataRowVersion.Original
            objParamOsoby = sqlCommandDeleteOsoby.Parameters.Add("@orygOsoba", SqlDbType.Char, 50, "OSOBA")
            objParamOsoby.SourceVersion = DataRowVersion.Original
            objParamOsoby = sqlCommandDeleteOsoby.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParamOsoby.SourceVersion = DataRowVersion.Original
            sqlDataAdapterOsoby.DeleteCommand = sqlCommandDeleteOsoby

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterOsoby.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionOsoby.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            If sqlConnectionOsoby.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterOsoby.FillSchema(DataSetOsoby, SchemaType.Mapped, "TABELA_Osoby")
                    sqlDataAdapterOsoby.Fill(DataSetOsoby)
                    sqlConnectionOsoby.Close()

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

        Dim dr As DataRow
        Dim Id As String
        For Each dr In DataSetOsoby.Tables(0).Rows
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
                sqlConnectionOsoby.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionOsoby.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterOsoby.Update(DataSetOsoby, "TABELA_Osoby")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterOsoby.Fill(DataSetOsoby)
                End If
                sqlConnectionOsoby.Close()
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

    Dim SQLConnectionKontakty As SqlClient.SqlConnection
    Dim SQLCommandSelectKontakty As SqlClient.SqlCommand
    Dim SQLCommandDeleteKontakty As SqlClient.SqlCommand
    Dim SQLDataAdapterKontakty As SqlClient.SqlDataAdapter

    Dim DataSetKontakty As New DataSet
    Dim DataViewKontakty As New DataView

    Private Sub UsunKontakty(ByVal IdKl As String)
        sSelectSQLKontakty = "SELECT * FROM KontaktyHandlowe WHERE IDENTKLIENTA='" & IdKl & "'"
        DataSetKontakty.Locale = New System.Globalization.CultureInfo("pl-PL")


        If ParL_DataBaseType = "MS ACCESS" Then
        Else
            SQLConnectionKontakty = New SqlClient.SqlConnection(sConnectionKontakty)
            SQLCommandSelectKontakty = New SqlClient.SqlCommand(sSelectSQLKontakty, SQLConnectionKontakty)
            SQLCommandDeleteKontakty = New SqlClient.SqlCommand(sDeleteSQLKontakty, SQLConnectionKontakty)
            SQLDataAdapterKontakty = New SqlClient.SqlDataAdapter(SQLCommandSelectKontakty)
            '---------------------------------------------
            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliKontakty As System.Data.Common.DataTableMapping
            MapowanieTabeliKontakty = SQLDataAdapterKontakty.TableMappings.Add("Table", "TABELA_Kontakty")

            'komendy DataAdapter
            Dim objParamKontakty As SqlClient.SqlParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            objParamKontakty = SQLCommandDeleteKontakty.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
            objParamKontakty.SourceVersion = DataRowVersion.Original
            objParamKontakty = SQLCommandDeleteKontakty.Parameters.Add("@orygData", SqlDbType.Char, 10, "DATAKONTAKTU")
            objParamKontakty.SourceVersion = DataRowVersion.Original
            objParamKontakty = SQLCommandDeleteKontakty.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParamKontakty.SourceVersion = DataRowVersion.Original
            SQLDataAdapterKontakty.DeleteCommand = SQLCommandDeleteKontakty

            ' Add the RowUpdated event handler.
            AddHandler SQLDataAdapterKontakty.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                SQLConnectionKontakty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            If SQLConnectionKontakty.State = ConnectionState.Open Then
                Try
                    SQLDataAdapterKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
                    SQLDataAdapterKontakty.Fill(DataSetKontakty)
                    SQLConnectionKontakty.Close()

                    'zdefiniuj unikalny klucz wyszukiwania w tabeli
                    DataSetKontakty.Tables(0).PrimaryKey = New DataColumn() {DataSetKontakty.Tables(0).Columns("IDENTKLIENTA"), DataSetKontakty.Tables(0).Columns("DATAKONTAKTU")}

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

        Dim dr As DataRow
        Dim Id As String
        For Each dr In DataSetKontakty.Tables(0).Rows
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
                SQLConnectionKontakty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If SQLConnectionKontakty.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    SQLDataAdapterKontakty.Update(DataSetKontakty, "TABELA_Kontakty")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    SQLDataAdapterKontakty.Fill(DataSetKontakty)
                End If
                SQLConnectionKontakty.Close()
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

    Dim sqlConnectionObroty As SqlClient.SqlConnection
    Dim sqlCommandSelectObroty As SqlClient.SqlCommand
    Dim sqlCommandDeleteObroty As SqlClient.SqlCommand
    Dim sqlDataAdapterObroty As SqlClient.SqlDataAdapter

    Dim DataSetObroty As New DataSet
    Dim DataViewObroty As New DataView

    Private Sub UsunObroty(ByVal IdKl As String)
        sSelectSQLObroty = "SELECT * FROM Obroty WHERE IDENTKLIENTA='" & IdKl & "'"
        DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
        Else
            sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionObroty)
            sqlCommandSelectObroty = New SqlClient.SqlCommand(sSelectSQLObroty, sqlConnectionObroty)
            sqlCommandDeleteObroty = New SqlClient.SqlCommand(sDeleteSQLObroty, sqlConnectionObroty)
            sqlDataAdapterObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)
            '---------------------------------------------
            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliObroty As System.Data.Common.DataTableMapping
            MapowanieTabeliObroty = sqlDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")

            'komendy DataAdapter
            Dim objParamObroty As SqlClient.SqlParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            objParamObroty = sqlCommandDeleteObroty.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
            objParamObroty.SourceVersion = DataRowVersion.Original
            objParamObroty = sqlCommandDeleteObroty.Parameters.Add("@orygData", SqlDbType.Char, 10, "DATA")
            objParamObroty.SourceVersion = DataRowVersion.Original
            objParamObroty = sqlCommandDeleteObroty.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParamObroty.SourceVersion = DataRowVersion.Original
            sqlDataAdapterObroty.DeleteCommand = sqlCommandDeleteObroty

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterObroty.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObroty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            If sqlConnectionObroty.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    sqlDataAdapterObroty.Fill(DataSetObroty)
                    sqlConnectionObroty.Close()

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

        Dim dr As DataRow
        Dim Id As String
        For Each dr In DataSetObroty.Tables(0).Rows
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
                sqlConnectionObroty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionObroty.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterObroty.Update(DataSetObroty, "TABELA_Obroty")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterObroty.Fill(DataSetObroty)
                End If
                sqlConnectionObroty.Close()
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

    Dim sqlConnectionObrotyMies As SqlClient.SqlConnection
    Dim sqlCommandSelectObrotyMies As SqlClient.SqlCommand
    Dim sqlCommandDeleteObrotyMies As SqlClient.SqlCommand
    Dim sqlDataAdapterObrotyMies As SqlClient.SqlDataAdapter

    Dim DataSetObrotyMies As New DataSet
    Dim DataViewObrotyMies As New DataView

    Private Sub UsunObrotyMies(ByVal IdKl As String)
        sSelectSQLObrotyMies = "SELECT * FROM ObrotyMies WHERE IDENTKLIENTA='" & IdKl & "'"
        DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
        Else
            sqlConnectionObrotyMies = New SqlClient.SqlConnection(sConnectionObrotyMies)
            sqlCommandSelectObrotyMies = New SqlClient.SqlCommand(sSelectSQLObrotyMies, sqlConnectionObrotyMies)
            sqlCommandDeleteObrotyMies = New SqlClient.SqlCommand(sDeleteSQLObrotyMies, sqlConnectionObrotyMies)
            sqlDataAdapterObrotyMies = New SqlClient.SqlDataAdapter(sqlCommandSelectObrotyMies)
            '---------------------------------------------
            'mapujemy domyslna nazwe tabeli
            Dim MapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            MapowanieTabeliObrotyMies = sqlDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")

            'komendy DataAdapter
            Dim objParamObrotyMies As SqlClient.SqlParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            objParamObrotyMies = sqlCommandDeleteObrotyMies.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
            objParamObrotyMies.SourceVersion = DataRowVersion.Original
            objParamObrotyMies = sqlCommandDeleteObrotyMies.Parameters.Add("@orygMies", SqlDbType.Char, 7, "MIESIAC")
            objParamObrotyMies.SourceVersion = DataRowVersion.Original
            objParamObrotyMies = sqlCommandDeleteObrotyMies.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            objParamObrotyMies.SourceVersion = DataRowVersion.Original
            sqlDataAdapterObrotyMies.DeleteCommand = sqlCommandDeleteObrotyMies

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterObrotyMies.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObrotyMies.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            If sqlConnectionObrotyMies.State = ConnectionState.Open Then
                Try
                    sqlDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    sqlDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    sqlConnectionObrotyMies.Close()

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

        Dim dr As DataRow
        Dim Id As String
        For Each dr In DataSetObrotyMies.Tables(0).Rows
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
                sqlConnectionObrotyMies.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionObrotyMies.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterObrotyMies.Update(DataSetObrotyMies, "TABELA_ObrotyMies")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                End If
                sqlConnectionObrotyMies.Close()
            End If
        End If
    End Sub






    Private Sub cmdOdswiez_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOdswiez.Click
        OdswiezBaze()
        PokazSumy(DataViewKlienci)
    End Sub













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
                'PokazSumy(DataViewKlienci)
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
                'PokazSumy(DataViewKlienci)
            End If
        End If
    End Sub



    Private Sub cbxCoAnalizowac_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxCoAnalizowac.SelectedIndexChanged
        If Not StartFormy Then
            StartFormy = True
            PrzygotujAnalize()
            WyliczSumyDlaTabeli()
            StartFormy = False
            dagKlienci.Refresh()
            '-------------------------------
            InicjujFiltryColumnDGV(Me.Panel4, dagKlienci, DataViewKlienci.Table.Columns.Count, stbFiltry, StartFormy)
            dagKlienci_CurrentCellChanged(sender, e)
        End If
    End Sub







    'Private Sub cmdDostosuj_Click(sender As Object, e As EventArgs) Handles cmdDostosuj.Click
    '    WyliczSumyDlaTabeli()
    '    '------------------
    '    MessageBox.Show("Zweryfikowa³em opisy kolumn w tabeli." & vbCrLf & _
    '        "Wyliczy³em sumy zbiorcze dla tabeli.", _
    '        "Ok, skoñczy³em.", _
    '        System.Windows.Forms.MessageBoxButtons.OK, _
    '        MessageBoxIcon.Information, _
    '        MessageBoxDefaultButton.Button1)
    'End Sub


    Private Sub WyliczSumyDlaTabeli()
        Me.Cursor = Cursors.WaitCursor
        dagKlienci.Columns.Item(37).Width = IIf(Chk101.Checked, 50, 0)
        dagKlienci.Columns.Item(38).Width = IIf(Chk102.Checked, 50, 0)
        dagKlienci.Columns.Item(39).Width = IIf(Chk103.Checked, 50, 0)
        dagKlienci.Columns.Item(40).Width = IIf(Chk104.Checked, 50, 0)
        dagKlienci.Columns.Item(41).Width = IIf(Chk105.Checked, 50, 0)
        dagKlienci.Columns.Item(42).Width = IIf(Chk106.Checked, 50, 0)
        dagKlienci.Columns.Item(43).Width = IIf(Chk107.Checked, 50, 0)
        dagKlienci.Columns.Item(44).Width = IIf(Chk108.Checked, 50, 0)
        dagKlienci.Columns.Item(45).Width = IIf(Chk109.Checked, 50, 0)
        dagKlienci.Columns.Item(46).Width = IIf(Chk110.Checked, 50, 0)
        dagKlienci.Columns.Item(47).Width = IIf(Chk111.Checked, 50, 0)
        dagKlienci.Columns.Item(48).Width = IIf(Chk112.Checked, 50, 0)

        dagKlienci.Columns.Item(49).Width = IIf(Chk201.Checked, 50, 0)
        dagKlienci.Columns.Item(50).Width = IIf(Chk202.Checked, 50, 0)
        dagKlienci.Columns.Item(51).Width = IIf(Chk203.Checked, 50, 0)
        dagKlienci.Columns.Item(52).Width = IIf(Chk204.Checked, 50, 0)
        dagKlienci.Columns.Item(53).Width = IIf(Chk205.Checked, 50, 0)
        dagKlienci.Columns.Item(54).Width = IIf(Chk206.Checked, 50, 0)
        dagKlienci.Columns.Item(55).Width = IIf(Chk207.Checked, 50, 0)
        dagKlienci.Columns.Item(56).Width = IIf(Chk208.Checked, 50, 0)
        dagKlienci.Columns.Item(57).Width = IIf(Chk209.Checked, 50, 0)
        dagKlienci.Columns.Item(58).Width = IIf(Chk210.Checked, 50, 0)
        dagKlienci.Columns.Item(59).Width = IIf(Chk211.Checked, 50, 0)
        dagKlienci.Columns.Item(60).Width = IIf(Chk212.Checked, 50, 0)
        '------------------------------
        dagKlienci.Columns.Item(37).HeaderText = IIf(Chk101.Checked, InnyMiesiac(Okres1, 0), "")
        dagKlienci.Columns.Item(38).HeaderText = IIf(Chk102.Checked, InnyMiesiac(Okres1, -1), "")
        dagKlienci.Columns.Item(39).HeaderText = IIf(Chk103.Checked, InnyMiesiac(Okres1, -2), "")
        dagKlienci.Columns.Item(40).HeaderText = IIf(Chk104.Checked, InnyMiesiac(Okres1, -3), "")
        dagKlienci.Columns.Item(41).HeaderText = IIf(Chk105.Checked, InnyMiesiac(Okres1, -4), "")
        dagKlienci.Columns.Item(42).HeaderText = IIf(Chk106.Checked, InnyMiesiac(Okres1, -5), "")
        dagKlienci.Columns.Item(43).HeaderText = IIf(Chk107.Checked, InnyMiesiac(Okres1, -6), "")
        dagKlienci.Columns.Item(44).HeaderText = IIf(Chk108.Checked, InnyMiesiac(Okres1, -7), "")
        dagKlienci.Columns.Item(45).HeaderText = IIf(Chk109.Checked, InnyMiesiac(Okres1, -8), "")
        dagKlienci.Columns.Item(46).HeaderText = IIf(Chk110.Checked, InnyMiesiac(Okres1, -9), "")
        dagKlienci.Columns.Item(47).HeaderText = IIf(Chk111.Checked, InnyMiesiac(Okres1, -10), "")
        dagKlienci.Columns.Item(48).HeaderText = IIf(Chk112.Checked, InnyMiesiac(Okres1, -11), "")

        dagKlienci.Columns.Item(49).HeaderText = IIf(Chk201.Checked, InnyMiesiac(Okres2, 0), "")
        dagKlienci.Columns.Item(50).HeaderText = IIf(Chk202.Checked, InnyMiesiac(Okres2, -1), "")
        dagKlienci.Columns.Item(51).HeaderText = IIf(Chk203.Checked, InnyMiesiac(Okres2, -2), "")
        dagKlienci.Columns.Item(52).HeaderText = IIf(Chk204.Checked, InnyMiesiac(Okres2, -3), "")
        dagKlienci.Columns.Item(53).HeaderText = IIf(Chk205.Checked, InnyMiesiac(Okres2, -4), "")
        dagKlienci.Columns.Item(54).HeaderText = IIf(Chk206.Checked, InnyMiesiac(Okres2, -5), "")
        dagKlienci.Columns.Item(55).HeaderText = IIf(Chk207.Checked, InnyMiesiac(Okres2, -6), "")
        dagKlienci.Columns.Item(56).HeaderText = IIf(Chk208.Checked, InnyMiesiac(Okres2, -7), "")
        dagKlienci.Columns.Item(57).HeaderText = IIf(Chk209.Checked, InnyMiesiac(Okres2, -8), "")
        dagKlienci.Columns.Item(58).HeaderText = IIf(Chk210.Checked, InnyMiesiac(Okres2, -9), "")
        dagKlienci.Columns.Item(59).HeaderText = IIf(Chk211.Checked, InnyMiesiac(Okres2, -10), "")
        dagKlienci.Columns.Item(60).HeaderText = IIf(Chk212.Checked, InnyMiesiac(Okres2, -11), "")
        '-----------------------
        dagKlienci.Columns.Item(37).Visible = IIf(Chk101.Checked, True, False)
        dagKlienci.Columns.Item(38).Visible = IIf(Chk102.Checked, True, False)
        dagKlienci.Columns.Item(39).Visible = IIf(Chk103.Checked, True, False)
        dagKlienci.Columns.Item(40).Visible = IIf(Chk104.Checked, True, False)
        dagKlienci.Columns.Item(41).Visible = IIf(Chk105.Checked, True, False)
        dagKlienci.Columns.Item(42).Visible = IIf(Chk106.Checked, True, False)
        dagKlienci.Columns.Item(43).Visible = IIf(Chk107.Checked, True, False)
        dagKlienci.Columns.Item(44).Visible = IIf(Chk108.Checked, True, False)
        dagKlienci.Columns.Item(45).Visible = IIf(Chk109.Checked, True, False)
        dagKlienci.Columns.Item(46).Visible = IIf(Chk110.Checked, True, False)
        dagKlienci.Columns.Item(47).Visible = IIf(Chk111.Checked, True, False)
        dagKlienci.Columns.Item(48).Visible = IIf(Chk112.Checked, True, False)

        dagKlienci.Columns.Item(49).Visible = IIf(Chk201.Checked, True, False)
        dagKlienci.Columns.Item(50).Visible = IIf(Chk202.Checked, True, False)
        dagKlienci.Columns.Item(51).Visible = IIf(Chk203.Checked, True, False)
        dagKlienci.Columns.Item(52).Visible = IIf(Chk204.Checked, True, False)
        dagKlienci.Columns.Item(53).Visible = IIf(Chk205.Checked, True, False)
        dagKlienci.Columns.Item(54).Visible = IIf(Chk206.Checked, True, False)
        dagKlienci.Columns.Item(55).Visible = IIf(Chk207.Checked, True, False)
        dagKlienci.Columns.Item(56).Visible = IIf(Chk208.Checked, True, False)
        dagKlienci.Columns.Item(57).Visible = IIf(Chk209.Checked, True, False)
        dagKlienci.Columns.Item(58).Visible = IIf(Chk210.Checked, True, False)
        dagKlienci.Columns.Item(59).Visible = IIf(Chk211.Checked, True, False)
        dagKlienci.Columns.Item(60).Visible = IIf(Chk212.Checked, True, False)

        PokazSumy()
        Me.Cursor = Cursors.Default
    End Sub


    Private Sub cmdPrzelicz_Click(sender As Object, e As EventArgs) Handles cmdPrzeliczSprzedaz.Click
        Par_DataAktAnalizy = SysData()

        Par_DaneDoAnalizy = IIf(chkPryzmatAtrament.Checked And chkPryzmatTonery.Checked, "T", IIf(chkPryzmatAtrament.Checked, "A", IIf(chkPryzmatTonery.Checked, "L", "N"))) &
                            IIf(chkOrygAtrament.Checked And chkOrygTonery.Checked, "T", IIf(chkOrygAtrament.Checked, "A", IIf(chkOrygTonery.Checked, "L", "N"))) &
                            IIf(chkNajem.Checked, "T", "N") &
                            IIf(chkStrony.Checked, "T", "N") &
                            IIf(chkDrukarki.Checked, "T", "N") &
                            IIf(chkSkup.Checked, "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Iloœci", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Wartoœci", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj wartoœæ Mar¿y", "T", "N") &
                   IIf(cbxCoAnalizowac.Text = "Analizuj procent Mar¿y", "T", "N")


        Par_MiesAnalizy1 = IIf(Chk101.Checked, "T", "N") &
                           IIf(Chk102.Checked, "T", "N") &
                           IIf(Chk103.Checked, "T", "N") &
                           IIf(Chk104.Checked, "T", "N") &
                           IIf(Chk105.Checked, "T", "N") &
                           IIf(Chk106.Checked, "T", "N") &
                           IIf(Chk107.Checked, "T", "N") &
                           IIf(Chk108.Checked, "T", "N") &
                           IIf(Chk109.Checked, "T", "N") &
                           IIf(Chk110.Checked, "T", "N") &
                           IIf(Chk111.Checked, "T", "N") &
                           IIf(Chk112.Checked, "T", "N")

        Par_MiesAnalizy2 = IIf(Chk201.Checked, "T", "N") &
                           IIf(Chk202.Checked, "T", "N") &
                           IIf(Chk203.Checked, "T", "N") &
                           IIf(Chk204.Checked, "T", "N") &
                           IIf(Chk205.Checked, "T", "N") &
                           IIf(Chk206.Checked, "T", "N") &
                           IIf(Chk207.Checked, "T", "N") &
                           IIf(Chk208.Checked, "T", "N") &
                           IIf(Chk209.Checked, "T", "N") &
                           IIf(Chk210.Checked, "T", "N") &
                           IIf(Chk211.Checked, "T", "N") &
                           IIf(Chk212.Checked, "T", "N")





        ZapiszParametry(Me)
        '---------------------------------------------------
        Dim Form As New PobierzObrotyzTOFIdoAnalizy2()
        Form.ShowDialog()
        '--------------------------
        'Par_OSTOkresAnalizy="RRRR-MM" & "RRRR-MM"
        'I okres analizy
        If Len(Trim(Mid(Par_OstOkresAnalizy, 1, 7))) = 0 Then
            Okres1 = Mid(SysData, 1, 7)
        Else
            Okres1 = Mid(Par_OstOkresAnalizy, 1, 7)
        End If

        'II okres analizy
        If Len(Trim(Mid(Par_OstOkresAnalizy, 8, 7))) = 0 Then
            Okres2 = MiesDoAnalizy(Okres1, -12)
        Else
            Okres2 = Mid(Par_OstOkresAnalizy, 8, 7)
        End If
        lblOkres1.Text = "SPRZEDA¯ : " & InnyMiesiac(Okres1, 0) & " : 0..-11"
        lblOkres2.Text = InnyMiesiac(Okres2, 0) & " : 0..-11"

        Chk101.Checked = (Mid(Par_MiesAnalizy1, 1, 1) = "T")
        Chk102.Checked = (Mid(Par_MiesAnalizy1, 2, 1) = "T")
        Chk103.Checked = (Mid(Par_MiesAnalizy1, 3, 1) = "T")
        Chk104.Checked = (Mid(Par_MiesAnalizy1, 4, 1) = "T")
        Chk105.Checked = (Mid(Par_MiesAnalizy1, 5, 1) = "T")
        Chk106.Checked = (Mid(Par_MiesAnalizy1, 6, 1) = "T")
        Chk107.Checked = (Mid(Par_MiesAnalizy1, 7, 1) = "T")
        Chk108.Checked = (Mid(Par_MiesAnalizy1, 8, 1) = "T")
        Chk109.Checked = (Mid(Par_MiesAnalizy1, 9, 1) = "T")
        Chk110.Checked = (Mid(Par_MiesAnalizy1, 10, 1) = "T")
        Chk111.Checked = (Mid(Par_MiesAnalizy1, 11, 1) = "T")
        Chk112.Checked = (Mid(Par_MiesAnalizy1, 12, 1) = "T")

        Chk201.Checked = (Mid(Par_MiesAnalizy2, 1, 1) = "T")
        Chk202.Checked = (Mid(Par_MiesAnalizy2, 2, 1) = "T")
        Chk203.Checked = (Mid(Par_MiesAnalizy2, 3, 1) = "T")
        Chk204.Checked = (Mid(Par_MiesAnalizy2, 4, 1) = "T")
        Chk205.Checked = (Mid(Par_MiesAnalizy2, 5, 1) = "T")
        Chk206.Checked = (Mid(Par_MiesAnalizy2, 6, 1) = "T")
        Chk207.Checked = (Mid(Par_MiesAnalizy2, 7, 1) = "T")
        Chk208.Checked = (Mid(Par_MiesAnalizy2, 8, 1) = "T")
        Chk209.Checked = (Mid(Par_MiesAnalizy2, 9, 1) = "T")
        Chk210.Checked = (Mid(Par_MiesAnalizy2, 10, 1) = "T")
        Chk211.Checked = (Mid(Par_MiesAnalizy2, 11, 1) = "T")
        Chk212.Checked = (Mid(Par_MiesAnalizy2, 12, 1) = "T")


        chkPryzmatAtrament.Checked = Mid(Par_DaneDoAnalizy, 1, 1) = "T" Or Mid(Par_DaneDoAnalizy, 1, 1) = "A"
        chkPryzmatTonery.Checked = Mid(Par_DaneDoAnalizy, 1, 1) = "T" Or Mid(Par_DaneDoAnalizy, 1, 1) = "L"
        chkOrygAtrament.Checked = Mid(Par_DaneDoAnalizy, 2, 1) = "T" Or Mid(Par_DaneDoAnalizy, 2, 1) = "A"
        chkOrygTonery.Checked = Mid(Par_DaneDoAnalizy, 2, 1) = "T" Or Mid(Par_DaneDoAnalizy, 2, 1) = "L"
        chkNajem.Checked = Mid(Par_DaneDoAnalizy, 3, 1) = "T"
        chkStrony.Checked = Mid(Par_DaneDoAnalizy, 3, 1) = "T"
        chkDrukarki.Checked = Mid(Par_DaneDoAnalizy, 3, 1) = "T"
        chkSkup.Checked = Mid(Par_DaneDoAnalizy, 3, 1) = "T"

        If Mid(Par_DaneDoAnalizy, 7, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne"
        If Mid(Par_DaneDoAnalizy, 8, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne"
        If Mid(Par_DaneDoAnalizy, 9, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Iloœci"
        If Mid(Par_DaneDoAnalizy, 10, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Wartoœci"
        If Mid(Par_DaneDoAnalizy, 11, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj wartoœæ Mar¿y"
        If Mid(Par_DaneDoAnalizy, 12, 1) = "T" Then cbxCoAnalizowac.Text = "Analizuj procent Mar¿y"

        'Par_DaneDoAnalizy = IIf(chkPryzmat.Checked, "T", IIf(chkPryzmatAtrament.Checked, "A", IIf(chkPryzmatTonery.Checked, "L", "N"))) &
        '                    IIf(chkOryg.Checked, "T", IIf(chkOrygAtrament.Checked, "A", IIf(chkOrygTonery.Checked, "L", "N"))) &
        '                    IIf(chkNajem.Checked, "T", "N") &
        '                    IIf(chkStrony.Checked, "T", "N") &
        '                    IIf(chkDrukarki.Checked, "T", "N") &
        '                    IIf(chkSkup.Checked, "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj Iloœciowo-Sumuj Iloœci", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj Wartoœciowo-Sumuj Wartoœci", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj wartoœæ Mar¿y", "T", "N") &
        '           IIf(cbxCoAnalizowac.Text = "Analizuj procent Mar¿y", "T", "N")



        '-------------------------------
        OdswiezBaze()
        'aktualizuj DataGrid
        dagKlienci.Update()
        'wyswietl ilosc rekordow
        stbKlienci.Panels(0).Text = "Il.rec.: " & DataViewKlienci.Count.ToString
        WyliczSumyDlaTabeli()
        'PokazSumy(DataViewKlienci)
    End Sub













    '=================================
    ' wyswietl informacje o sumach...
    '===================================

    Private Sub PokazSumy()
        Dim dbSelectRoboczy As String = ""

        Dim dbConnectionRoboczy As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectRoboczy As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterRoboczy As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionRoboczy As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectRoboczy As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterRoboczy As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetRoboczy As New DataSet
        Dim DataViewRoboczy As New DataView




        Dim Val As Double = 0
        Dim Suma1 As Double = 0
        Dim Suma2 As Double = 0


        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne")
        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Iloœci")
        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne")
        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Wartoœci")
        'cbxCoAnalizowac.Items.Add("Analizuj wartoœæ Mar¿y")
        'cbxCoAnalizowac.Items.Add("Analizuj procent Mar¿y")


        Me.Cursor = Cursors.WaitCursor



        lblM00.Text = InnyMiesiac(Okres1, 0)
        lblM01.Text = InnyMiesiac(Okres1, -1)
        lblM02.Text = InnyMiesiac(Okres1, -2)
        lblM03.Text = InnyMiesiac(Okres1, -3)
        lblM04.Text = InnyMiesiac(Okres1, -4)
        lblM05.Text = InnyMiesiac(Okres1, -5)
        lblM06.Text = InnyMiesiac(Okres1, -6)
        lblM07.Text = InnyMiesiac(Okres1, -7)
        lblM08.Text = InnyMiesiac(Okres1, -8)
        lblM09.Text = InnyMiesiac(Okres1, -9)
        lblM10.Text = InnyMiesiac(Okres1, -10)
        lblM11.Text = InnyMiesiac(Okres1, -11)

        lblM20.Text = InnyMiesiac(Okres2, 0)
        lblM21.Text = InnyMiesiac(Okres2, -1)
        lblM22.Text = InnyMiesiac(Okres2, -2)
        lblM23.Text = InnyMiesiac(Okres2, -3)
        lblM24.Text = InnyMiesiac(Okres2, -4)
        lblM25.Text = InnyMiesiac(Okres2, -5)
        lblM26.Text = InnyMiesiac(Okres2, -6)
        lblM27.Text = InnyMiesiac(Okres2, -7)
        lblM28.Text = InnyMiesiac(Okres2, -8)
        lblM29.Text = InnyMiesiac(Okres2, -9)
        lblM30.Text = InnyMiesiac(Okres2, -10)
        lblM31.Text = InnyMiesiac(Okres2, -11)

        valM00.Text = "0.00"
        valM01.Text = "0.00"
        valM02.Text = "0.00"
        valM03.Text = "0.00"
        valM04.Text = "0.00"
        valM05.Text = "0.00"
        valM06.Text = "0.00"
        valM07.Text = "0.00"
        valM08.Text = "0.00"
        valM09.Text = "0.00"
        valM10.Text = "0.00"
        valM11.Text = "0.00"

        valM20.Text = "0.00"
        valM21.Text = "0.00"
        valM22.Text = "0.00"
        valM23.Text = "0.00"
        valM24.Text = "0.00"
        valM25.Text = "0.00"
        valM26.Text = "0.00"
        valM27.Text = "0.00"
        valM28.Text = "0.00"
        valM29.Text = "0.00"
        valM30.Text = "0.00"
        valM31.Text = "0.00"
        '------------



        Dim AnalIlosciowo As Boolean = False
        Dim AnalWartMarzy As Boolean = False
        Dim AnalProcMarzy As Boolean = False
        Dim SumujAktywnosci As Boolean = False

        Select Case cbxCoAnalizowac.Text
            Case "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne"
                AnalIlosciowo = True
                SumujAktywnosci = True
                AnalWartMarzy = False
                AnalProcMarzy = False
            Case "Analizuj Iloœciowo-Sumuj Iloœci"
                AnalIlosciowo = True
                SumujAktywnosci = False
                AnalWartMarzy = False
                AnalProcMarzy = False
            Case "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne"
                AnalIlosciowo = False
                SumujAktywnosci = True
                AnalWartMarzy = False
                AnalProcMarzy = False
            Case "Analizuj Wartoœciowo-Sumuj Wartoœci"
                AnalIlosciowo = False
                SumujAktywnosci = False
                AnalWartMarzy = False
                AnalProcMarzy = False
            Case "Analizuj wartoœæ Mar¿y"
                AnalIlosciowo = False
                SumujAktywnosci = False
                AnalWartMarzy = True
                AnalProcMarzy = False
            Case "Analizuj procent Mar¿y"
                AnalIlosciowo = False
                SumujAktywnosci = False
                AnalWartMarzy = False
                AnalProcMarzy = True
        End Select


        dbSelectRoboczy = "SELECT TOP 1 "
        If SumujAktywnosci Then
            Label18.Text = "Sum.SPRZED.Aktywnoœci"
            dbSelectRoboczy &=
                                         "(SELECT SUM(WARTZA00) FROM ANALIZYZAKUPU WHERE ((SELECT AKTYWNY FROM KLIENCI WHERE KLIENCI.IDENTKLIENTA=ANALIZYZAKUPU.IDENTKLIENTA)='TRUE')) AS ILOSCAA," &
                                         "(SELECT SUM(ILOSZA00) FROM ANALIZYZAKUPU WHERE ((SELECT AKTYWNY FROM KLIENCI WHERE KLIENCI.IDENTKLIENTA=ANALIZYZAKUPU.IDENTKLIENTA)='TRUE')) AS ILOSCBB,"
        Else
            'nie aktywnoœci - sumy
            If AnalIlosciowo Then
                Label18.Text = "Sum.SPRZED.Iloœci"
                dbSelectRoboczy &=
                    "(SELECT SUM(ILOSZABM + ILOSZA01 + ILOSZA02 + ILOSZA03 + ILOSZA04 + ILOSZA05 + ILOSZA06 + ILOSZA07 + ILOSZA08 + ILOSZA09 + ILOSZA10 + ILOSZA11) FROM ANALIZYZAKUPU WHERE ((SELECT AKTYWNY FROM KLIENCI WHERE KLIENCI.IDENTKLIENTA=ANALIZYZAKUPU.IDENTKLIENTA)='TRUE')) AS ILOSCAA," &
                    "(SELECT SUM(ILOSZA12 + ILOSZA13 + ILOSZA14 + ILOSZA15 + ILOSZA16 + ILOSZA17 + ILOSZA18 + ILOSZA19 + ILOSZA20 + ILOSZA21 + ILOSZA22 + ILOSZA23) FROM ANALIZYZAKUPU WHERE ((SELECT AKTYWNY FROM KLIENCI WHERE KLIENCI.IDENTKLIENTA=ANALIZYZAKUPU.IDENTKLIENTA)='TRUE')) AS ILOSCBB,"
            ElseIf AnalWartMarzy Then
                Label18.Text = "Sum.SPRZED.Mar¿a"
                dbSelectRoboczy &=
                    "(SELECT SUM(MARZABM + MARZA01 + MARZA02 + MARZA03 + MARZA04 + MARZA05 + MARZA06 + MARZA07 + MARZA08 + MARZA09 + MARZA10 + MARZA11) FROM ANALIZYZAKUPU WHERE ((SELECT AKTYWNY FROM KLIENCI WHERE KLIENCI.IDENTKLIENTA=ANALIZYZAKUPU.IDENTKLIENTA)='TRUE')) AS ILOSCAA," &
                    "(SELECT SUM(MARZA12 + MARZA13 + MARZA14 + MARZA15 + MARZA16 + MARZA17 + MARZA18 + MARZA19 + MARZA20 + MARZA21 + MARZA22 + MARZA23) FROM ANALIZYZAKUPU WHERE ((SELECT AKTYWNY FROM KLIENCI WHERE KLIENCI.IDENTKLIENTA=ANALIZYZAKUPU.IDENTKLIENTA)='TRUE')) AS ILOSCBB,"
            ElseIf AnalProcMarzy Then
                Label18.Text = "Sum.SPRZED.Proc.Mar¿y"
                dbSelectRoboczy &=
                    "(SELECT (SUM((PROCMBM + PROCM01 + PROCM02 + PROCM03 + PROCM04 + PROCM05 + PROCM06 + PROCM07 + PROCM08 + PROCM09 + PROCM10 + PROCM11) / 12) / (SELECT COUNT(*) FROM KLIENCI WHERE (AKTYWNY='TRUE'))) FROM ANALIZYZAKUPU WHERE ((SELECT AKTYWNY FROM KLIENCI WHERE KLIENCI.IDENTKLIENTA=ANALIZYZAKUPU.IDENTKLIENTA)='TRUE')) AS ILOSCAA," &
                    "(SELECT (SUM((PROCM12 + PROCM13 + PROCM14 + PROCM15 + PROCM16 + PROCM17 + PROCM18 + PROCM19 + PROCM20 + PROCM21 + PROCM22 + PROCM23) / 12) / (SELECT COUNT(*) FROM KLIENCI WHERE (AKTYWNY='TRUE'))) FROM ANALIZYZAKUPU WHERE ((SELECT AKTYWNY FROM KLIENCI WHERE KLIENCI.IDENTKLIENTA=ANALIZYZAKUPU.IDENTKLIENTA)='TRUE')) AS ILOSCBB,"
            Else
                Label18.Text = "Sum.SPRZED.Wartoœci"
                dbSelectRoboczy &=
                    "(SELECT SUM(WARTZABM + WARTZA01 + WARTZA02 + WARTZA03 + WARTZA04 + WARTZA05 + WARTZA06 + WARTZA07 + WARTZA08 + WARTZA09 + WARTZA10 + WARTZA11) FROM ANALIZYZAKUPU WHERE ((SELECT AKTYWNY FROM KLIENCI WHERE KLIENCI.IDENTKLIENTA=ANALIZYZAKUPU.IDENTKLIENTA)='TRUE')) AS ILOSCAA," &
                    "(SELECT SUM(WARTZA12 + WARTZA13 + WARTZA14 + WARTZA15 + WARTZA16 + WARTZA17 + WARTZA18 + WARTZA19 + WARTZA20 + WARTZA21 + WARTZA22 + WARTZA23) FROM ANALIZYZAKUPU WHERE ((SELECT AKTYWNY FROM KLIENCI WHERE KLIENCI.IDENTKLIENTA=ANALIZYZAKUPU.IDENTKLIENTA)='TRUE')) AS ILOSCBB,"
            End If
        End If



        If AnalIlosciowo Then
            'Tylko Ilosc
            dbSelectRoboczy &=
                                                   "(SELECT SUM(ILOSZABM) FROM ANALIZYZAKUPU) AS ILZAKBM," &
                                                   "(SELECT SUM(ILOSZA01) FROM ANALIZYZAKUPU) AS ILZAK01," &
                                                   "(SELECT SUM(ILOSZA02) FROM ANALIZYZAKUPU) AS ILZAK02," &
                                                   "(SELECT SUM(ILOSZA03) FROM ANALIZYZAKUPU) AS ILZAK03," &
                                                   "(SELECT SUM(ILOSZA04) FROM ANALIZYZAKUPU) AS ILZAK04," &
                                                   "(SELECT SUM(ILOSZA05) FROM ANALIZYZAKUPU) AS ILZAK05," &
                                                   "(SELECT SUM(ILOSZA06) FROM ANALIZYZAKUPU) AS ILZAK06," &
                                                   "(SELECT SUM(ILOSZA07) FROM ANALIZYZAKUPU) AS ILZAK07," &
                                                   "(SELECT SUM(ILOSZA08) FROM ANALIZYZAKUPU) AS ILZAK08," &
                                                   "(SELECT SUM(ILOSZA09) FROM ANALIZYZAKUPU) AS ILZAK09," &
                                                   "(SELECT SUM(ILOSZA10) FROM ANALIZYZAKUPU) AS ILZAK10," &
                                                   "(SELECT SUM(ILOSZA11) FROM ANALIZYZAKUPU) AS ILZAK11," &
                                                 "(SELECT SUM(ILOSZA12) FROM ANALIZYZAKUPU) AS ILZAK12," &
                                                 "(SELECT SUM(ILOSZA13) FROM ANALIZYZAKUPU) AS ILZAK13," &
                                                 "(SELECT SUM(ILOSZA14) FROM ANALIZYZAKUPU) AS ILZAK14," &
                                                 "(SELECT SUM(ILOSZA15) FROM ANALIZYZAKUPU) AS ILZAK15," &
                                                 "(SELECT SUM(ILOSZA16) FROM ANALIZYZAKUPU) AS ILZAK16," &
                                                 "(SELECT SUM(ILOSZA17) FROM ANALIZYZAKUPU) AS ILZAK17," &
                                                 "(SELECT SUM(ILOSZA18) FROM ANALIZYZAKUPU) AS ILZAK18," &
                                                 "(SELECT SUM(ILOSZA19) FROM ANALIZYZAKUPU) AS ILZAK19," &
                                                 "(SELECT SUM(ILOSZA20) FROM ANALIZYZAKUPU) AS ILZAK20," &
                                                 "(SELECT SUM(ILOSZA21) FROM ANALIZYZAKUPU) AS ILZAK21," &
                                                 "(SELECT SUM(ILOSZA22) FROM ANALIZYZAKUPU) AS ILZAK22," &
                                                 "(SELECT SUM(ILOSZA23) FROM ANALIZYZAKUPU) AS ILZAK23," &
                                                 "(SELECT SUM(ILOSZA24) FROM ANALIZYZAKUPU) AS ILZAK24, "

        ElseIf AnalWartMarzy Then
            'wartosc marzy
            dbSelectRoboczy &=
                                                   "(SELECT SUM(MARZABM) FROM ANALIZYZAKUPU) AS ILZAKBM," &
                                                   "(SELECT SUM(MARZA01) FROM ANALIZYZAKUPU) AS ILZAK01," &
                                                   "(SELECT SUM(MARZA02) FROM ANALIZYZAKUPU) AS ILZAK02," &
                                                   "(SELECT SUM(MARZA03) FROM ANALIZYZAKUPU) AS ILZAK03," &
                                                   "(SELECT SUM(MARZA04) FROM ANALIZYZAKUPU) AS ILZAK04," &
                                                   "(SELECT SUM(MARZA05) FROM ANALIZYZAKUPU) AS ILZAK05," &
                                                   "(SELECT SUM(MARZA06) FROM ANALIZYZAKUPU) AS ILZAK06," &
                                                   "(SELECT SUM(MARZA07) FROM ANALIZYZAKUPU) AS ILZAK07," &
                                                   "(SELECT SUM(MARZA08) FROM ANALIZYZAKUPU) AS ILZAK08," &
                                                   "(SELECT SUM(MARZA09) FROM ANALIZYZAKUPU) AS ILZAK09," &
                                                   "(SELECT SUM(MARZA10) FROM ANALIZYZAKUPU) AS ILZAK10," &
                                                   "(SELECT SUM(MARZA11) FROM ANALIZYZAKUPU) AS ILZAK11," &
                                                 "(SELECT SUM(MARZA12) FROM ANALIZYZAKUPU) AS ILZAK12," &
                                                 "(SELECT SUM(MARZA13) FROM ANALIZYZAKUPU) AS ILZAK13," &
                                                 "(SELECT SUM(MARZA14) FROM ANALIZYZAKUPU) AS ILZAK14," &
                                                 "(SELECT SUM(MARZA15) FROM ANALIZYZAKUPU) AS ILZAK15," &
                                                 "(SELECT SUM(MARZA16) FROM ANALIZYZAKUPU) AS ILZAK16," &
                                                 "(SELECT SUM(MARZA17) FROM ANALIZYZAKUPU) AS ILZAK17," &
                                                 "(SELECT SUM(MARZA18) FROM ANALIZYZAKUPU) AS ILZAK18," &
                                                 "(SELECT SUM(MARZA19) FROM ANALIZYZAKUPU) AS ILZAK19," &
                                                 "(SELECT SUM(MARZA20) FROM ANALIZYZAKUPU) AS ILZAK20," &
                                                 "(SELECT SUM(MARZA21) FROM ANALIZYZAKUPU) AS ILZAK21," &
                                                 "(SELECT SUM(MARZA22) FROM ANALIZYZAKUPU) AS ILZAK22," &
                                                 "(SELECT SUM(MARZA23) FROM ANALIZYZAKUPU) AS ILZAK23," &
                                                 "(SELECT SUM(MARZA24) FROM ANALIZYZAKUPU) AS ILZAK24, "


        ElseIf AnalProcMarzy Then
            'procent marzy
            dbSelectRoboczy &=
                                                   "(SELECT SUM(PROCMBM) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAKBM," &
                                                   "(SELECT SUM(PROCM01) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK01," &
                                                   "(SELECT SUM(PROCM02) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK02," &
                                                   "(SELECT SUM(PROCM03) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK03," &
                                                   "(SELECT SUM(PROCM04) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK04," &
                                                   "(SELECT SUM(PROCM05) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK05," &
                                                   "(SELECT SUM(PROCM06) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK06," &
                                                   "(SELECT SUM(PROCM07) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK07," &
                                                   "(SELECT SUM(PROCM08) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK08," &
                                                   "(SELECT SUM(PROCM09) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK09," &
                                                   "(SELECT SUM(PROCM10) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK10," &
                                                   "(SELECT SUM(PROCM11) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK11," &
                                                 "(SELECT SUM(PROCM12) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK12," &
                                                 "(SELECT SUM(PROCM13) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK13," &
                                                 "(SELECT SUM(PROCM14) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK14," &
                                                 "(SELECT SUM(PROCM15) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK15," &
                                                 "(SELECT SUM(PROCM16) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK16," &
                                                 "(SELECT SUM(PROCM17) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK17," &
                                                 "(SELECT SUM(PROCM18) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK18," &
                                                 "(SELECT SUM(PROCM19) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK19," &
                                                 "(SELECT SUM(PROCM20) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK20," &
                                                 "(SELECT SUM(PROCM21) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK21," &
                                                 "(SELECT SUM(PROCM22) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK22," &
                                                 "(SELECT SUM(PROCM23) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK23," &
                                                 "(SELECT SUM(PROCM24) / (SELECT COUNT(*) FROM KLIENCI) FROM ANALIZYZAKUPU) AS ILZAK24, "


        Else
            'Tylko Wartosc
            dbSelectRoboczy &=
                                                   "(SELECT SUM(WARTZABM) FROM ANALIZYZAKUPU) AS ILZAKBM," &
                                                   "(SELECT SUM(WARTZA01) FROM ANALIZYZAKUPU) AS ILZAK01," &
                                                   "(SELECT SUM(WARTZA02) FROM ANALIZYZAKUPU) AS ILZAK02," &
                                                   "(SELECT SUM(WARTZA03) FROM ANALIZYZAKUPU) AS ILZAK03," &
                                                   "(SELECT SUM(WARTZA04) FROM ANALIZYZAKUPU) AS ILZAK04," &
                                                   "(SELECT SUM(WARTZA05) FROM ANALIZYZAKUPU) AS ILZAK05," &
                                                   "(SELECT SUM(WARTZA06) FROM ANALIZYZAKUPU) AS ILZAK06," &
                                                   "(SELECT SUM(WARTZA07) FROM ANALIZYZAKUPU) AS ILZAK07," &
                                                   "(SELECT SUM(WARTZA08) FROM ANALIZYZAKUPU) AS ILZAK08," &
                                                   "(SELECT SUM(WARTZA09) FROM ANALIZYZAKUPU) AS ILZAK09," &
                                                   "(SELECT SUM(WARTZA10) FROM ANALIZYZAKUPU) AS ILZAK10," &
                                                   "(SELECT SUM(WARTZA11) FROM ANALIZYZAKUPU) AS ILZAK11," &
                                                 "(SELECT SUM(WARTZA12) FROM ANALIZYZAKUPU) AS ILZAK12," &
                                                 "(SELECT SUM(WARTZA13) FROM ANALIZYZAKUPU) AS ILZAK13," &
                                                 "(SELECT SUM(WARTZA14) FROM ANALIZYZAKUPU) AS ILZAK14," &
                                                 "(SELECT SUM(WARTZA15) FROM ANALIZYZAKUPU) AS ILZAK15," &
                                                 "(SELECT SUM(WARTZA16) FROM ANALIZYZAKUPU) AS ILZAK16," &
                                                 "(SELECT SUM(WARTZA17) FROM ANALIZYZAKUPU) AS ILZAK17," &
                                                 "(SELECT SUM(WARTZA18) FROM ANALIZYZAKUPU) AS ILZAK18," &
                                                 "(SELECT SUM(WARTZA19) FROM ANALIZYZAKUPU) AS ILZAK19," &
                                                 "(SELECT SUM(WARTZA20) FROM ANALIZYZAKUPU) AS ILZAK20," &
                                                 "(SELECT SUM(WARTZA21) FROM ANALIZYZAKUPU) AS ILZAK21," &
                                                 "(SELECT SUM(WARTZA22) FROM ANALIZYZAKUPU) AS ILZAK22," &
                                                 "(SELECT SUM(WARTZA23) FROM ANALIZYZAKUPU) AS ILZAK23," &
                                                 "(SELECT SUM(WARTZA24) FROM ANALIZYZAKUPU) AS ILZAK24, "

        End If


        dbSelectRoboczy &=
                         "(SELECT SUM(PLANKROTKOOKRESOWY) FROM KLIENCI) AS SUMAPLANK, " &
                         "(SELECT SUM(PLANDLUGOOKRESOWY) FROM KLIENCI) AS SUMAPLAND "

        dbSelectRoboczy &= " FROM Klienci WHERE (AKTYWNY='TRUE') "


        DataSetRoboczy = New DataSet
        DataViewRoboczy = New DataView
        DataSetRoboczy.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
        Else
            sqlConnectionRoboczy = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelectRoboczy = New SqlClient.SqlCommand(dbSelectRoboczy, sqlConnectionRoboczy)
            'sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLRoboczy, sqlConnection)
            'sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLRoboczy, sqlConnection)
            'sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLRoboczy, sqlConnection)
            sqlDataAdapterRoboczy = New SqlClient.SqlDataAdapter(sqlCommandSelectRoboczy)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliRoboczy As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliRoboczy = sqlDataAdapterRoboczy.TableMappings.Add("Table", "TABELA_Roboczy")

            '' Add the RowUpdated event handler.
            'AddHandler sqlDataAdapter.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionRoboczy.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionRoboczy.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterRoboczy.FillSchema(DataSetRoboczy, SchemaType.Mapped, "TABELA_Roboczy")
                    'dbDataAdapterRoboczy.Fill(DataSetRoboczy)
                    'dbConnectionRoboczy.Close()
                Else
                    sqlDataAdapterRoboczy.FillSchema(DataSetRoboczy, SchemaType.Mapped, "TABELA_Roboczy")
                    sqlDataAdapterRoboczy.Fill(DataSetRoboczy)
                    sqlConnectionRoboczy.Close()
                End If

                lblPlanK.Text = ""
                lblPlanD.Text = ""

                val0011.Text = ""
                val1223.Text = ""
                Suma1 = 0
                Suma2 = 0



                If DataSetRoboczy.Tables(0).Rows.Count > 0 Then
                    lblPlanD.Text = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "SUMAPLAND").ToString("N0")
                    lblPlanK.Text = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "SUMAPLANK").ToString("N0")

                    If Chk101.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAKBM")
                        valM00.Text = Val.ToString("N2")
                        Suma1 += Val
                    End If
                    If Chk102.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK01")
                        valM01.Text = Val.ToString("N2")
                        Suma1 += Val
                    End If
                    If Chk103.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK02")
                        valM02.Text = Val.ToString("N2")
                        Suma1 += Val
                    End If
                    If Chk104.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK03")
                        valM03.Text = Val.ToString("N2")
                        Suma1 += Val
                    End If
                    If Chk105.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK04")
                        valM04.Text = Val.ToString("N2")
                        Suma1 += Val
                    End If
                    If Chk106.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK05")
                        valM05.Text = Val.ToString("N2")
                        Suma1 += Val
                    End If
                    If Chk107.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK06")
                        valM06.Text = Val.ToString("N2")
                        Suma1 += Val
                    End If
                    If Chk108.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK07")
                        valM07.Text = Val.ToString("N2")
                        Suma1 += Val
                    End If
                    If Chk109.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK08")
                        valM08.Text = Val.ToString("N2")
                        Suma1 += Val
                    End If
                    If Chk110.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK09")
                        valM09.Text = Val.ToString("N2")
                        Suma1 += Val
                    End If
                    If Chk111.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK10")
                        valM10.Text = Val.ToString("N2")
                        Suma1 += Val
                    End If
                    If Chk112.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK11")
                        valM11.Text = Val.ToString("N2")
                        Suma1 += Val
                    End If


                    If Chk201.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK12")
                        valM20.Text = Val.ToString("N2")
                        Suma2 += Val
                    End If
                    If Chk202.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK13")
                        valM21.Text = Val.ToString("N2")
                        Suma2 += Val
                    End If
                    If Chk203.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK14")
                        valM22.Text = Val.ToString("N2")
                        Suma2 += Val
                    End If
                    If Chk204.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK15")
                        valM23.Text = Val.ToString("N2")
                        Suma2 += Val
                    End If
                    If Chk205.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK16")
                        valM24.Text = Val.ToString("N2")
                        Suma2 += Val
                    End If
                    If Chk206.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK17")
                        valM25.Text = Val.ToString("N2")
                        Suma2 += Val
                    End If
                    If Chk207.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK18")
                        valM26.Text = Val.ToString("N2")
                        Suma2 += Val
                    End If
                    If Chk208.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK19")
                        valM27.Text = Val.ToString("N2")
                        Suma2 += Val
                    End If
                    If Chk209.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK20")
                        valM28.Text = Val.ToString("N2")
                        Suma2 += Val
                    End If
                    If Chk210.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK21")
                        valM29.Text = Val.ToString("N2")
                        Suma2 += Val
                    End If
                    If Chk211.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK22")
                        valM30.Text = Val.ToString("N2")
                        Suma2 += Val
                    End If
                    If Chk212.Checked Then
                        Val = GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILZAK23")
                        valM31.Text = Val.ToString("N2")
                        Suma2 += Val
                    End If

                    'val0011.Text = Suma1.ToString("N2")
                    'val1223.Text = Suma2.ToString("N2")
                    val0011.Text = (GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILOSCAA")).ToString("N2")
                    val1223.Text = (GetDblDRField(DataSetRoboczy.Tables(0).Rows(0), "ILOSCBB")).ToString("N2")

                End If



            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End Try
            '--------------------------------------------------
        End If








        Dim dbSelectObroty As String = ""

        Dim dbConnectionObroty As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectObroty As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterObroty As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionObroty As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectObroty As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterObroty As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetObroty As New DataSet
        Dim DataViewObroty As New DataView

        ''Obroty
        'Public oIdentObr As String            'IDENTKLIENTA     Text, 6
        'Public oDataObr As String             'DATA             Text,10
        'Public oLWartSprzedObr As Double      'LWARTSPRZED      Double
        'Public oLIloSprzedObr As Double       'LILOPRZED        Double
        'Public oLMarSprzedObr As Double       'LMARPRZED        Double
        'Public oAWartSprzedObr As Double      'AWARTSPRZED      Double
        'Public oAIloSprzedObr As Double       'AILOSPRZED       Double
        'Public oAMARSprzedObr As Double       'AMARSPRZED       Double
        'Public oLOWartSprzedObr As Double      'LOWARTSPRZED      Double
        'Public oLOIloSprzedObr As Double       'LOILOPRZED        Double
        'Public oLOMARSprzedObr As Double       'LOMARPRZED        Double
        'Public oAOWartSprzedObr As Double      'AOWARTSPRZED      Double
        'Public oAOIloSprzedObr As Double       'AOILOSPRZED       Double
        'Public oAOMARSprzedObr As Double       'AOMARSPRZED       Double
        'Public oNAJEMWartSprzedObr As Double      'NAJEMWARTSPRZED      Double
        'Public oNAJEMIloSprzedObr As Double       'NAJEMILOPRZED        Double
        'Public oNAJEMMARSprzedObr As Double       'NAJEMMARPRZED        Double
        'Public oSTRONYWartSprzedObr As Double      'STRONYWARTSPRZED      Double
        'Public oSTRONYIloSprzedObr As Double       'STRONYILOPRZED        Double
        'Public oSTRONYMARSprzedObr As Double       'STRONYMARPRZED        Double
        'Public oDRUKARKIWartSprzedObr As Double      'DRUKARKIWARTSPRZED      Double
        'Public oDRUKARKIIloSprzedObr As Double       'DRUKARKIILOPRZED        Double
        'Public oDRUKARKIMARSprzedObr As Double       'DRUKARKIMARPRZED        Double
        'Public oSKUPWartSprzedObr As Double      'SKUPWARTSPRZED      Double
        'Public oSKUPIloSprzedObr As Double       'SKUPILOPRZED        Double
        'Public oSKUPMARSprzedObr As Double       'SKUPMARPRZED        Double
        'Public oWersjaObr As Integer          'WERSJA           Integer



        Dim brok As String = Mid(SysData, 1, 4)
        dbSelectObroty = "SELECT " &
        "ISNULL((select sum(skupwartsprzed) from obroty where substring(data,1,7)='" & brok & "-01'),0) + ISNULL((select sum(skupwartsprzed) from obrotymies where miesiac='" & brok & "-01'),0) as sty," &
        "ISNULL((select sum(skupwartsprzed) from obroty where substring(data,1,7)='" & brok & "-02'),0) + ISNULL((select sum(skupwartsprzed) from obrotymies where miesiac='" & brok & "-02'),0) as lut," &
        "ISNULL((select sum(skupwartsprzed) from obroty where substring(data,1,7)='" & brok & "-03'),0) + ISNULL((select sum(skupwartsprzed) from obrotymies where miesiac='" & brok & "-03'),0) as mar," &
        "ISNULL((select sum(skupwartsprzed) from obroty where substring(data,1,7)='" & brok & "-04'),0) + ISNULL((select sum(skupwartsprzed) from obrotymies where miesiac='" & brok & "-04'),0) as kwi," &
        "ISNULL((select sum(skupwartsprzed) from obroty where substring(data,1,7)='" & brok & "-05'),0) + ISNULL((select sum(skupwartsprzed) from obrotymies where miesiac='" & brok & "-05'),0) as maj," &
        "ISNULL((select sum(skupwartsprzed) from obroty where substring(data,1,7)='" & brok & "-06'),0) + ISNULL((select sum(skupwartsprzed) from obrotymies where miesiac='" & brok & "-06'),0) as cze," &
        "ISNULL((select sum(skupwartsprzed) from obroty where substring(data,1,7)='" & brok & "-07'),0) + ISNULL((select sum(skupwartsprzed) from obrotymies where miesiac='" & brok & "-07'),0) as lip," &
        "ISNULL((select sum(skupwartsprzed) from obroty where substring(data,1,7)='" & brok & "-08'),0) + ISNULL((select sum(skupwartsprzed) from obrotymies where miesiac='" & brok & "-08'),0) as sie," &
        "ISNULL((select sum(skupwartsprzed) from obroty where substring(data,1,7)='" & brok & "-09'),0) + ISNULL((select sum(skupwartsprzed) from obrotymies where miesiac='" & brok & "-09'),0) as wrz," &
        "ISNULL((select sum(skupwartsprzed) from obroty where substring(data,1,7)='" & brok & "-10'),0) + ISNULL((select sum(skupwartsprzed) from obrotymies where miesiac='" & brok & "-10'),0) as paz," &
        "ISNULL((select sum(skupwartsprzed) from obroty where substring(data,1,7)='" & brok & "-11'),0) + ISNULL((select sum(skupwartsprzed) from obrotymies where miesiac='" & brok & "-11'),0) as lis," &
        "ISNULL((select sum(skupwartsprzed) from obroty where substring(data,1,7)='" & brok & "-12'),0) + ISNULL((select sum(skupwartsprzed) from obrotymies where miesiac='" & brok & "-12'),0) as gru"

        DataSetObroty = New DataSet
        DataViewObroty = New DataView
        DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then

        Else
            sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionKlienci)
            sqlCommandSelectObroty = New SqlClient.SqlCommand(dbSelectObroty, sqlConnectionObroty)
            'sqlCommandDelete = New SqlClient.SqlCommand(sDeleteSQLObroty, sqlConnection)
            'sqlCommandUpdate = New SqlClient.SqlCommand(sUpdateSQLObroty, sqlConnection)
            'sqlCommandInsert = New SqlClient.SqlCommand(sInsertSQLObroty, sqlConnection)
            sqlDataAdapterObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliObroty As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliObroty = sqlDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")

            '' Add the RowUpdated event handler.
            'AddHandler sqlDataAdapter.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObroty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = sqlConnectionObroty.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    'dbDataAdapterObroty.Fill(DataSetObroty)
                    'dbConnectionObroty.Close()
                Else
                    sqlDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    sqlDataAdapterObroty.Fill(DataSetObroty)
                    sqlConnectionObroty.Close()
                End If




            Catch Ex As System.Exception
                'ViewInLog(ex, Me.Name(), Nothing)
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End Try
            '--------------------------------------------------
        End If
        Me.Cursor = Cursors.Default
    End Sub



    Private Sub PokazSumy(ByRef dvKlienci As DataView)
        Dim Val As Double = 0
        Dim Suma1 As Double = 0
        Dim Suma2 As Double = 0
        Dim SumaPlanK As Double = 0
        Dim SumaPlanD As Double = 0

        Me.Cursor = Cursors.WaitCursor

        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne")
        'cbxCoAnalizowac.Items.Add("Analizuj Iloœciowo-Sumuj Iloœci")
        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne")
        'cbxCoAnalizowac.Items.Add("Analizuj Wartoœciowo-Sumuj Wartoœci")
        'cbxCoAnalizowac.Items.Add("Analizuj wartoœæ Mar¿y")
        'cbxCoAnalizowac.Items.Add("Analizuj procent Mar¿y")


        Dim AnalIlosciowo As Boolean = False
        Dim AnalWartMarzy As Boolean = False
        Dim AnalProcMarzy As Boolean = False
        Dim SumujAktywnosci As Boolean = False

        Select Case cbxCoAnalizowac.Text
            Case "Analizuj Iloœciowo-Sumuj Aktywnoœci miesiêczne"
                AnalIlosciowo = True
                SumujAktywnosci = True
                AnalWartMarzy = False
                AnalProcMarzy = False
            Case "Analizuj Iloœciowo-Sumuj Iloœci"
                AnalIlosciowo = True
                SumujAktywnosci = False
                AnalWartMarzy = False
                AnalProcMarzy = False
            Case "Analizuj Wartoœciowo-Sumuj Aktywnoœci miesiêczne"
                AnalIlosciowo = False
                SumujAktywnosci = True
                AnalWartMarzy = False
                AnalProcMarzy = False
            Case "Analizuj Wartoœciowo-Sumuj Wartoœci"
                AnalIlosciowo = False
                SumujAktywnosci = False
                AnalWartMarzy = False
                AnalProcMarzy = False
            Case "Analizuj wartoœæ Mar¿y"
                AnalIlosciowo = False
                SumujAktywnosci = False
                AnalWartMarzy = True
                AnalProcMarzy = False
            Case "Analizuj procent Mar¿y"
                AnalIlosciowo = False
                SumujAktywnosci = False
                AnalWartMarzy = False
                AnalProcMarzy = True
        End Select


        Try
            val0011.Text = ""
            val1223.Text = ""
            Suma1 = 0
            Suma2 = 0
            SumaPlanK = 0
            SumaPlanD = 0

            lblM00.Text = InnyMiesiac(Okres1, 0)
            lblM01.Text = InnyMiesiac(Okres1, -1)
            lblM02.Text = InnyMiesiac(Okres1, -2)
            lblM03.Text = InnyMiesiac(Okres1, -3)
            lblM04.Text = InnyMiesiac(Okres1, -4)
            lblM05.Text = InnyMiesiac(Okres1, -5)
            lblM06.Text = InnyMiesiac(Okres1, -6)
            lblM07.Text = InnyMiesiac(Okres1, -7)
            lblM08.Text = InnyMiesiac(Okres1, -8)
            lblM09.Text = InnyMiesiac(Okres1, -9)
            lblM10.Text = InnyMiesiac(Okres1, -10)
            lblM11.Text = InnyMiesiac(Okres1, -11)

            lblM20.Text = InnyMiesiac(Okres2, 0)
            lblM21.Text = InnyMiesiac(Okres2, -1)
            lblM22.Text = InnyMiesiac(Okres2, -2)
            lblM23.Text = InnyMiesiac(Okres2, -3)
            lblM24.Text = InnyMiesiac(Okres2, -4)
            lblM25.Text = InnyMiesiac(Okres2, -5)
            lblM26.Text = InnyMiesiac(Okres2, -6)
            lblM27.Text = InnyMiesiac(Okres2, -7)
            lblM28.Text = InnyMiesiac(Okres2, -8)
            lblM29.Text = InnyMiesiac(Okres2, -9)
            lblM30.Text = InnyMiesiac(Okres2, -10)
            lblM31.Text = InnyMiesiac(Okres2, -11)

            valM00.Text = "0"
            valM01.Text = "0"
            valM02.Text = "0"
            valM03.Text = "0"
            valM04.Text = "0"
            valM05.Text = "0"
            valM06.Text = "0"
            valM07.Text = "0"
            valM08.Text = "0"
            valM09.Text = "0"
            valM10.Text = "0"
            valM11.Text = "0"

            valM20.Text = "0"
            valM21.Text = "0"
            valM22.Text = "0"
            valM23.Text = "0"
            valM24.Text = "0"
            valM25.Text = "0"
            valM26.Text = "0"
            valM27.Text = "0"
            valM28.Text = "0"
            valM29.Text = "0"
            valM30.Text = "0"
            valM31.Text = "0"
            '------------

            If dvKlienci.Count > 0 Then
                Dim i As Integer = 0
                For i = 0 To dvKlienci.Count - 1
                    If GetLogDRVField(dvKlienci.Item(i), "AKTYWNY") Then
                        'Public oPlanDlugookresowyKli As Integer         'PLANDLUGOOKRESOWY     Integer
                        'Public oPlanKrotkookresowyKli As Integer        'PLANKROTKOOKRESOWY    Integer
                        SumaPlanK += GetDblDRVField(dvKlienci.Item(i), "PLANKROTKOOKRESOWY")
                        SumaPlanD += GetDblDRVField(dvKlienci.Item(i), "PLANDLUGOOKRESOWY")
                        '---------------------------
                        If Chk101.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILOSCBM")
                            valM00.Text = (CDbl(valM00.Text) + Val).ToString("N0")
                            Suma1 += Val
                        End If
                        If Chk102.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK01")
                            valM01.Text = (CDbl(valM01.Text) + Val).ToString("N0")
                            Suma1 += Val
                        End If
                        If Chk103.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK02")
                            valM02.Text = (CDbl(valM02.Text) + Val).ToString("N0")
                            Suma1 += Val
                        End If
                        If Chk104.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK03")
                            valM03.Text = (CDbl(valM03.Text) + Val).ToString("N0")
                            Suma1 += Val
                        End If
                        If Chk105.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK04")
                            valM04.Text = (CDbl(valM04.Text) + Val).ToString("N0")
                            Suma1 += Val
                        End If
                        If Chk106.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK05")
                            valM05.Text = (CDbl(valM05.Text) + Val).ToString("N0")
                            Suma1 += Val
                        End If
                        If Chk107.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK06")
                            valM06.Text = (CDbl(valM06.Text) + Val).ToString("N0")
                            Suma1 += Val
                        End If
                        If Chk108.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK07")
                            valM07.Text = (CDbl(valM07.Text) + Val).ToString("N0")
                            Suma1 += Val
                        End If
                        If Chk109.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK08")
                            valM08.Text = (CDbl(valM08.Text) + Val).ToString("N0")
                            Suma1 += Val
                        End If
                        If Chk110.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK09")
                            valM09.Text = (CDbl(valM09.Text) + Val).ToString("N0")
                            Suma1 += Val
                        End If
                        If Chk111.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK10")
                            valM10.Text = (CDbl(valM10.Text) + Val).ToString("N0")
                            Suma1 += Val
                        End If
                        If Chk112.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK11")
                            valM11.Text = (CDbl(valM11.Text) + Val).ToString("N0")
                            Suma1 += Val
                        End If


                        If Chk201.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK12")
                            valM20.Text = (CDbl(valM20.Text) + Val).ToString("N0")
                            Suma2 += Val
                        End If
                        If Chk202.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK13")
                            valM21.Text = (CDbl(valM21.Text) + Val).ToString("N0")
                            Suma2 += Val
                        End If
                        If Chk203.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK14")
                            valM22.Text = (CDbl(valM22.Text) + Val).ToString("N0")
                            Suma2 += Val
                        End If
                        If Chk204.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK15")
                            valM23.Text = (CDbl(valM23.Text) + Val).ToString("N0")
                            Suma2 += Val
                        End If
                        If Chk205.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK16")
                            valM24.Text = (CDbl(valM24.Text) + Val).ToString("N0")
                            Suma2 += Val
                        End If
                        If Chk206.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK17")
                            valM25.Text = (CDbl(valM25.Text) + Val).ToString("N0")
                            Suma2 += Val
                        End If
                        If Chk207.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK18")
                            valM26.Text = (CDbl(valM26.Text) + Val).ToString("N0")
                            Suma2 += Val
                        End If
                        If Chk208.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK19")
                            valM27.Text = (CDbl(valM27.Text) + Val).ToString("N0")
                            Suma2 += Val
                        End If
                        If Chk209.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK20")
                            valM28.Text = (CDbl(valM28.Text) + Val).ToString("N0")
                            Suma2 += Val
                        End If
                        If Chk210.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK21")
                            valM29.Text = (CDbl(valM29.Text) + Val).ToString("N0")
                            Suma2 += Val
                        End If
                        If Chk211.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK22")
                            valM30.Text = (CDbl(valM30.Text) + Val).ToString("N0")
                            Suma2 += Val
                        End If
                        If Chk212.Checked Then
                            Val = GetDblDRVField(dvKlienci.Item(i), "ILZAK23")
                            valM31.Text = (CDbl(valM31.Text) + Val).ToString("N0")
                            Suma2 += Val
                        End If

                        val0011.Text = Suma1.ToString("N0")
                        val1223.Text = Suma2.ToString("N0")
                        '------------------
                    End If
                Next
                lblPlanD.Text = SumaPlanD.ToString("N0")
                lblPlanK.Text = SumaPlanK.ToString("N0")

            End If



        Catch Ex As System.Exception
            'ViewInLog(ex, Me.Name(), Nothing)
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message & vbCrLf & Ex.StackTrace, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            Me.Cursor = Cursors.Default
        End Try

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








    '*************************************************
    '** obs³uga menu kontekstowego
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
        ZapiszParametryAnalizy()
        OdswiezBaze()
        PokazSumy(DataViewKlienci)
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

            'aktualizuj Katalog klientow
            CzyscFiltryDGV(Me.Panel4,
                        dagKlienci,
                        DataViewKlienci.Table.Columns.Count,
                        DataViewKlienci,
                        stbKlienci)

            'CzyscFiltry(Me.panel4,
            '            dagKlienci,
            '            DataViewKlienci.Table.Columns.Count,
            '            DataViewKlienci,
            '            stbKlienci,
            '            Trim(oSchematFiltrowania))

            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(79).ColumnName, "Marker  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(80).ColumnName, "Marker P.  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(81).ColumnName, "Aktywny  ", 50, True)

            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Panel4.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi079" Or ctrl.Name = "txtFi080" Then
                        ctrl.Text = "TRUE"
                        LiMark += 1
                        If LiMark = 2 Then
                            Exit For
                        End If
                    End If
                End If
            Next
            dagKlienci.Update()
            '------------------------------
            PokazSumy(DataViewKlienci)
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

            'aktualizuj Katalog klientow
            CzyscFiltryDGV(Me.Panel4,
                        dagKlienci,
                        DataViewKlienci.Table.Columns.Count,
                        DataViewKlienci,
                        stbKlienci)

            'CzyscFiltry(Me.panel4,
            '            dagKlienci,
            '            DataViewKlienci.Table.Columns.Count,
            '            DataViewKlienci,
            '            stbKlienci,
            '            Trim(oSchematFiltrowania))

            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(79).ColumnName, "Marker  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(80).ColumnName, "Marker P.  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(81).ColumnName, "Aktywny  ", 50, True)

            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Panel4.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi079" Or ctrl.Name = "txtFi080" Then
                        ctrl.Text = "TRUE"
                        LiMark += 1
                        If LiMark = 2 Then
                            Exit For
                        End If
                    End If
                End If
            Next
            dagKlienci.Update()
            '------------------------------
            PokazSumy(DataViewKlienci)
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

            'aktualizuj Katalog klientow
            CzyscFiltryDGV(Me.Panel4,
                        dagKlienci,
                        DataViewKlienci.Table.Columns.Count,
                        DataViewKlienci,
                        stbKlienci)

            'CzyscFiltry(Me.panel4,
            '            dagKlienci,
            '            DataViewKlienci.Table.Columns.Count,
            '            DataViewKlienci,
            '            stbKlienci,
            '            Trim(oSchematFiltrowania))

            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(79).ColumnName, "Marker  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(80).ColumnName, "Marker P.  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(81).ColumnName, "Aktywny  ", 50, True)

            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Panel4.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi079" Or ctrl.Name = "txtFi080" Then
                        ctrl.Text = "TRUE"
                        LiMark += 1
                        If LiMark = 2 Then
                            Exit For
                        End If
                    End If
                End If
            Next
            dagKlienci.Update()
            '------------------------------
            PokazSumy(DataViewKlienci)
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

            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(79).ColumnName, "Marker  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(80).ColumnName, "Marker P.  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(81).ColumnName, "Aktywny  ", 50, True)

            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Panel4.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi079" Or ctrl.Name = "txtFi080" Then
                        ctrl.Text = "TRUE"
                        LiMark += 1
                        If LiMark = 2 Then
                            Exit For
                        End If
                    End If
                End If
            Next
            dagKlienci.Update()
            '------------------------------
            PokazSumy(DataViewKlienci)
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

            'aktualizuj Katalog klientow
            CzyscFiltryDGV(Me.Panel4,
                        dagKlienci,
                        DataViewKlienci.Table.Columns.Count,
                        DataViewKlienci,
                        stbKlienci)

            'CzyscFiltry(Me.panel4,
            '            dagKlienci,
            '            DataViewKlienci.Table.Columns.Count,
            '            DataViewKlienci,
            '            stbKlienci,
            '            Trim(oSchematFiltrowania))

            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(79).ColumnName, "Marker  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(80).ColumnName, "Marker P.  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(81).ColumnName, "Aktywny  ", 50, True)

            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Panel4.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi079" Or ctrl.Name = "txtFi080" Then
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
            '------------------------------
            PokazSumy(DataViewKlienci)
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub menuEWybierzWgWzrostu_Click(sender As Object, e As EventArgs) Handles menuEWybierzWgWzrostu.Click
        oWybranoAkcjeMarketingowa = False
        'definiuj delegat
        Dim deleg As delegateAktualizuj = AddressOf AktualizujBaze

        Dim Form As New AkcjaMarketingowaWybierzWgObrotow(DataSetKlienci, cbxCoAnalizowac.Text, deleg)
        Form.ShowDialog()

        If oWybranoAkcjeMarketingowa Then
            Me.Cursor = Cursors.WaitCursor
            oLokalnyFiltr = ""
            oNazwaSchematu = ""
            oSchematFiltrowania = ""
            stbKlienci.Panels(4).Text = "Szablon : "
            OdswiezBaze()

            'aktualizuj Katalog klientow
            CzyscFiltryDGV(Me.Panel4,
                        dagKlienci,
                        DataViewKlienci.Table.Columns.Count,
                        DataViewKlienci,
                        stbKlienci)

            'CzyscFiltry(Me.panel4,
            '            dagKlienci,
            '            DataViewKlienci.Table.Columns.Count,
            '            DataViewKlienci,
            '            stbKlienci,
            '            Trim(oSchematFiltrowania))

            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(79).ColumnName, "Marker  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(80).ColumnName, "Marker P.  ", 50, True)
            'LogColumnView(dagKlienci, DataSetKlienci.Tables(0).Columns(81).ColumnName, "Aktywny  ", 50, True)

            Dim ctrl As Control
            Dim LiMark As Integer = 0
            For Each ctrl In Me.Panel4.Controls
                If TypeOf (ctrl) Is System.Windows.Forms.TextBox Then
                    If ctrl.Name = "txtFi079" Or ctrl.Name = "txtFi080" Then
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
            '------------------------------
            PokazSumy(DataViewKlienci)
            Me.Cursor = Cursors.Default
        End If
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

    Private Sub menuZmienWskAkt_Click(sender As Object, e As EventArgs) Handles menuZmienWskAkt.Click
        ZmianaWskAkt()
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

        If IO.File.Exists(oKatParametry & "\" & oPlikSzerokosciKolumnAnaliza) Then
            IO.File.Delete(oKatParametry & "\" & oPlikSzerokosciKolumnAnaliza)
        End If

        'zapisz do pliku tekstowego
        FileNum = FreeFile()    'kolejny nr otwartego zbioru
        FileOpen(FileNum, oKatParametry & "\" & oPlikSzerokosciKolumnAnaliza, OpenMode.Output)
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

        If IO.File.Exists(oKatParametry & "\" & oPlikSzerokosciKolumnAnaliza) Then
            FileNum = FreeFile()    'kolejny nr otwartego zbioru
            FileOpen(FileNum, oKatParametry & "\" & oPlikSzerokosciKolumnAnaliza, OpenMode.Input)
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

    Private Sub cmdZaznacz2_Click(sender As Object, e As EventArgs) Handles cmdZaznacz2.Click
        Chk101.Checked = True
        Chk102.Checked = True
        Chk103.Checked = True
        Chk104.Checked = True
        Chk105.Checked = True
        Chk106.Checked = True
        Chk107.Checked = True
        Chk108.Checked = True
        Chk109.Checked = True
        Chk110.Checked = True
        Chk111.Checked = True
        Chk112.Checked = True
    End Sub

    Private Sub cmdOdznacz2_Click(sender As Object, e As EventArgs) Handles cmdOdznacz2.Click
        Chk101.Checked = False
        Chk102.Checked = False
        Chk103.Checked = False
        Chk104.Checked = False
        Chk105.Checked = False
        Chk106.Checked = False
        Chk107.Checked = False
        Chk108.Checked = False
        Chk109.Checked = False
        Chk110.Checked = False
        Chk111.Checked = False
        Chk112.Checked = False
    End Sub


    Private Sub cmdZaznacz3_Click(sender As Object, e As EventArgs) Handles cmdZaznacz3.Click
        Chk201.Checked = True
        Chk202.Checked = True
        Chk203.Checked = True
        Chk204.Checked = True
        Chk205.Checked = True
        Chk206.Checked = True
        Chk207.Checked = True
        Chk208.Checked = True
        Chk209.Checked = True
        Chk210.Checked = True
        Chk211.Checked = True
        Chk212.Checked = True
    End Sub

    Private Sub cmdOdznacz3_Click(sender As Object, e As EventArgs) Handles cmdOdznacz3.Click
        Chk201.Checked = False
        Chk202.Checked = False
        Chk203.Checked = False
        Chk204.Checked = False
        Chk205.Checked = False
        Chk206.Checked = False
        Chk207.Checked = False
        Chk208.Checked = False
        Chk209.Checked = False
        Chk210.Checked = False
        Chk211.Checked = False
        Chk212.Checked = False
    End Sub

End Class
