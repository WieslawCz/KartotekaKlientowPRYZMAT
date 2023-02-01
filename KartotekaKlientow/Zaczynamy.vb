Imports Microsoft.Win32     ' odczyt danych z rejestru

Public Class Zaczynamy
    Inherits System.Windows.Forms.Form
    'Inherits KartotekaKlientow.Softart.myDropShadowForm

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
    Friend WithEvents tmrStart As System.Windows.Forms.Timer
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents picTlo As System.Windows.Forms.PictureBox
    Friend WithEvents lblTytul As System.Windows.Forms.Label
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents lblWersja As System.Windows.Forms.Label
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuKlienci As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuProgram As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSystem2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSysInfo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuInternet As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuNotatnik As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuKalkulator As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuKalendarz As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuKoniec As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOAutorach As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOProgramie As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuLokalne As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuKatalogiISlowniki As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuBazaDanych As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSzabony As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuBackup As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRestore As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuZTofi As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuKartotekaKlientow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAnalizaZakupuKlientow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator7 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuDoAnalizy As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuKatalogPracownik�w As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuKatalogAkcjiMarketingowych As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuKatalogOs�bKontaktowych As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuKatalogObrot�w As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuKatalogObrot�wMiesi�cznych As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuSprawdzeniePoprawnosci As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SprawdzenieKataloguOs�bKontaktowychToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SprawdzenieKataloguObrot�wToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SprawdzenieKataloguObrot�wMiesi�cznychToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAktualizacjaStruktur As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents KatalogKontakt�wHandlowychToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents S�ownikRodzaj�wKontakt�wToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuParametry As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents menuAnalizaKontakt�wHandlowych As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RaportAktywno�ciPrzedstawicieliHandlowychToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuDaneDoRaportu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SprawdzenieKataloguKontakt�wHandlowychToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuHarmonogram As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator10 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator12 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator13 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents KatalogCoDalejToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuWartZakPopRok As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuZmianaPracownika As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuZmianaHasla As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuSprawdzenieKataloguCoDalej As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tstExportDanych As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents menuAnalizaWskaznikowa As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuAnalizaKlienci80 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents picSoftart As PictureBox
    Friend WithEvents lblAutor As Label
    Friend WithEvents menuAnalizaMarketingowa As ToolStripMenuItem
    Friend WithEvents menuRankingWskaznikow As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents mnuKorektaRZU As ToolStripMenuItem
    Friend WithEvents S�ownikCoDalejToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents menuZmianaNazwKatKlientow As ToolStripMenuItem
    Friend WithEvents menuZmianaNazwKatKontaktow As ToolStripMenuItem
    Friend WithEvents KatalogLogAktywnosci As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator8 As ToolStripSeparator
    Friend WithEvents menuKatalogBranze As ToolStripMenuItem
    Friend WithEvents menuKatalogPodbranze As ToolStripMenuItem
    Friend WithEvents mnuMiesieczne As System.Windows.Forms.ToolStripMenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Zaczynamy))
        Me.tmrStart = New System.Windows.Forms.Timer(Me.components)
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.lblWersja = New System.Windows.Forms.Label()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.lblTytul = New System.Windows.Forms.Label()
        Me.picTlo = New System.Windows.Forms.PictureBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.mnuKlienci = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuKartotekaKlientow = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAnalizaZakupuKlientow = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuAnalizaKontakt�wHandlowych = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuHarmonogram = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator10 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuDaneDoRaportu = New System.Windows.Forms.ToolStripMenuItem()
        Me.RaportAktywno�ciPrzedstawicieliHandlowychToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tstExportDanych = New System.Windows.Forms.ToolStripSeparator()
        Me.menuAnalizaWskaznikowa = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuAnalizaKlienci80 = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuAnalizaMarketingowa = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuRankingWskaznikow = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuProgram = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOAutorach = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOProgramie = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuLokalne = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuParametry = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuKatalogiISlowniki = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuKatalogPracownik�w = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuKatalogAkcjiMarketingowych = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuKatalogOs�bKontaktowych = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuKatalogBranze = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuKatalogPodbranze = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuKatalogObrot�w = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuKatalogObrot�wMiesi�cznych = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator12 = New System.Windows.Forms.ToolStripSeparator()
        Me.S�ownikRodzaj�wKontakt�wToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.KatalogKontakt�wHandlowychToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator13 = New System.Windows.Forms.ToolStripSeparator()
        Me.S�ownikCoDalejToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.KatalogCoDalejToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator8 = New System.Windows.Forms.ToolStripSeparator()
        Me.KatalogLogAktywnosci = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuBazaDanych = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuAktualizacjaStruktur = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSprawdzeniePoprawnosci = New System.Windows.Forms.ToolStripMenuItem()
        Me.SprawdzenieKataloguOs�bKontaktowychToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SprawdzenieKataloguObrot�wToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SprawdzenieKataloguObrot�wMiesi�cznychToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SprawdzenieKataloguKontakt�wHandlowychToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuSprawdzenieKataloguCoDalej = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuZmianaNazwKatKlientow = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuZmianaNazwKatKontaktow = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuKorektaRZU = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuSzabony = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuBackup = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuRestore = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuZTofi = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMiesieczne = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuDoAnalizy = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuWartZakPopRok = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSystem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuSysInfo = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuInternet = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuNotatnik = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuKalkulator = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuKalendarz = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuZmianaPracownika = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuZmianaHasla = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuKoniec = New System.Windows.Forms.ToolStripMenuItem()
        Me.picSoftart = New System.Windows.Forms.PictureBox()
        Me.lblAutor = New System.Windows.Forms.Label()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picTlo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.picSoftart, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tmrStart
        '
        '
        'lblWersja
        '
        Me.lblWersja.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblWersja.ForeColor = System.Drawing.Color.Purple
        Me.lblWersja.Location = New System.Drawing.Point(14, 139)
        Me.lblWersja.Name = "lblWersja"
        Me.lblWersja.Size = New System.Drawing.Size(255, 20)
        Me.lblWersja.TabIndex = 15
        Me.lblWersja.Text = "Wersja 1.0 (listopad 2004)"
        Me.lblWersja.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'picLogo
        '
        Me.picLogo.Image = CType(resources.GetObject("picLogo.Image"), System.Drawing.Image)
        Me.picLogo.Location = New System.Drawing.Point(61, 105)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(152, 31)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 14
        Me.picLogo.TabStop = False
        '
        'lblTytul
        '
        Me.lblTytul.BackColor = System.Drawing.SystemColors.Control
        Me.lblTytul.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTytul.Font = New System.Drawing.Font("Arial", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTytul.ForeColor = System.Drawing.Color.Navy
        Me.lblTytul.Location = New System.Drawing.Point(12, 33)
        Me.lblTytul.Name = "lblTytul"
        Me.lblTytul.Size = New System.Drawing.Size(263, 129)
        Me.lblTytul.TabIndex = 13
        Me.lblTytul.Text = "Kartoteka Klient�w firmy"
        Me.lblTytul.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'picTlo
        '
        Me.picTlo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.picTlo.BackColor = System.Drawing.Color.Silver
        Me.picTlo.Image = CType(resources.GetObject("picTlo.Image"), System.Drawing.Image)
        Me.picTlo.Location = New System.Drawing.Point(0, 24)
        Me.picTlo.Name = "picTlo"
        Me.picTlo.Size = New System.Drawing.Size(481, 381)
        Me.picTlo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picTlo.TabIndex = 11
        Me.picTlo.TabStop = False
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuKlienci, Me.mnuProgram, Me.mnuSystem2, Me.ToolStripMenuItem2, Me.mnuKoniec})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(481, 24)
        Me.MenuStrip1.TabIndex = 16
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'mnuKlienci
        '
        Me.mnuKlienci.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuKartotekaKlientow, Me.mnuAnalizaZakupuKlientow, Me.ToolStripSeparator9, Me.menuAnalizaKontakt�wHandlowych, Me.menuHarmonogram, Me.ToolStripSeparator10, Me.mnuDaneDoRaportu, Me.RaportAktywno�ciPrzedstawicieliHandlowychToolStripMenuItem, Me.tstExportDanych, Me.menuAnalizaWskaznikowa, Me.menuAnalizaKlienci80, Me.menuAnalizaMarketingowa, Me.menuRankingWskaznikow})
        Me.mnuKlienci.Name = "mnuKlienci"
        Me.mnuKlienci.Size = New System.Drawing.Size(54, 20)
        Me.mnuKlienci.Text = "Klienci"
        '
        'mnuKartotekaKlientow
        '
        Me.mnuKartotekaKlientow.Name = "mnuKartotekaKlientow"
        Me.mnuKartotekaKlientow.Size = New System.Drawing.Size(558, 22)
        Me.mnuKartotekaKlientow.Text = "Kartoteka Klient�w"
        '
        'mnuAnalizaZakupuKlientow
        '
        Me.mnuAnalizaZakupuKlientow.Name = "mnuAnalizaZakupuKlientow"
        Me.mnuAnalizaZakupuKlientow.Size = New System.Drawing.Size(558, 22)
        Me.mnuAnalizaZakupuKlientow.Text = "Analiza zakup�w klient�w"
        '
        'ToolStripSeparator9
        '
        Me.ToolStripSeparator9.Name = "ToolStripSeparator9"
        Me.ToolStripSeparator9.Size = New System.Drawing.Size(555, 6)
        '
        'menuAnalizaKontakt�wHandlowych
        '
        Me.menuAnalizaKontakt�wHandlowych.Name = "menuAnalizaKontakt�wHandlowych"
        Me.menuAnalizaKontakt�wHandlowych.Size = New System.Drawing.Size(558, 22)
        Me.menuAnalizaKontakt�wHandlowych.Text = "Analiza kontakt�w handlowych z klientami"
        '
        'menuHarmonogram
        '
        Me.menuHarmonogram.Name = "menuHarmonogram"
        Me.menuHarmonogram.Size = New System.Drawing.Size(558, 22)
        Me.menuHarmonogram.Text = "Harmonogram planowanych kontakt�w z klientami"
        Me.menuHarmonogram.Visible = False
        '
        'ToolStripSeparator10
        '
        Me.ToolStripSeparator10.Name = "ToolStripSeparator10"
        Me.ToolStripSeparator10.Size = New System.Drawing.Size(555, 6)
        '
        'mnuDaneDoRaportu
        '
        Me.mnuDaneDoRaportu.Name = "mnuDaneDoRaportu"
        Me.mnuDaneDoRaportu.Size = New System.Drawing.Size(558, 22)
        Me.mnuDaneDoRaportu.Text = "Wprowadzanie danych uzupe�niaj�cych do Raportu Aktywno�ci Przedstawicieli Handlow" &
    "ych"
        '
        'RaportAktywno�ciPrzedstawicieliHandlowychToolStripMenuItem
        '
        Me.RaportAktywno�ciPrzedstawicieliHandlowychToolStripMenuItem.Name = "RaportAktywno�ciPrzedstawicieliHandlowychToolStripMenuItem"
        Me.RaportAktywno�ciPrzedstawicieliHandlowychToolStripMenuItem.Size = New System.Drawing.Size(558, 22)
        Me.RaportAktywno�ciPrzedstawicieliHandlowychToolStripMenuItem.Text = "Raport Aktywno�ci Przedstawicieli Handlowych"
        '
        'tstExportDanych
        '
        Me.tstExportDanych.Name = "tstExportDanych"
        Me.tstExportDanych.Size = New System.Drawing.Size(555, 6)
        '
        'menuAnalizaWskaznikowa
        '
        Me.menuAnalizaWskaznikowa.Name = "menuAnalizaWskaznikowa"
        Me.menuAnalizaWskaznikowa.Size = New System.Drawing.Size(558, 22)
        Me.menuAnalizaWskaznikowa.Text = "Eksport danych do Raportu EXCEL - Analiza wska�nikowa"
        '
        'menuAnalizaKlienci80
        '
        Me.menuAnalizaKlienci80.Name = "menuAnalizaKlienci80"
        Me.menuAnalizaKlienci80.Size = New System.Drawing.Size(558, 22)
        Me.menuAnalizaKlienci80.Text = "Eksport danych do Raportu EXCEL - Analiza Klienci 80%"
        '
        'menuAnalizaMarketingowa
        '
        Me.menuAnalizaMarketingowa.Name = "menuAnalizaMarketingowa"
        Me.menuAnalizaMarketingowa.Size = New System.Drawing.Size(558, 22)
        Me.menuAnalizaMarketingowa.Text = "Eksport danych do Raportu EXCEL - Analiza marketingowa"
        '
        'menuRankingWskaznikow
        '
        Me.menuRankingWskaznikow.Name = "menuRankingWskaznikow"
        Me.menuRankingWskaznikow.Size = New System.Drawing.Size(558, 22)
        Me.menuRankingWskaznikow.Text = "Eksport danych do Raportu EXCEL - Ranking wska�nik�w"
        '
        'mnuProgram
        '
        Me.mnuProgram.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuOAutorach, Me.mnuOProgramie, Me.ToolStripSeparator2, Me.mnuLokalne, Me.menuParametry, Me.ToolStripSeparator5, Me.mnuKatalogiISlowniki, Me.mnuBazaDanych, Me.mnuAktualizacjaStruktur, Me.mnuSprawdzeniePoprawnosci, Me.ToolStripMenuItem1, Me.ToolStripSeparator6, Me.mnuSzabony, Me.ToolStripSeparator3, Me.mnuBackup, Me.mnuRestore, Me.ToolStripSeparator4, Me.mnuZTofi, Me.mnuMiesieczne, Me.ToolStripSeparator7, Me.mnuDoAnalizy, Me.mnuWartZakPopRok})
        Me.mnuProgram.Name = "mnuProgram"
        Me.mnuProgram.Size = New System.Drawing.Size(65, 20)
        Me.mnuProgram.Text = "Program"
        '
        'mnuOAutorach
        '
        Me.mnuOAutorach.Name = "mnuOAutorach"
        Me.mnuOAutorach.Size = New System.Drawing.Size(462, 22)
        Me.mnuOAutorach.Text = "O autorach"
        '
        'mnuOProgramie
        '
        Me.mnuOProgramie.Name = "mnuOProgramie"
        Me.mnuOProgramie.Size = New System.Drawing.Size(462, 22)
        Me.mnuOProgramie.Text = "O programie"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(459, 6)
        '
        'mnuLokalne
        '
        Me.mnuLokalne.Name = "mnuLokalne"
        Me.mnuLokalne.Size = New System.Drawing.Size(462, 22)
        Me.mnuLokalne.Text = "Parametry lokalne programu"
        '
        'menuParametry
        '
        Me.menuParametry.Name = "menuParametry"
        Me.menuParametry.Size = New System.Drawing.Size(462, 22)
        Me.menuParametry.Text = "Parametry programu"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(459, 6)
        '
        'mnuKatalogiISlowniki
        '
        Me.mnuKatalogiISlowniki.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuKatalogPracownik�w, Me.mnuKatalogAkcjiMarketingowych, Me.mnuKatalogOs�bKontaktowych, Me.menuKatalogBranze, Me.menuKatalogPodbranze, Me.ToolStripSeparator11, Me.mnuKatalogObrot�w, Me.mnuKatalogObrot�wMiesi�cznych, Me.ToolStripSeparator12, Me.S�ownikRodzaj�wKontakt�wToolStripMenuItem, Me.KatalogKontakt�wHandlowychToolStripMenuItem, Me.ToolStripSeparator13, Me.S�ownikCoDalejToolStripMenuItem, Me.KatalogCoDalejToolStripMenuItem, Me.ToolStripSeparator8, Me.KatalogLogAktywnosci})
        Me.mnuKatalogiISlowniki.Name = "mnuKatalogiISlowniki"
        Me.mnuKatalogiISlowniki.Size = New System.Drawing.Size(462, 22)
        Me.mnuKatalogiISlowniki.Text = "Katalogi i s�owniki"
        '
        'mnuKatalogPracownik�w
        '
        Me.mnuKatalogPracownik�w.Name = "mnuKatalogPracownik�w"
        Me.mnuKatalogPracownik�w.Size = New System.Drawing.Size(251, 22)
        Me.mnuKatalogPracownik�w.Text = "Katalog pracownik�w"
        '
        'mnuKatalogAkcjiMarketingowych
        '
        Me.mnuKatalogAkcjiMarketingowych.Name = "mnuKatalogAkcjiMarketingowych"
        Me.mnuKatalogAkcjiMarketingowych.Size = New System.Drawing.Size(251, 22)
        Me.mnuKatalogAkcjiMarketingowych.Text = "Katalog akcji marketingowych"
        '
        'mnuKatalogOs�bKontaktowych
        '
        Me.mnuKatalogOs�bKontaktowych.Name = "mnuKatalogOs�bKontaktowych"
        Me.mnuKatalogOs�bKontaktowych.Size = New System.Drawing.Size(251, 22)
        Me.mnuKatalogOs�bKontaktowych.Text = "Katalog os�b kontaktowych"
        '
        'menuKatalogBranze
        '
        Me.menuKatalogBranze.Name = "menuKatalogBranze"
        Me.menuKatalogBranze.Size = New System.Drawing.Size(251, 22)
        Me.menuKatalogBranze.Text = "Katalog Bran�e"
        '
        'menuKatalogPodbranze
        '
        Me.menuKatalogPodbranze.Name = "menuKatalogPodbranze"
        Me.menuKatalogPodbranze.Size = New System.Drawing.Size(251, 22)
        Me.menuKatalogPodbranze.Text = "Katalog PodBran�e"
        '
        'ToolStripSeparator11
        '
        Me.ToolStripSeparator11.Name = "ToolStripSeparator11"
        Me.ToolStripSeparator11.Size = New System.Drawing.Size(248, 6)
        '
        'mnuKatalogObrot�w
        '
        Me.mnuKatalogObrot�w.Name = "mnuKatalogObrot�w"
        Me.mnuKatalogObrot�w.Size = New System.Drawing.Size(251, 22)
        Me.mnuKatalogObrot�w.Text = "Katalog obrot�w"
        '
        'mnuKatalogObrot�wMiesi�cznych
        '
        Me.mnuKatalogObrot�wMiesi�cznych.Name = "mnuKatalogObrot�wMiesi�cznych"
        Me.mnuKatalogObrot�wMiesi�cznych.Size = New System.Drawing.Size(251, 22)
        Me.mnuKatalogObrot�wMiesi�cznych.Text = "Katalog obrot�w miesi�cznych"
        '
        'ToolStripSeparator12
        '
        Me.ToolStripSeparator12.Name = "ToolStripSeparator12"
        Me.ToolStripSeparator12.Size = New System.Drawing.Size(248, 6)
        '
        'S�ownikRodzaj�wKontakt�wToolStripMenuItem
        '
        Me.S�ownikRodzaj�wKontakt�wToolStripMenuItem.Name = "S�ownikRodzaj�wKontakt�wToolStripMenuItem"
        Me.S�ownikRodzaj�wKontakt�wToolStripMenuItem.Size = New System.Drawing.Size(251, 22)
        Me.S�ownikRodzaj�wKontakt�wToolStripMenuItem.Text = "S�ownik Rodzaj�w Kontakt�w"
        '
        'KatalogKontakt�wHandlowychToolStripMenuItem
        '
        Me.KatalogKontakt�wHandlowychToolStripMenuItem.Name = "KatalogKontakt�wHandlowychToolStripMenuItem"
        Me.KatalogKontakt�wHandlowychToolStripMenuItem.Size = New System.Drawing.Size(251, 22)
        Me.KatalogKontakt�wHandlowychToolStripMenuItem.Text = "Katalog kontakt�w handlowych"
        '
        'ToolStripSeparator13
        '
        Me.ToolStripSeparator13.Name = "ToolStripSeparator13"
        Me.ToolStripSeparator13.Size = New System.Drawing.Size(248, 6)
        '
        'S�ownikCoDalejToolStripMenuItem
        '
        Me.S�ownikCoDalejToolStripMenuItem.Name = "S�ownikCoDalejToolStripMenuItem"
        Me.S�ownikCoDalejToolStripMenuItem.Size = New System.Drawing.Size(251, 22)
        Me.S�ownikCoDalejToolStripMenuItem.Text = "S�ownik Co Dalej-Rodzaj dzia�ania"
        '
        'KatalogCoDalejToolStripMenuItem
        '
        Me.KatalogCoDalejToolStripMenuItem.Name = "KatalogCoDalejToolStripMenuItem"
        Me.KatalogCoDalejToolStripMenuItem.Size = New System.Drawing.Size(251, 22)
        Me.KatalogCoDalejToolStripMenuItem.Text = "Katalog Co Dalej"
        '
        'ToolStripSeparator8
        '
        Me.ToolStripSeparator8.Name = "ToolStripSeparator8"
        Me.ToolStripSeparator8.Size = New System.Drawing.Size(248, 6)
        '
        'KatalogLogAktywnosci
        '
        Me.KatalogLogAktywnosci.Name = "KatalogLogAktywnosci"
        Me.KatalogLogAktywnosci.Size = New System.Drawing.Size(251, 22)
        Me.KatalogLogAktywnosci.Text = "Log aktywno�ci pracownik�w"
        '
        'mnuBazaDanych
        '
        Me.mnuBazaDanych.Name = "mnuBazaDanych"
        Me.mnuBazaDanych.Size = New System.Drawing.Size(462, 22)
        Me.mnuBazaDanych.Text = "Baza danych programu"
        '
        'mnuAktualizacjaStruktur
        '
        Me.mnuAktualizacjaStruktur.Name = "mnuAktualizacjaStruktur"
        Me.mnuAktualizacjaStruktur.Size = New System.Drawing.Size(462, 22)
        Me.mnuAktualizacjaStruktur.Text = "Aktualizacja struktur Bazy Danych"
        '
        'mnuSprawdzeniePoprawnosci
        '
        Me.mnuSprawdzeniePoprawnosci.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SprawdzenieKataloguOs�bKontaktowychToolStripMenuItem, Me.SprawdzenieKataloguObrot�wToolStripMenuItem, Me.SprawdzenieKataloguObrot�wMiesi�cznychToolStripMenuItem, Me.SprawdzenieKataloguKontakt�wHandlowychToolStripMenuItem, Me.menuSprawdzenieKataloguCoDalej, Me.menuZmianaNazwKatKlientow, Me.menuZmianaNazwKatKontaktow})
        Me.mnuSprawdzeniePoprawnosci.Name = "mnuSprawdzeniePoprawnosci"
        Me.mnuSprawdzeniePoprawnosci.Size = New System.Drawing.Size(462, 22)
        Me.mnuSprawdzeniePoprawnosci.Text = "Sprawdzenie poprawno�ci danych w katalogach"
        '
        'SprawdzenieKataloguOs�bKontaktowychToolStripMenuItem
        '
        Me.SprawdzenieKataloguOs�bKontaktowychToolStripMenuItem.Name = "SprawdzenieKataloguOs�bKontaktowychToolStripMenuItem"
        Me.SprawdzenieKataloguOs�bKontaktowychToolStripMenuItem.Size = New System.Drawing.Size(399, 22)
        Me.SprawdzenieKataloguOs�bKontaktowychToolStripMenuItem.Text = "Sprawdzenie Katalogu os�b kontaktowych"
        '
        'SprawdzenieKataloguObrot�wToolStripMenuItem
        '
        Me.SprawdzenieKataloguObrot�wToolStripMenuItem.Name = "SprawdzenieKataloguObrot�wToolStripMenuItem"
        Me.SprawdzenieKataloguObrot�wToolStripMenuItem.Size = New System.Drawing.Size(399, 22)
        Me.SprawdzenieKataloguObrot�wToolStripMenuItem.Text = "Sprawdzenie Katalogu obrot�w"
        '
        'SprawdzenieKataloguObrot�wMiesi�cznychToolStripMenuItem
        '
        Me.SprawdzenieKataloguObrot�wMiesi�cznychToolStripMenuItem.Name = "SprawdzenieKataloguObrot�wMiesi�cznychToolStripMenuItem"
        Me.SprawdzenieKataloguObrot�wMiesi�cznychToolStripMenuItem.Size = New System.Drawing.Size(399, 22)
        Me.SprawdzenieKataloguObrot�wMiesi�cznychToolStripMenuItem.Text = "Sprawdzenie Katalogu obrot�w miesi�cznych"
        '
        'SprawdzenieKataloguKontakt�wHandlowychToolStripMenuItem
        '
        Me.SprawdzenieKataloguKontakt�wHandlowychToolStripMenuItem.Name = "SprawdzenieKataloguKontakt�wHandlowychToolStripMenuItem"
        Me.SprawdzenieKataloguKontakt�wHandlowychToolStripMenuItem.Size = New System.Drawing.Size(399, 22)
        Me.SprawdzenieKataloguKontakt�wHandlowychToolStripMenuItem.Text = "Sprawdzenie Katalogu kontakt�w handlowych"
        '
        'menuSprawdzenieKataloguCoDalej
        '
        Me.menuSprawdzenieKataloguCoDalej.Name = "menuSprawdzenieKataloguCoDalej"
        Me.menuSprawdzenieKataloguCoDalej.Size = New System.Drawing.Size(399, 22)
        Me.menuSprawdzenieKataloguCoDalej.Text = "Sprawdzenie Katalogu CoDalej "
        '
        'menuZmianaNazwKatKlientow
        '
        Me.menuZmianaNazwKatKlientow.Name = "menuZmianaNazwKatKlientow"
        Me.menuZmianaNazwKatKlientow.Size = New System.Drawing.Size(399, 22)
        Me.menuZmianaNazwKatKlientow.Text = "Zmiana nazw w Katalogu Klient�w (Anonimizacja)"
        '
        'menuZmianaNazwKatKontaktow
        '
        Me.menuZmianaNazwKatKontaktow.Name = "menuZmianaNazwKatKontaktow"
        Me.menuZmianaNazwKatKontaktow.Size = New System.Drawing.Size(399, 22)
        Me.menuZmianaNazwKatKontaktow.Text = "Zmiana nazw w Katalogu Os�b Kontaktowych (Anonimizacja)"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuKorektaRZU})
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(462, 22)
        Me.ToolStripMenuItem1.Text = "Korekta b��dnych zapis�w w katalogach"
        '
        'mnuKorektaRZU
        '
        Me.mnuKorektaRZU.Name = "mnuKorektaRZU"
        Me.mnuKorektaRZU.Size = New System.Drawing.Size(482, 22)
        Me.mnuKorektaRZU.Text = "Korekta b��dnych warto�ci Rozszerzonego Zakresu Us�ug w Katalogu Klient�w"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(459, 6)
        '
        'mnuSzabony
        '
        Me.mnuSzabony.Name = "mnuSzabony"
        Me.mnuSzabony.Size = New System.Drawing.Size(462, 22)
        Me.mnuSzabony.Text = "Edycja szablonu filtrowania"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(459, 6)
        '
        'mnuBackup
        '
        Me.mnuBackup.Name = "mnuBackup"
        Me.mnuBackup.Size = New System.Drawing.Size(462, 22)
        Me.mnuBackup.Text = "Kopia archiwalna bazy danych"
        '
        'mnuRestore
        '
        Me.mnuRestore.Name = "mnuRestore"
        Me.mnuRestore.Size = New System.Drawing.Size(462, 22)
        Me.mnuRestore.Text = "Odtwarzanie danych z kopii archiwalnej"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(459, 6)
        '
        'mnuZTofi
        '
        Me.mnuZTofi.Name = "mnuZTofi"
        Me.mnuZTofi.Size = New System.Drawing.Size(462, 22)
        Me.mnuZTofi.Text = "Pobranie danych z programu Tofi"
        '
        'mnuMiesieczne
        '
        Me.mnuMiesieczne.Name = "mnuMiesieczne"
        Me.mnuMiesieczne.Size = New System.Drawing.Size(462, 22)
        Me.mnuMiesieczne.Text = "Przepisanie danych o obrotach bie��cych do danych miesi�cznych"
        '
        'ToolStripSeparator7
        '
        Me.ToolStripSeparator7.Name = "ToolStripSeparator7"
        Me.ToolStripSeparator7.Size = New System.Drawing.Size(459, 6)
        '
        'mnuDoAnalizy
        '
        Me.mnuDoAnalizy.Name = "mnuDoAnalizy"
        Me.mnuDoAnalizy.Size = New System.Drawing.Size(462, 22)
        Me.mnuDoAnalizy.Text = "Pobranie danych o obrotach do analizy zakupu klient�w"
        '
        'mnuWartZakPopRok
        '
        Me.mnuWartZakPopRok.Name = "mnuWartZakPopRok"
        Me.mnuWartZakPopRok.Size = New System.Drawing.Size(462, 22)
        Me.mnuWartZakPopRok.Text = "Zmiana roku kalendarzowego - Inicjowanie nowego roku rozliczeniowego"
        '
        'mnuSystem2
        '
        Me.mnuSystem2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuSysInfo, Me.mnuInternet, Me.ToolStripSeparator1, Me.mnuNotatnik, Me.mnuKalkulator, Me.mnuKalendarz})
        Me.mnuSystem2.Name = "mnuSystem2"
        Me.mnuSystem2.Size = New System.Drawing.Size(57, 20)
        Me.mnuSystem2.Text = "System"
        '
        'mnuSysInfo
        '
        Me.mnuSysInfo.Name = "mnuSysInfo"
        Me.mnuSysInfo.Size = New System.Drawing.Size(199, 22)
        Me.mnuSysInfo.Text = "System Info"
        '
        'mnuInternet
        '
        Me.mnuInternet.Name = "mnuInternet"
        Me.mnuInternet.Size = New System.Drawing.Size(199, 22)
        Me.mnuInternet.Text = "Po��czenie z Internetem"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(196, 6)
        '
        'mnuNotatnik
        '
        Me.mnuNotatnik.Name = "mnuNotatnik"
        Me.mnuNotatnik.Size = New System.Drawing.Size(199, 22)
        Me.mnuNotatnik.Text = "Notatnik"
        '
        'mnuKalkulator
        '
        Me.mnuKalkulator.Name = "mnuKalkulator"
        Me.mnuKalkulator.Size = New System.Drawing.Size(199, 22)
        Me.mnuKalkulator.Text = "Kalkulator"
        '
        'mnuKalendarz
        '
        Me.mnuKalendarz.Name = "mnuKalendarz"
        Me.mnuKalendarz.Size = New System.Drawing.Size(199, 22)
        Me.mnuKalendarz.Text = "Kalendarz"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuZmianaPracownika, Me.mnuZmianaHasla})
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(74, 20)
        Me.ToolStripMenuItem2.Text = "Pracownik"
        '
        'mnuZmianaPracownika
        '
        Me.mnuZmianaPracownika.Name = "mnuZmianaPracownika"
        Me.mnuZmianaPracownika.Size = New System.Drawing.Size(295, 22)
        Me.mnuZmianaPracownika.Text = "Zmiana bie��cego u�ytkownika programu"
        '
        'mnuZmianaHasla
        '
        Me.mnuZmianaHasla.Name = "mnuZmianaHasla"
        Me.mnuZmianaHasla.Size = New System.Drawing.Size(295, 22)
        Me.mnuZmianaHasla.Text = "Zmiana has�a bie��cego u�ytkownika"
        '
        'mnuKoniec
        '
        Me.mnuKoniec.Name = "mnuKoniec"
        Me.mnuKoniec.Size = New System.Drawing.Size(55, 20)
        Me.mnuKoniec.Text = "Koniec"
        '
        'picSoftart
        '
        Me.picSoftart.Location = New System.Drawing.Point(164, 316)
        Me.picSoftart.Name = "picSoftart"
        Me.picSoftart.Size = New System.Drawing.Size(110, 20)
        Me.picSoftart.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picSoftart.TabIndex = 18
        Me.picSoftart.TabStop = False
        '
        'lblAutor
        '
        Me.lblAutor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblAutor.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblAutor.ForeColor = System.Drawing.Color.Navy
        Me.lblAutor.Location = New System.Drawing.Point(12, 315)
        Me.lblAutor.Name = "lblAutor"
        Me.lblAutor.Size = New System.Drawing.Size(263, 22)
        Me.lblAutor.TabIndex = 19
        Me.lblAutor.Text = "Autorem programu jest"
        Me.lblAutor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Zaczynamy
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(481, 354)
        Me.Controls.Add(Me.picSoftart)
        Me.Controls.Add(Me.lblAutor)
        Me.Controls.Add(Me.lblWersja)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.lblTytul)
        Me.Controls.Add(Me.picTlo)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.Name = "Zaczynamy"
        Me.Text = " "
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picTlo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.picSoftart, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Dim Program_CoRobimy As String = ""

    Private Sub frmZaczynamy_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.picSoftart.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        'ustawienia �rodowiska aplikacji
        Dim ThisSystem As OperatingSystem = Environment.OSVersion
        If (ThisSystem.Platform = PlatformID.Win32NT) AndAlso
           (ThisSystem.Version.CompareTo(New Version(5, 1, 0, 0)) >= 0) Then
            Application.EnableVisualStyles()
        End If
        '---------------------------------------------------
        'analiza parametrow wywolania
        ' /H - uruchom tylko HARMONOGRAM KONTAKT�W
        ' puste - menu programu
        Program_CoRobimy = ""
        For Each argument As String In My.Application.CommandLineArgs
            If Mid(argument, 1, 2) = "/H" Then
                'Program_CoRobimy = "H"
                'MenuStrip1.Visible = False
            Else
            End If
        Next
        '---------------------------------------------------
        Dim Proc() As Process = Nothing
        ' sprawdz czy ten proces juz jest uruchomiony
        Dim ModuleName As String = Process.GetCurrentProcess.MainModule.ModuleName
        Dim ProcName As String = System.IO.Path.GetFileNameWithoutExtension(ModuleName)
        Dim versionText As String = Environment.OSVersion.Version.ToString()

        'Proc = Process.GetProcessesByName(ProcName)
        'If Proc.Length > 1 Then
        '    ' There is more than one process with this name.
        '    ' Therefore, this instance should end.
        '    MessageBox.Show("Ten program jest ju� raz uruchomiony na tym komputerze" & vbCrLf & "Musz� si� wy��czy�...", _
        '        "Obs�uga magazynu", _
        '        System.Windows.Forms.MessageBoxButtons.OK, _
        '        MessageBoxIcon.Information, _
        '        MessageBoxDefaultButton.Button1)
        '    End
        'End If
        '---------------------------------------------------
        'Numer wersji
        lblWersja.Text = "Wersja " & Trim(Str(System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMajorPart)) &
                         "." & Trim(Str(System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMinorPart)) &
                         "." & Trim(Str(System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileBuildPart))
        'lblAutor.Text = "Autorem programu jest"
        '---------------------------------------------------
        'sprawdz parametry wywolania
        'Dim CommandParam As String = Command()
        'Dim CommandUser As String = ""
        'Dim SeparPoz As Integer = 0
        ''Jesli Admin - udost�pnij archiwizowanie
        'SeparPoz = InStr(CommandParam, "/U")
        'If SeparPoz > 0 Then
        '    'pobierz uzytkownika
        '    CommandUser = Mid(CommandParam, SeparPoz + 2)
        '    SeparPoz = InStr(CommandParam, " ")
        '    If SeparPoz > 0 Then
        '        CommandUser = Mid(CommandUser, 1, SeparPoz - 1)
        '    End If
        '    'commanduser zawiera ident uzytkownika
        '    If CommandUser = "Admin" Then
        '        Me.mnuBackup.Enabled = True
        '        Me.mnuRestore.Enabled = True
        '    End If
        'End If
        '---------------------------------------------------
        'If GetRegionalOption() = "pl-PL" Or GetRegionalOption() = "sk-SK" Then
        'Else
        '    MessageBox.Show("Opcje regionalne i j�zykowe tego komputera" & vbCrLf & _
        '        "ustawione s� inaczej niz wymaga tego program (" & GetRegionalOption() & ")." & vbCrLf & _
        '        "Prosz� ustawi� opcje na pl-PL i ponownie uruchomi� program...", _
        '        "Uwaga :", _
        '        System.Windows.Forms.MessageBoxButtons.OK, _
        '        MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        '    End
        'End If
        '---------------------------------------------------
        'Me.Visible = False
        'definicje zmiennych
        DefiniujZmienneGlobalne()
        tmrStart.Interval = 5
        tmrStart.Enabled = True
    End Sub


    Private Sub tmrStart_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrStart.Tick
        'wylacz ten zegar
        tmrStart.Enabled = False
        System.Windows.Forms.Application.DoEvents()
        '---------------------------------------------------
        'definiuj parametry programu
        If Len(Dir(oKatParametry & "\" & oPlikParametry)) = 0 Then
            'brak pliku z parametrami
            InicjujParametryLokalne()
            Dim Form As New ParametryLokalne
            Form.ShowDialog()
        Else
            PobierzParametryLokalne()
        End If
        ConnectionStrings()
        If ParL_DataBaseType = "MS ACCESS" Then
            Me.mnuRestore.Enabled = True
        Else
            Me.mnuRestore.Enabled = False
        End If

        'sprawdz wersje danych
        oWersjaBazyDanych = WersjaBazyDanych()
        'sprawdz poprawnosc definicji BD
        Do While oWersjaBazyDanych = 0
            If MessageBox.Show("W parametrach programu b��dnie zdefiniowano baz� danych." & vbCrLf &
                "Czy chcesz edytowa� parametry programu ?",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then

                Dim Form As New ParametryLokalne
                Form.ShowDialog()
                System.Windows.Forms.Application.DoEvents()
                oWersjaBazyDanych = WersjaBazyDanych()
            Else
                System.Windows.Forms.Application.DoEvents()
                Me.Close()
                End
            End If
        Loop

        If oWersjaBazyDanych <> oAktualnaWersjaBazyDanych Then
            If MessageBox.Show("Dost�pna Baza Danych ma wersj� " & (oWersjaBazyDanych / 100).ToString("N2") & vbCrLf &
                "Program wymaga Bazy Danych co najmniej w wersji " & (oAktualnaWersjaBazyDanych / 100).ToString("N2") & vbCrLf &
                "Czy mam wykona� aktualizacj� struktur Bazy Danych ?",
                "Prosz� o decyzje :",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then


                System.Windows.Forms.Application.DoEvents()
                Dim Form As New AktualizacjaStrukturBazyDanych
                Form.ShowDialog()
                'sprawdz wersje bazy danych
                oWersjaBazyDanych = WersjaBazyDanych()
                If oWersjaBazyDanych <> oAktualnaWersjaBazyDanych Then
                    System.Windows.Forms.Application.DoEvents()
                    Me.Close()
                    End
                Else
                End If
            Else
                System.Windows.Forms.Application.DoEvents()
                Me.Close()
                End
            End If
        End If

        '--------------------------
        PobierzParametry()
        '--------------------------
        Select Case Program_CoRobimy
            Case "H"
                'uruchom Harmonogram KOntakt�w
                Dim Form As New KatalogHarmonogramWizyt
                Form.ShowDialog()

                System.Windows.Forms.Application.DoEvents()
                Me.Close()
                End
            Case Else
                '--------------------------
                'wybor operatora programu
                Program_STOP = False
                Dim Form2 As New frmIdentyfikacjaUzytkownika()
                System.Windows.Forms.Application.DoEvents()
                Using Form2
                    Form2.ShowDialog()
                End Using
                Form2 = Nothing

                If Program_STOP Then
                    'koniec pracy programu ......!!!
                    mnuKoniec_Click(sender, e)
                Else
                    Me.Text = "Kartoteka Klient�w firmy PRYZMAT [ " & Program_IdUzytkownika & " " & Program_NazwaUzytkownika & " ]"

                    mnuProgram.Visible = True
                    mnuLokalne.Visible = True
                    menuParametry.Visible = True
                    mnuKatalogiISlowniki.Visible = True
                    mnuBazaDanych.Visible = True
                    mnuAktualizacjaStruktur.Visible = True
                    mnuSprawdzeniePoprawnosci.Visible = True
                    mnuSzabony.Visible = True
                    mnuBackup.Visible = True
                    mnuRestore.Visible = True
                    mnuZTofi.Visible = True
                    mnuMiesieczne.Visible = True
                    mnuDoAnalizy.Visible = True
                    mnuWartZakPopRok.Visible = True

                    ToolStripSeparator2.Visible = True
                    ToolStripSeparator3.Visible = True
                    ToolStripSeparator4.Visible = True
                    ToolStripSeparator5.Visible = True
                    ToolStripSeparator6.Visible = True
                    ToolStripSeparator7.Visible = True

                    If Program_IdUzytkownika = Program_AdministratorLogin Then
                        menuAnalizaWskaznikowa.Visible = True
                        menuAnalizaKlienci80.Visible = True
                        tstExportDanych.Visible = True
                        mnuProgram.Enabled = True
                    ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaAdministrator Then
                        menuAnalizaWskaznikowa.Visible = True
                        menuAnalizaKlienci80.Visible = True
                        tstExportDanych.Visible = True
                        mnuProgram.Enabled = True
                    ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaSzef Then
                        menuAnalizaWskaznikowa.Visible = True
                        menuAnalizaKlienci80.Visible = True
                        tstExportDanych.Visible = True
                        mnuProgram.Enabled = True

                    ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaPracownikUprzywilejowany Then
                        menuAnalizaWskaznikowa.Visible = False
                        menuAnalizaKlienci80.Visible = False
                        tstExportDanych.Visible = False

                        mnuProgram.Visible = True
                        mnuLokalne.Visible = False
                        menuParametry.Visible = False
                        mnuKatalogiISlowniki.Visible = False
                        mnuBazaDanych.Visible = False
                        mnuAktualizacjaStruktur.Visible = False
                        mnuSprawdzeniePoprawnosci.Visible = False
                        mnuSzabony.Visible = False
                        mnuBackup.Visible = True
                        mnuRestore.Visible = False
                        mnuZTofi.Visible = True
                        mnuMiesieczne.Visible = False
                        mnuDoAnalizy.Visible = True
                        mnuWartZakPopRok.Visible = False

                        ToolStripSeparator2.Visible = False
                        ToolStripSeparator3.Visible = False
                        ToolStripSeparator4.Visible = False
                        ToolStripSeparator5.Visible = False
                        ToolStripSeparator6.Visible = False
                        ToolStripSeparator7.Visible = False


                    Else        'pracownik
                        menuAnalizaWskaznikowa.Visible = False
                        menuAnalizaKlienci80.Visible = False
                        tstExportDanych.Visible = False

                        mnuProgram.Enabled = False
                    End If

                End If
        End Select

    End Sub




    '****************************************************
    '** obs�uga menu programu
    '****************************************************

    Private Sub mnuKoniec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuKoniec.Click
        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Application.Exit()
        'Me.Close()
        End
    End Sub

    Private Sub mnuSysInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSysInfo.Click
        'Const RegistryPath As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
        'Dim Key As Microsoft.Win32.RegistryKey
        'Dim Path As String

        'System.Windows.Forms.Application.DoEvents()
        'Key = Registry.LocalMachine.OpenSubKey(RegistryPath, True)
        'Path = CType(Key.GetValue("PATH"), String)
        'If (Dir(Path) <> "") Then
        '    Shell(Path, AppWinStyle.NormalFocus, True)
        'Else    ' Error - File Can Not Be Found...
        '    MessageBox.Show("Nie mog� pobra� informacji o systemie...", "Przykro mi",
        '        System.Windows.Forms.MessageBoxButtons.OK,
        '        MessageBoxIcon.Information,
        '        MessageBoxDefaultButton.Button1)
        'End If
        Try
            Shell("MSINFO32.EXE", AppWinStyle.NormalFocus, True)
        Catch ex As Exception
            MessageBox.Show("B��d podczas pr�by wywo�ania MSINFO (" + ex.ToString() & ")")
        End Try
    End Sub


    Private Sub mnuInternet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuInternet.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New _frmPolaczenieSieciowe
        Form.ShowDialog()
    End Sub

    Private Sub mnuNotatnik_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNotatnik.Click
        'Shell("notepad", AppWinStyle.NormalFocus, False)
        Dim proc As New Process
        System.Windows.Forms.Application.DoEvents()
        proc.StartInfo.FileName = "Notepad.exe"
        proc.StartInfo.Arguments = ""
        proc.Start()
    End Sub

    Private Sub mnuKalkulator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuKalkulator.Click
        System.Windows.Forms.Application.DoEvents()
        Shell("calc", AppWinStyle.NormalFocus, False)
    End Sub

    Private Sub mnuKalendarz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuKalendarz.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New Kalendarz
        Form.Show()
    End Sub

    Private Sub mnuOAutorach_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOAutorach.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New _frmOAutorach
        Form.ShowDialog()
    End Sub

    Private Sub mnuOProgramie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOProgramie.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New _frmOProgramie
        Form.ShowDialog()
    End Sub

    Private Sub mnuLokalne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLokalne.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New ParametryLokalne
        Form.ShowDialog()
        '-----------------------
        ConnectionStrings()
        If ParL_DataBaseType = "MS ACCESS" Then
            Me.mnuRestore.Enabled = True
        Else
            Me.mnuRestore.Enabled = False
        End If

    End Sub

    Private Sub menuParametry_Click(sender As Object, e As EventArgs) Handles menuParametry.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New Parametry
        Form.ShowDialog()
    End Sub



    Private Sub mnuBazaDanych_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuBazaDanych.Click
        System.Windows.Forms.Application.DoEvents()
        oPlik = oPlikKartoteki
        oKomunikat = "Katalog Klient�w"
        If ParL_DataBaseType = "MS ACCESS" Then
            Dim Form As New Katalogi
            Form.ShowDialog()
        Else
            Dim FormSQL As New _frmKatalogi
            'Dim FormSQL As New KatalogiSQL
            FormSQL.ShowDialog()
        End If

        'oPlik = oPlikKartoteki
        'oKomunikat = ""
        'Dim Form As New _frmKatalogi
        'System.Windows.Forms.Application.DoEvents()
        'Using Form
        '    Form.ShowDialog()
        'End Using
        'Form = Nothing
    End Sub

    Private Sub mnuSzabony_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSzabony.Click
        System.Windows.Forms.Application.DoEvents()
        oCoMamRobic = "EDYTUJ"
        oNazwaSchematu = ""
        oSchematFiltrowania = ""
        Dim Form As New SchematFiltrowania("KatalogKlientow")
        Form.ShowDialog()
    End Sub

    Private Sub mnuBackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuBackup.Click
        System.Windows.Forms.Application.DoEvents()
        If ParL_DataBaseType = "MS ACCESS" Then
            Dim Form As New ArchiwizacjaDanych
            Form.ShowDialog()
        Else
            Dim FormSQL As New ArchiwizacjaDanychSQL
            FormSQL.ShowDialog()
        End If
    End Sub

    Private Sub OdmnuRestore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRestore.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New OdtwarzanieDanych
        Form.ShowDialog()
    End Sub


    Private Sub mnuZTofi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuZTofi.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New PobierzObrotyzTOFI
        Form.ShowDialog()
    End Sub

    Private Sub mnuMiesieczne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMiesieczne.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New PobierzDaneMiesieczne
        Form.ShowDialog()
    End Sub


    Private Sub mnuKartotekaKlientow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuKartotekaKlientow.Click
        oCoMamRobic = "EDYTUJ"
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New KatalogKlientow
        Form.Show()
        'Form.ShowDialog()
    End Sub

    Private Sub mnuAnalizaZakupuKlientow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAnalizaZakupuKlientow.Click
        oCoMamRobic = "EDYTUJ"
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New KatalogKlientowAnaliza
        'Form.Show()
        Form.ShowDialog()
    End Sub

    Private Sub mnuDoAnalizy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDoAnalizy.Click
        System.Windows.Forms.Application.DoEvents()
        'Dim Form As New PobierzObrotyzTOFIdoAnalizy
        Dim Form As New PobierzObrotyzTOFIdoAnalizy2
        Form.ShowDialog()
    End Sub

    Private Sub mnuKatalogPracownik�w_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuKatalogPracownik�w.Click
        System.Windows.Forms.Application.DoEvents()
        oPracownik = ""
        oCoMamRobic = "EDYTUJ"
        Dim Form As New KatalogUzytkownikow
        Form.ShowDialog()
    End Sub

    Private Sub mnuKatalogAkcjiMarketingowych_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuKatalogAkcjiMarketingowych.Click
        System.Windows.Forms.Application.DoEvents()
        oCoMamRobic = "EDYTUJ"
        oAkcja = ""
        Dim Form As New KatalogAkcjiOpis
        Form.ShowDialog()
    End Sub

    Private Sub mnuKatalogOs�bKontaktowych_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuKatalogOs�bKontaktowych.Click
        System.Windows.Forms.Application.DoEvents()
        oPracownik = ""
        oIdentKli = ""
        oCoMamRobic = "EDYTUJ"
        Dim Form As New KatalogOsob
        Form.ShowDialog()
    End Sub

    Private Sub mnuKatalogObrot�w_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuKatalogObrot�w.Click
        System.Windows.Forms.Application.DoEvents()
        oPracownik = ""
        oIdentKli = ""
        oCoMamRobic = "EDYTUJ"
        Dim Form As New KatalogObrotow(True)
        Form.ShowDialog()
    End Sub

    Private Sub mnuKatalogObrot�wMiesi�cznych_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuKatalogObrot�wMiesi�cznych.Click
        System.Windows.Forms.Application.DoEvents()
        oPracownik = ""
        oIdentKli = ""
        oCoMamRobic = "EDYTUJ"
        Dim Form As New KatalogObrotowMies(True)
        Form.ShowDialog()
    End Sub

    Private Sub SprawdzenieKataloguOs�bKontaktowychToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SprawdzenieKataloguOs�bKontaktowychToolStripMenuItem.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New SprawdzKatalogOsob
        Form.ShowDialog()
    End Sub

    Private Sub SprawdzenieKataloguObrot�wToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SprawdzenieKataloguObrot�wToolStripMenuItem.Click
        System.Windows.Forms.Application.DoEvents()
        Dim form As New SprawdzKatalogObrotow
        form.ShowDialog()
    End Sub

    Private Sub SprawdzenieKataloguObrot�wMiesi�cznychToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SprawdzenieKataloguObrot�wMiesi�cznychToolStripMenuItem.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New SprawdzKatalogObrotowMies
        Form.ShowDialog()
    End Sub

    Private Sub SprawdzenieKataloguKontakt�wHandlowychToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SprawdzenieKataloguKontakt�wHandlowychToolStripMenuItem.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New SprawdzKatalogKontaktow
        Form.ShowDialog()
    End Sub

    Private Sub menuSprawdzenieKataloguCoDalej_Click(sender As Object, e As EventArgs) Handles menuSprawdzenieKataloguCoDalej.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New SprawdzKatalogCoDalej
        Form.ShowDialog()
    End Sub

    Private Sub mnuAktualizacjaStruktur_Click(sender As Object, e As EventArgs) Handles mnuAktualizacjaStruktur.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New AktualizacjaStrukturBazyDanych
        Form.ShowDialog()
    End Sub

    Private Sub KatalogKontakt�wHandlowychToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KatalogKontakt�wHandlowychToolStripMenuItem.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New KatalogKontaktow("")
        Form.ShowDialog()
    End Sub

    Private Sub S�ownikRodzaj�wKontakt�wToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles S�ownikRodzaj�wKontakt�wToolStripMenuItem.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New KatalogRodzajowKontaktow
        Form.ShowDialog()
    End Sub

    Private Sub menuAnalizaKontakt�wHandlowych_Click(sender As Object, e As EventArgs) Handles menuAnalizaKontakt�wHandlowych.Click
        System.Windows.Forms.Application.DoEvents()
        oCoMamRobic = "EDYTUJ"
        Dim Form As New KatalogKlientowIKontaktow(Nothing, Nothing)
        Form.ShowDialog()
    End Sub

    Private Sub RaportAktywno�ciPrzedstawicieliHandlowychToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RaportAktywno�ciPrzedstawicieliHandlowychToolStripMenuItem.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New frmDrukujRaportAktywnosci
        'Dim Form As New frmDrukujRaportAktywnosci0
        Form.ShowDialog()
    End Sub

    Private Sub mnuDaneDoRaportu_Click(sender As Object, e As EventArgs) Handles mnuDaneDoRaportu.Click
        System.Windows.Forms.Application.DoEvents()
        oCoMamRobic = "EDYTUJ"
        Dim Form As New KatalogDaneDoRaportu
        Form.ShowDialog()
    End Sub

    Private Sub menuHarmonogram_Click(sender As Object, e As EventArgs) Handles menuHarmonogram.Click
        Dim Form As New KatalogHarmonogramWizyt
        Form.Show()
    End Sub


    Private Sub KatalogCoDalejToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KatalogCoDalejToolStripMenuItem.Click
        System.Windows.Forms.Application.DoEvents()
        oCoMamRobic = "EDYTUJ"
        Dim Form As New KatalogCoDalej("")
        Form.ShowDialog()
    End Sub

    Private Sub S�ownikCoDalejToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles S�ownikCoDalejToolStripMenuItem.Click
        System.Windows.Forms.Application.DoEvents()
        oCoMamRobic = "EDYTUJ"
        Dim Form As New KatalogSlownikCoDalej()
        Form.ShowDialog()
    End Sub

    Private Sub KatalogLogAktywnosci_Click(sender As Object, e As EventArgs) Handles KatalogLogAktywnosci.Click
        System.Windows.Forms.Application.DoEvents()
        oCoMamRobic = "EDYTUJ"
        Dim Form As New frmKatalogLogAktywnosci()
        Form.ShowDialog()
    End Sub




    Private Sub mnuWartZakPopRok_Click(sender As Object, e As EventArgs) Handles mnuWartZakPopRok.Click
        Dim Form As New WyliczZakupyPopRok
        Form.ShowDialog()
    End Sub


    Private Sub mnuZmianaPracownika_Click(sender As Object, e As EventArgs) Handles mnuZmianaPracownika.Click
        Dim Form As New frmIdentyfikacjaUzytkownika()
        System.Windows.Forms.Application.DoEvents()
        Using Form
            Form.ShowDialog()
        End Using
        Form = Nothing

        If Program_STOP Then
            'koniec pracy programu ......!!!
            mnuKoniec_Click(sender, e)
        Else
            Me.Text = "Kartoteka Klient�w firmy PRYZMAT [ " & Program_IdUzytkownika & " " & Program_NazwaUzytkownika & " ]"

            mnuProgram.Visible = True
            mnuLokalne.Visible = True
            menuParametry.Visible = True
            mnuKatalogiISlowniki.Visible = True
            mnuBazaDanych.Visible = True
            mnuAktualizacjaStruktur.Visible = True
            mnuSprawdzeniePoprawnosci.Visible = True
            mnuSzabony.Visible = True
            mnuBackup.Visible = True
            mnuRestore.Visible = True
            mnuZTofi.Visible = True
            mnuMiesieczne.Visible = True
            mnuDoAnalizy.Visible = True
            mnuWartZakPopRok.Visible = True

            ToolStripSeparator2.Visible = True
            ToolStripSeparator3.Visible = True
            ToolStripSeparator4.Visible = True
            ToolStripSeparator5.Visible = True
            ToolStripSeparator6.Visible = True
            ToolStripSeparator7.Visible = True

            If Program_IdUzytkownika = Program_AdministratorLogin Then
                menuAnalizaWskaznikowa.Visible = True
                menuAnalizaKlienci80.Visible = True
                tstExportDanych.Visible = True
                mnuProgram.Enabled = True
            ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaAdministrator Then
                menuAnalizaWskaznikowa.Visible = True
                menuAnalizaKlienci80.Visible = True
                tstExportDanych.Visible = True
                mnuProgram.Enabled = True
            ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaSzef Then
                menuAnalizaWskaznikowa.Visible = True
                menuAnalizaKlienci80.Visible = True
                tstExportDanych.Visible = True
                mnuProgram.Enabled = True

            ElseIf OdczytajUprawnieniaUzytkownika(oUprawnieniaUzytkownika) = Program_RolaPracownikUprzywilejowany Then
                menuAnalizaWskaznikowa.Visible = False
                menuAnalizaKlienci80.Visible = False
                tstExportDanych.Visible = False

                mnuProgram.Visible = True
                mnuLokalne.Visible = False
                menuParametry.Visible = False
                mnuKatalogiISlowniki.Visible = False
                mnuBazaDanych.Visible = False
                mnuAktualizacjaStruktur.Visible = False
                mnuSprawdzeniePoprawnosci.Visible = False
                mnuSzabony.Visible = False
                mnuBackup.Visible = True
                mnuRestore.Visible = False
                mnuZTofi.Visible = True
                mnuMiesieczne.Visible = False
                mnuDoAnalizy.Visible = True
                mnuWartZakPopRok.Visible = False

                ToolStripSeparator2.Visible = False
                ToolStripSeparator3.Visible = False
                ToolStripSeparator4.Visible = False
                ToolStripSeparator5.Visible = False
                ToolStripSeparator6.Visible = False
                ToolStripSeparator7.Visible = False

            Else        'pracownik
                menuAnalizaWskaznikowa.Visible = False
                menuAnalizaKlienci80.Visible = False
                tstExportDanych.Visible = False

                mnuProgram.Enabled = False
            End If
        End If
    End Sub

    Private Sub mnuZmianaHasla_Click(sender As Object, e As EventArgs) Handles mnuZmianaHasla.Click
        If Program_IdUzytkownika = Program_AdministratorLogin Then
            MessageBox.Show("Za pomoc� tej funkcji" & vbCrLf & "nie mo�na zmieni� has�a SUPERVISORa...", "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            IdentUzytkownika(Program_IdUzytkownika)
            Dim Form As New frmZmianaHasla(True)
            System.Windows.Forms.Application.DoEvents()
            Using Form
                Form.ShowDialog()
            End Using
            Form = Nothing
            'ZapiszNoweHaslo(u_identUzytkownika,u_kluczUzytkownika))
        End If
    End Sub









    Private Sub menuAnalizaWskaznikowa_Click(sender As Object, e As EventArgs) Handles menuAnalizaWskaznikowa.Click
        Dim Form As New frmEksportDanychAnalizaWskaznikowa
        System.Windows.Forms.Application.DoEvents()
        Using Form
            Form.ShowDialog()
        End Using
        Form = Nothing
    End Sub

    Private Sub menuAnalizaKlienci80_Click(sender As Object, e As EventArgs) Handles menuAnalizaKlienci80.Click
        Dim Form As New frmEksportDanychAnalizaKlienci80
        System.Windows.Forms.Application.DoEvents()
        Using Form
            Form.ShowDialog()
        End Using
        Form = Nothing
    End Sub

    Private Sub menuAnalizaMarketingowa_Click(sender As Object, e As EventArgs) Handles menuAnalizaMarketingowa.Click
        Dim Form As New frmEksportDanychAnalizaMarketingowa
        System.Windows.Forms.Application.DoEvents()
        Using Form
            Form.ShowDialog()
        End Using
        Form = Nothing
    End Sub

    Private Sub menuRankingWskaznikow_Click(sender As Object, e As EventArgs) Handles menuRankingWskaznikow.Click
        Dim Form As New frmEksportDanychRankingWskaznikow
        System.Windows.Forms.Application.DoEvents()
        Using Form
            Form.ShowDialog()
        End Using
        Form = Nothing
    End Sub






    Private Sub mnuKorektaRZU_Click(sender As Object, e As EventArgs) Handles mnuKorektaRZU.Click
        Dim Form As New frmKorektaRZU
        System.Windows.Forms.Application.DoEvents()
        Using Form
            Form.ShowDialog()
        End Using
        Form = Nothing
    End Sub





    '=========================
    ' ZMIANA NAZW W KATALOGACH KLIENTOW I KONTAKTOW
    ' UWAGA - Zmiana realizowana BEZPOWROTNIWE....
    '===================================

    Private Sub MenuZmianaNazwKatKlientow_Click(sender As Object, e As EventArgs) Handles menuZmianaNazwKatKlientow.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New ZmienNazwyKatalogKlientow
        Form.ShowDialog()
    End Sub

    Private Sub MenuZmianaNazwKatKontaktow_Click(sender As Object, e As EventArgs) Handles menuZmianaNazwKatKontaktow.Click
        System.Windows.Forms.Application.DoEvents()
        Dim Form As New ZmienNazwyKatalogOsob
        Form.ShowDialog()
    End Sub

    Private Sub menuKatalogBranze_Click(sender As Object, e As EventArgs) Handles menuKatalogBranze.Click
        System.Windows.Forms.Application.DoEvents()
        oIdBranzy = ""
        oCoMamRobic = "EDYTUJ"
        Dim Form As New KatalogBranze
        Form.ShowDialog()
    End Sub

    Private Sub menuKatalogPodbranze_Click(sender As Object, e As EventArgs) Handles menuKatalogPodbranze.Click
        System.Windows.Forms.Application.DoEvents()
        oIdBranzy = ""
        oIdPodBranzy = ""
        oCoMamRobic = "EDYTUJ"
        Dim Form As New KatalogPodBranze
        Form.ShowDialog()
    End Sub

End Class
