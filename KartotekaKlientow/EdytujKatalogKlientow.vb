Public Class EdytujKatalogKlientow
    Inherits System.Windows.Forms.Form
    'Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "
    '====================================================
    'zmienne do œledzenia JAK ZAMKNIETO FORME
    Private _closeClick As Boolean
    'Private components As System.ComponentModel.Container
    Public Const SC_CLOSE As Integer = 61536
    Friend WithEvents txtWartGran As TextBox
    Friend WithEvents Label153 As Label
    Friend WithEvents lblNryKartPayback As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents lblTelefonKom_Akcje As Label
    Friend WithEvents Label133 As Label
    Friend WithEvents lblTelefon_Akcje As Label
    Friend WithEvents Label135 As Label
    Friend WithEvents lblStanowisko_Akcje As Label
    Friend WithEvents Label148 As Label
    Friend WithEvents lblTelefonKom_Obroty As Label
    Friend WithEvents Label127 As Label
    Friend WithEvents lblTelefon_Obroty As Label
    Friend WithEvents Label129 As Label
    Friend WithEvents lblStanowisko_Obroty As Label
    Friend WithEvents Label131 As Label
    Friend WithEvents lblTelefonKom_ObrotyMies As Label
    Friend WithEvents Label121 As Label
    Friend WithEvents lblTelefon_ObrotyMies As Label
    Friend WithEvents Label123 As Label
    Friend WithEvents lblStanowisko_ObrotyMies As Label
    Friend WithEvents Label125 As Label
    Friend WithEvents lblTelefonKom_Analizy As Label
    Friend WithEvents Label115 As Label
    Friend WithEvents lblTelefon_Analizy As Label
    Friend WithEvents Label117 As Label
    Friend WithEvents lblStanowisko_Analizy As Label
    Friend WithEvents Label119 As Label
    Friend WithEvents lblTelefonKom_UDod As Label
    Friend WithEvents Label105 As Label
    Friend WithEvents lblTelefon_UDod As Label
    Friend WithEvents Label111 As Label
    Friend WithEvents lblStanowisko_UDod As Label
    Friend WithEvents Label113 As Label
    Friend WithEvents lblTelefonKom_Charakterystyka As Label
    Friend WithEvents Label64 As Label
    Friend WithEvents lblTelefon_Charakterystyka As Label
    Friend WithEvents Label74 As Label
    Friend WithEvents lblStanowisko_Charakterystyka As Label
    Friend WithEvents Label84 As Label
    Friend WithEvents lblTelefonKom_Sprzedaz As Label
    Friend WithEvents Label62 As Label
    Friend WithEvents lblTelefon_Sprzedaz As Label
    Friend WithEvents Label41 As Label
    Friend WithEvents lblStanowisko_Sprzedaz As Label
    Friend WithEvents Label39 As Label
    Friend WithEvents lblKlientNieaktywny As Label
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents txtPlanKrotkookresowy As TextBox
    Friend WithEvents txtPlanDlugookresowy As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label22 As Label
    Friend WithEvents cbxSposobWyboruDostawcy As ComboBox
    Friend WithEvents Label38 As Label
    Friend WithEvents Label37 As Label
    Friend WithEvents lblPotencjalDruku As Label
    Friend WithEvents Label36 As Label
    Friend WithEvents Label35 As Label
    Friend WithEvents txtLiczbaWydrukow As TextBox
    Friend WithEvents Label29 As Label
    Friend WithEvents txtLiczbaUrzadzen As TextBox
    Friend WithEvents Label28 As Label
    Friend WithEvents Label20 As Label
    Friend WithEvents chkZainstalowanyMonitor As CheckBox
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents Label23 As Label
    Friend WithEvents chkZaintKorzystaZNajmu As CheckBox
    Friend WithEvents chkZaintRozliczaStrony As CheckBox
    Friend WithEvents chkAktKorzystaZNajmu As CheckBox
    Friend WithEvents chkAktRozliczaStrony As CheckBox
    Friend WithEvents chkAktZakupMatEksp As CheckBox
    Friend WithEvents Label63 As Label
    Friend WithEvents Panel14 As Panel
    Friend WithEvents Label58 As Label
    Friend WithEvents Panel10 As Panel
    Friend WithEvents Label53 As Label
    Friend WithEvents Label47 As Label
    Friend WithEvents Label42 As Label
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Panel5 As Panel
    Friend WithEvents Label40 As Label
    Friend WithEvents Label66 As Label
    Friend WithEvents dtpDataWeryfikKrytSegmentacji As DateTimePicker
    Friend WithEvents cmdWybierzPodbranze As Button
    Friend WithEvents cmdWybierzBranze As Button
    Friend WithEvents txtLiczbaPracZdalnie As TextBox
    Friend WithEvents Label78 As Label
    Friend WithEvents txtLiczbaZatrudnionych As TextBox
    Friend WithEvents Label79 As Label
    Friend WithEvents txtPodbranza As TextBox
    Friend WithEvents Label67 As Label
    Friend WithEvents txtBranza As TextBox
    Friend WithEvents Label75 As Label
    Public Const WM_SYSCOMMAND As Integer = 274

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
    Friend WithEvents cmdWycofajSie As System.Windows.Forms.Button
    Friend WithEvents cmdPrzywroc As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdKontaktPromotor As System.Windows.Forms.Button
    Friend WithEvents cmdWysylaj As System.Windows.Forms.Button
    Friend WithEvents btnPop As System.Windows.Forms.Button
    Friend WithEvents btnNast As System.Windows.Forms.Button
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents DodajKartêPayback As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UsuñKartêPayback As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EdytujKartêPayback As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents MenuKontakty As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents menuSingleLine As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuMultiLine As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmdCoDalej As System.Windows.Forms.Button
    Friend WithEvents MenuCoDalej As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents menuSingleLineCoDalej As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuMultiLineCoDalej As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuRZUKontakty As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents menuRZUSingleLine As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuRZUMultiLine As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuOdswiezCoDalej As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents menuOdswiezKontakty As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents menuOdswiezRZUKontakty As ToolStripMenuItem
    Friend WithEvents MenuOsoby As ContextMenuStrip
    Friend WithEvents MenuOsobySingleL As ToolStripMenuItem
    Friend WithEvents menuOsobyMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents menuOsobyOdswiez As ToolStripMenuItem
    Friend WithEvents menuOsobyEdytuj As ToolStripMenuItem
    Friend WithEvents menuOsobyDodaj As ToolStripMenuItem
    Friend WithEvents menuOsobyUsun As ToolStripMenuItem
    Friend WithEvents AkcjeSpec As TabPage
    Friend WithEvents cmdWszystkoAkcjeSpec As Button
    Friend WithEvents stbFiltryAkcjeSpec As StatusBar
    Friend WithEvents dagAkcjeSpec As DataGridView
    Friend WithEvents stbAkcjeSpec As StatusBar
    Friend WithEvents lblNazwa2_Akcje As Label
    Friend WithEvents Label86 As Label
    Friend WithEvents lblNazwa1_Akcje As Label
    Friend WithEvents lblNazwaHandlowa_Akcje As Label
    Friend WithEvents Label92 As Label
    Friend WithEvents lblTofi_Akcje As Label
    Friend WithEvents lblIdent_Akcje As Label
    Friend WithEvents Panel8 As Panel
    Friend WithEvents Label95 As Label
    Friend WithEvents Label96 As Label
    Friend WithEvents Label97 As Label
    Friend WithEvents Obroty As TabPage
    Friend WithEvents cmdWszystkoObroty As Button
    Friend WithEvents stbFiltryObroty As StatusBar
    Friend WithEvents dagObroty As DataGridView
    Friend WithEvents stbObroty As StatusBar
    Friend WithEvents lblNazwa2_Obroty As Label
    Friend WithEvents Label68 As Label
    Friend WithEvents cmdEdytujObroty As Button
    Friend WithEvents cmdusuñObroty As Button
    Friend WithEvents cmdDodajObroty As Button
    Friend WithEvents lblNazwa1_Obroty As Label
    Friend WithEvents lblNazwaHandlowa_Obroty As Label
    Friend WithEvents Label85 As Label
    Friend WithEvents lblTofi_Obroty As Label
    Friend WithEvents lblIdent_Obroty As Label
    Friend WithEvents Panel7 As Panel
    Friend WithEvents Label89 As Label
    Friend WithEvents Label90 As Label
    Friend WithEvents Label91 As Label
    Friend WithEvents ObrotyMies As TabPage
    Friend WithEvents cmdWszystkoObrotyMies As Button
    Friend WithEvents stbFiltryObrotyMies As StatusBar
    Friend WithEvents dagObrotyMies As DataGridView
    Friend WithEvents stbObrotyMies As StatusBar
    Friend WithEvents lblNazwa2_ObrotyMies As Label
    Friend WithEvents Label69 As Label
    Friend WithEvents cmdEdytujObrotyMies As Button
    Friend WithEvents cmdUsuñObrotyMies As Button
    Friend WithEvents cmdDodajObrotyMies As Button
    Friend WithEvents lblNazwa1_ObrotyMies As Label
    Friend WithEvents lblNazwaHandlowa_ObrotyMies As Label
    Friend WithEvents Label76 As Label
    Friend WithEvents lblTofi_ObrotyMies As Label
    Friend WithEvents lblIdent_ObrotyMies As Label
    Friend WithEvents Panel6 As Panel
    Friend WithEvents Label80 As Label
    Friend WithEvents Label81 As Label
    Friend WithEvents Label82 As Label
    Friend WithEvents AnalizyZakupu As TabPage
    Friend WithEvents cmdWszystkoAnalizyZakupu As Button
    Friend WithEvents stbFiltryAnalizyZakupu As StatusBar
    Friend WithEvents dagAnalizyZakupu As DataGridView
    Friend WithEvents stbAnalizyZakupu As StatusBar
    Friend WithEvents lblNazwa2_Analizy As Label
    Friend WithEvents Label93 As Label
    Friend WithEvents lblNazwa1_Analizy As Label
    Friend WithEvents lblNazwaHandlowa_Analizy As Label
    Friend WithEvents Label99 As Label
    Friend WithEvents lblTOFI_Analizy As Label
    Friend WithEvents lblIdent_Analizy As Label
    Friend WithEvents Panel13 As Panel
    Friend WithEvents Label102 As Label
    Friend WithEvents Label103 As Label
    Friend WithEvents Label104 As Label
    Friend WithEvents cmdWszystkoOsoby As Button
    Friend WithEvents stbFiltryOsoby As StatusBar
    Friend WithEvents dagOsoby As DataGridView
    Friend WithEvents stbOsoby As StatusBar
    Friend WithEvents cmdEdytujOsoby As Button
    Friend WithEvents cmdUsuñOsoby As Button
    Friend WithEvents cmdDodajOsoby As Button
    Friend WithEvents UslugiDod As TabPage
    Friend WithEvents stbRZUKontakty As StatusBar
    Friend WithEvents cmdFi As Button
    Friend WithEvents cmdWszystkoRZUKontakty As Button
    Friend WithEvents stbFiltryRZUKontakty As StatusBar
    Friend WithEvents dtpDataOb As DateTimePicker
    Friend WithEvents Label109 As Label
    Friend WithEvents Label71 As Label
    Friend WithEvents cmdPokazR9 As Button
    Friend WithEvents cmdWybierzR9 As Button
    Friend WithEvents cmdPokazR8 As Button
    Friend WithEvents cmdWybierzR8 As Button
    Friend WithEvents cmdPokazR7 As Button
    Friend WithEvents cmdWybierzR7 As Button
    Friend WithEvents cmdPokazR6 As Button
    Friend WithEvents cmdWybierzR6 As Button
    Friend WithEvents cmdPokazR5 As Button
    Friend WithEvents cmdWybierzR5 As Button
    Friend WithEvents cmdPokazR4 As Button
    Friend WithEvents cmdWybierzR4 As Button
    Friend WithEvents cmdPokazR3 As Button
    Friend WithEvents cmdWybierzR3 As Button
    Friend WithEvents cmdPokazR2 As Button
    Friend WithEvents cmdWybierzR2 As Button
    Friend WithEvents cmdPokazR1 As Button
    Friend WithEvents cmdWybierzR1 As Button
    Friend WithEvents txtRaport09 As TextBox
    Friend WithEvents txtRaport08 As TextBox
    Friend WithEvents txtRaport07 As TextBox
    Friend WithEvents txtRaport06 As TextBox
    Friend WithEvents txtRaport05 As TextBox
    Friend WithEvents txtRaport04 As TextBox
    Friend WithEvents txtRaport03 As TextBox
    Friend WithEvents txtRaport02 As TextBox
    Friend WithEvents txtRaport01 As TextBox
    Friend WithEvents txtzakupyPopRok As TextBox
    Friend WithEvents Label70 As Label
    Friend WithEvents dtpU9 As DateTimePicker
    Friend WithEvents dtpU8 As DateTimePicker
    Friend WithEvents dtpU7 As DateTimePicker
    Friend WithEvents dtpU6 As DateTimePicker
    Friend WithEvents dtpU5 As DateTimePicker
    Friend WithEvents dtpU4 As DateTimePicker
    Friend WithEvents dtpU3 As DateTimePicker
    Friend WithEvents dtpU2 As DateTimePicker
    Friend WithEvents dtpU1 As DateTimePicker
    Friend WithEvents Label61 As Label
    Friend WithEvents chkU9 As CheckBox
    Friend WithEvents chkU8 As CheckBox
    Friend WithEvents chkU7 As CheckBox
    Friend WithEvents chkU6 As CheckBox
    Friend WithEvents chkU5 As CheckBox
    Friend WithEvents chkU4 As CheckBox
    Friend WithEvents chkU3 As CheckBox
    Friend WithEvents chkU2 As CheckBox
    Friend WithEvents chkU1 As CheckBox
    Friend WithEvents Label25 As Label
    Friend WithEvents Label24 As Label
    Friend WithEvents lblNazwa2_UDod As Label
    Friend WithEvents lblNazwa1_UDod As Label
    Friend WithEvents Label57 As Label
    Friend WithEvents lblTofi_UDod As Label
    Friend WithEvents lblNazwaHandlowa_UDod As Label
    Friend WithEvents lblIdent_UDod As Label
    Friend WithEvents Label72 As Label
    Friend WithEvents Label73 As Label
    Friend WithEvents Label94 As Label
    Friend WithEvents Label98 As Label
    Friend WithEvents Charakterystyka As TabPage
    Friend WithEvents cmdPoka¿Plik As Button
    Friend WithEvents cmdWybierzPlik As Button
    Friend WithEvents lblOfertaPlik As Label
    Friend WithEvents Label59 As Label
    Friend WithEvents txtTZlozenia As TextBox
    Friend WithEvents txtUwagiCharakterystyka As TextBox
    Friend WithEvents txtWykazUrzadzen As TextBox
    Friend WithEvents txtPKontakt As TextBox
    Friend WithEvents Label45 As Label
    Friend WithEvents lblNazwa1_Charakterystyka As Label
    Friend WithEvents Label55 As Label
    Friend WithEvents Label54 As Label
    Friend WithEvents Label52 As Label
    Friend WithEvents chbRodzL As CheckBox
    Friend WithEvents chbRodzA As CheckBox
    Friend WithEvents chbRodzT As CheckBox
    Friend WithEvents cmdPKontakt As Button
    Friend WithEvents Label34 As Label
    Friend WithEvents Label27 As Label
    Friend WithEvents Label21 As Label
    Friend WithEvents lblNazwaHandlowa_Charakterystyka As Label
    Friend WithEvents Label19 As Label
    Friend WithEvents lblTofi_Charakterystyka As Label
    Friend WithEvents lblNazwa2_Charakterystyka As Label
    Friend WithEvents lblIdent_Charakterystyka As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label17 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents Label60 As Label
    Friend WithEvents Opiekun As TabPage
    Friend WithEvents txtNryKartPayBack As TextBox
    Friend WithEvents txtSprzedazNKonT As TextBox
    Friend WithEvents txtSprzedazOKonT As TextBox
    Friend WithEvents txtSprzedazOpiekun As TextBox
    Friend WithEvents txtOpisPotencjalu As TextBox
    Friend WithEvents Label43 As Label
    Friend WithEvents txtSprzedazUwagi As TextBox
    Friend WithEvents Label44 As Label
    Friend WithEvents SplitContainer6 As SplitContainer
    Friend WithEvents Panel9 As Panel
    Friend WithEvents dagCoDalej As DataGridView
    Friend WithEvents Label46 As Label
    Friend WithEvents Label56 As Label
    Friend WithEvents dagKontakty As DataGridView
    Friend WithEvents Label11 As Label
    Friend WithEvents cmdHarmonogramWizyt As Button
    Friend WithEvents cmdEdytujNumeryKartPayBack As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents chkKartyPayback As CheckBox
    Friend WithEvents Label88 As Label
    Friend WithEvents Label48 As Label
    Friend WithEvents Label77 As Label
    Friend WithEvents Label87 As Label
    Friend WithEvents Panel11 As Panel
    Friend WithEvents cbxSprzedazNKontR As ComboBox
    Friend WithEvents cbxSprzedazOKontR As ComboBox
    Friend WithEvents Panel12 As Panel
    Friend WithEvents lblOstTransak As Label
    Friend WithEvents Label83 As Label
    Friend WithEvents cbxSprzedazNKonD As ComboBox
    Friend WithEvents lblosoba_Sprzedaz As Label
    Friend WithEvents Label65 As Label
    Friend WithEvents lbladres_Sprzedaz As Label
    Friend WithEvents cmdSprzedazOpiekun As Button
    Friend WithEvents cbxSprzedazOKonD As ComboBox
    Friend WithEvents cmdSprzedazNKonT As Button
    Friend WithEvents cmdSprzedazOKonT As Button
    Friend WithEvents Label49 As Label
    Friend WithEvents Label50 As Label
    Friend WithEvents Label51 As Label
    Friend WithEvents lblNazwa_Sprzedaz As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents lblTofi_Sprzedaz As Label
    Friend WithEvents lblIdent_Sprzedaz As Label
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Label30 As Label
    Friend WithEvents Label31 As Label
    Friend WithEvents Label32 As Label
    Friend WithEvents Identyfikacja As TabPage
    Friend WithEvents cmdNowyNrKarty As Button
    Friend WithEvents txtIDDomu As TextBox
    Friend WithEvents txtRejon As TextBox
    Friend WithEvents txtNrDomu As TextBox
    Friend WithEvents txtMiejscowosc As TextBox
    Friend WithEvents txtNazwa1 As TextBox
    Friend WithEvents txtKtoWpisal As TextBox
    Friend WithEvents txtUlica As TextBox
    Friend WithEvents txtTofi As TextBox
    Friend WithEvents txteMail As TextBox
    Friend WithEvents txtFax As TextBox
    Friend WithEvents txtTelefon As TextBox
    Friend WithEvents txtNIP As TextBox
    Friend WithEvents txtKodP As TextBox
    Friend WithEvents txtIdent As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents cmdNrKarty As Button
    Friend WithEvents Label26 As Label
    Friend WithEvents cmdKtoWpisal As Button
    Friend WithEvents Label33 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label16 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents menuAnalizyZakupu As ContextMenuStrip
    Friend WithEvents menuAnalizyZakupuSingleL As ToolStripMenuItem
    Friend WithEvents menuAnalizyZakupuMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator6 As ToolStripSeparator
    Friend WithEvents menuAnalizyZakupuOdswiez As ToolStripMenuItem
    Friend WithEvents menuObrotyMies As ContextMenuStrip
    Friend WithEvents menuObrotyMiesEdytuj As ToolStripMenuItem
    Friend WithEvents menuObrotyMiesDodaj As ToolStripMenuItem
    Friend WithEvents menuObrotyMiesUsun As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator7 As ToolStripSeparator
    Friend WithEvents menuObrotyMiesSingleL As ToolStripMenuItem
    Friend WithEvents menuObrotyMiesMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator8 As ToolStripSeparator
    Friend WithEvents menuObrotyMiesOdswiez As ToolStripMenuItem
    Friend WithEvents menuObroty As ContextMenuStrip
    Friend WithEvents menuObrotyEdytuj As ToolStripMenuItem
    Friend WithEvents menuObrotyDodaj As ToolStripMenuItem
    Friend WithEvents menuObrotyUsun As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator9 As ToolStripSeparator
    Friend WithEvents menuObrotySingleL As ToolStripMenuItem
    Friend WithEvents menuObrotyMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator10 As ToolStripSeparator
    Friend WithEvents menuObrotyOdswiez As ToolStripMenuItem
    Friend WithEvents menuAkcjeSpec As ContextMenuStrip
    Friend WithEvents menuAkcjeSpecSingleL As ToolStripMenuItem
    Friend WithEvents menuAkcjeSpecMultiL As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator11 As ToolStripSeparator
    Friend WithEvents menuAkcjeSpecOdswiez As ToolStripMenuItem
    Friend WithEvents PanelAkcjeSpec As Panel
    Friend WithEvents PanelUslugiDod As Panel
    Friend WithEvents PanelSprzedaz As Panel
    Friend WithEvents dagRZUKontakty As DataGridView
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EdytujKatalogKlientow))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.btnNast = New System.Windows.Forms.Button()
        Me.btnPop = New System.Windows.Forms.Button()
        Me.cmdWysylaj = New System.Windows.Forms.Button()
        Me.MenuCoDalej = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuSingleLineCoDalej = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuMultiLineCoDalej = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuOdswiezCoDalej = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuKontakty = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuSingleLine = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuMultiLine = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuOdswiezKontakty = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuRZUKontakty = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuRZUSingleLine = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuRZUMultiLine = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuOdswiezRZUKontakty = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.DodajKartêPayback = New System.Windows.Forms.ToolStripMenuItem()
        Me.UsuñKartêPayback = New System.Windows.Forms.ToolStripMenuItem()
        Me.EdytujKartêPayback = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmdKontaktPromotor = New System.Windows.Forms.Button()
        Me.cmdWycofajSie = New System.Windows.Forms.Button()
        Me.cmdPrzywroc = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdWybierzR1 = New System.Windows.Forms.Button()
        Me.cmdPokazR1 = New System.Windows.Forms.Button()
        Me.cmdWybierzR2 = New System.Windows.Forms.Button()
        Me.cmdPokazR2 = New System.Windows.Forms.Button()
        Me.cmdWybierzR3 = New System.Windows.Forms.Button()
        Me.cmdPokazR3 = New System.Windows.Forms.Button()
        Me.cmdWybierzR4 = New System.Windows.Forms.Button()
        Me.cmdPokazR4 = New System.Windows.Forms.Button()
        Me.cmdWybierzR5 = New System.Windows.Forms.Button()
        Me.cmdPokazR5 = New System.Windows.Forms.Button()
        Me.cmdWybierzR6 = New System.Windows.Forms.Button()
        Me.cmdPokazR6 = New System.Windows.Forms.Button()
        Me.cmdWybierzR7 = New System.Windows.Forms.Button()
        Me.cmdPokazR7 = New System.Windows.Forms.Button()
        Me.cmdWybierzR8 = New System.Windows.Forms.Button()
        Me.cmdPokazR8 = New System.Windows.Forms.Button()
        Me.cmdWybierzR9 = New System.Windows.Forms.Button()
        Me.cmdPokazR9 = New System.Windows.Forms.Button()
        Me.cmdWybierzPlik = New System.Windows.Forms.Button()
        Me.cmdPoka¿Plik = New System.Windows.Forms.Button()
        Me.cmdCoDalej = New System.Windows.Forms.Button()
        Me.MenuOsoby = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuOsobyEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuOsobyDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuOsobyUsun = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.MenuOsobySingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuOsobyMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuOsobyOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.AkcjeSpec = New System.Windows.Forms.TabPage()
        Me.PanelAkcjeSpec = New System.Windows.Forms.Panel()
        Me.lblTelefonKom_Akcje = New System.Windows.Forms.Label()
        Me.Label133 = New System.Windows.Forms.Label()
        Me.lblTelefon_Akcje = New System.Windows.Forms.Label()
        Me.Label135 = New System.Windows.Forms.Label()
        Me.lblStanowisko_Akcje = New System.Windows.Forms.Label()
        Me.Label148 = New System.Windows.Forms.Label()
        Me.cmdWszystkoAkcjeSpec = New System.Windows.Forms.Button()
        Me.lblTofi_Akcje = New System.Windows.Forms.Label()
        Me.dagAkcjeSpec = New System.Windows.Forms.DataGridView()
        Me.menuAkcjeSpec = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuAkcjeSpecSingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuAkcjeSpecMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuAkcjeSpecOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.stbFiltryAkcjeSpec = New System.Windows.Forms.StatusBar()
        Me.stbAkcjeSpec = New System.Windows.Forms.StatusBar()
        Me.lblNazwa2_Akcje = New System.Windows.Forms.Label()
        Me.Label86 = New System.Windows.Forms.Label()
        Me.lblNazwa1_Akcje = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa_Akcje = New System.Windows.Forms.Label()
        Me.Label92 = New System.Windows.Forms.Label()
        Me.lblIdent_Akcje = New System.Windows.Forms.Label()
        Me.Label95 = New System.Windows.Forms.Label()
        Me.Label96 = New System.Windows.Forms.Label()
        Me.Label97 = New System.Windows.Forms.Label()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.Obroty = New System.Windows.Forms.TabPage()
        Me.lblTelefonKom_Obroty = New System.Windows.Forms.Label()
        Me.Label127 = New System.Windows.Forms.Label()
        Me.lblTelefon_Obroty = New System.Windows.Forms.Label()
        Me.Label129 = New System.Windows.Forms.Label()
        Me.lblStanowisko_Obroty = New System.Windows.Forms.Label()
        Me.Label131 = New System.Windows.Forms.Label()
        Me.cmdWszystkoObroty = New System.Windows.Forms.Button()
        Me.stbFiltryObroty = New System.Windows.Forms.StatusBar()
        Me.dagObroty = New System.Windows.Forms.DataGridView()
        Me.menuObroty = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuObrotyEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuObrotyDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuObrotyUsun = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuObrotySingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuObrotyMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator10 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuObrotyOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.stbObroty = New System.Windows.Forms.StatusBar()
        Me.lblNazwa2_Obroty = New System.Windows.Forms.Label()
        Me.Label68 = New System.Windows.Forms.Label()
        Me.cmdEdytujObroty = New System.Windows.Forms.Button()
        Me.cmdusuñObroty = New System.Windows.Forms.Button()
        Me.cmdDodajObroty = New System.Windows.Forms.Button()
        Me.lblNazwa1_Obroty = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa_Obroty = New System.Windows.Forms.Label()
        Me.Label85 = New System.Windows.Forms.Label()
        Me.lblTofi_Obroty = New System.Windows.Forms.Label()
        Me.lblIdent_Obroty = New System.Windows.Forms.Label()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.Label89 = New System.Windows.Forms.Label()
        Me.Label90 = New System.Windows.Forms.Label()
        Me.Label91 = New System.Windows.Forms.Label()
        Me.ObrotyMies = New System.Windows.Forms.TabPage()
        Me.lblTelefonKom_ObrotyMies = New System.Windows.Forms.Label()
        Me.Label121 = New System.Windows.Forms.Label()
        Me.lblTelefon_ObrotyMies = New System.Windows.Forms.Label()
        Me.Label123 = New System.Windows.Forms.Label()
        Me.lblStanowisko_ObrotyMies = New System.Windows.Forms.Label()
        Me.Label125 = New System.Windows.Forms.Label()
        Me.cmdWszystkoObrotyMies = New System.Windows.Forms.Button()
        Me.stbFiltryObrotyMies = New System.Windows.Forms.StatusBar()
        Me.dagObrotyMies = New System.Windows.Forms.DataGridView()
        Me.menuObrotyMies = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuObrotyMiesEdytuj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuObrotyMiesDodaj = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuObrotyMiesUsun = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuObrotyMiesSingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuObrotyMiesMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator8 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuObrotyMiesOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.stbObrotyMies = New System.Windows.Forms.StatusBar()
        Me.lblNazwa2_ObrotyMies = New System.Windows.Forms.Label()
        Me.Label69 = New System.Windows.Forms.Label()
        Me.cmdEdytujObrotyMies = New System.Windows.Forms.Button()
        Me.cmdUsuñObrotyMies = New System.Windows.Forms.Button()
        Me.cmdDodajObrotyMies = New System.Windows.Forms.Button()
        Me.lblNazwa1_ObrotyMies = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa_ObrotyMies = New System.Windows.Forms.Label()
        Me.Label76 = New System.Windows.Forms.Label()
        Me.lblTofi_ObrotyMies = New System.Windows.Forms.Label()
        Me.lblIdent_ObrotyMies = New System.Windows.Forms.Label()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.Label80 = New System.Windows.Forms.Label()
        Me.Label81 = New System.Windows.Forms.Label()
        Me.Label82 = New System.Windows.Forms.Label()
        Me.AnalizyZakupu = New System.Windows.Forms.TabPage()
        Me.lblTelefonKom_Analizy = New System.Windows.Forms.Label()
        Me.Label115 = New System.Windows.Forms.Label()
        Me.lblTelefon_Analizy = New System.Windows.Forms.Label()
        Me.Label117 = New System.Windows.Forms.Label()
        Me.lblStanowisko_Analizy = New System.Windows.Forms.Label()
        Me.Label119 = New System.Windows.Forms.Label()
        Me.cmdWszystkoAnalizyZakupu = New System.Windows.Forms.Button()
        Me.stbFiltryAnalizyZakupu = New System.Windows.Forms.StatusBar()
        Me.dagAnalizyZakupu = New System.Windows.Forms.DataGridView()
        Me.menuAnalizyZakupu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.menuAnalizyZakupuSingleL = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuAnalizyZakupuMultiL = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuAnalizyZakupuOdswiez = New System.Windows.Forms.ToolStripMenuItem()
        Me.stbAnalizyZakupu = New System.Windows.Forms.StatusBar()
        Me.lblNazwa2_Analizy = New System.Windows.Forms.Label()
        Me.Label93 = New System.Windows.Forms.Label()
        Me.lblNazwa1_Analizy = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa_Analizy = New System.Windows.Forms.Label()
        Me.Label99 = New System.Windows.Forms.Label()
        Me.lblTOFI_Analizy = New System.Windows.Forms.Label()
        Me.lblIdent_Analizy = New System.Windows.Forms.Label()
        Me.Panel13 = New System.Windows.Forms.Panel()
        Me.Label102 = New System.Windows.Forms.Label()
        Me.Label103 = New System.Windows.Forms.Label()
        Me.Label104 = New System.Windows.Forms.Label()
        Me.cmdWszystkoOsoby = New System.Windows.Forms.Button()
        Me.stbFiltryOsoby = New System.Windows.Forms.StatusBar()
        Me.dagOsoby = New System.Windows.Forms.DataGridView()
        Me.stbOsoby = New System.Windows.Forms.StatusBar()
        Me.cmdEdytujOsoby = New System.Windows.Forms.Button()
        Me.cmdUsuñOsoby = New System.Windows.Forms.Button()
        Me.cmdDodajOsoby = New System.Windows.Forms.Button()
        Me.UslugiDod = New System.Windows.Forms.TabPage()
        Me.PanelUslugiDod = New System.Windows.Forms.Panel()
        Me.lblTelefonKom_UDod = New System.Windows.Forms.Label()
        Me.Label105 = New System.Windows.Forms.Label()
        Me.lblTelefon_UDod = New System.Windows.Forms.Label()
        Me.Label111 = New System.Windows.Forms.Label()
        Me.lblStanowisko_UDod = New System.Windows.Forms.Label()
        Me.Label113 = New System.Windows.Forms.Label()
        Me.txtWartGran = New System.Windows.Forms.TextBox()
        Me.Label153 = New System.Windows.Forms.Label()
        Me.dagRZUKontakty = New System.Windows.Forms.DataGridView()
        Me.stbRZUKontakty = New System.Windows.Forms.StatusBar()
        Me.cmdFi = New System.Windows.Forms.Button()
        Me.cmdWszystkoRZUKontakty = New System.Windows.Forms.Button()
        Me.stbFiltryRZUKontakty = New System.Windows.Forms.StatusBar()
        Me.dtpDataOb = New System.Windows.Forms.DateTimePicker()
        Me.Label109 = New System.Windows.Forms.Label()
        Me.Label71 = New System.Windows.Forms.Label()
        Me.txtRaport09 = New System.Windows.Forms.TextBox()
        Me.txtRaport08 = New System.Windows.Forms.TextBox()
        Me.txtRaport07 = New System.Windows.Forms.TextBox()
        Me.txtRaport06 = New System.Windows.Forms.TextBox()
        Me.txtRaport05 = New System.Windows.Forms.TextBox()
        Me.txtRaport04 = New System.Windows.Forms.TextBox()
        Me.txtRaport03 = New System.Windows.Forms.TextBox()
        Me.txtRaport02 = New System.Windows.Forms.TextBox()
        Me.txtRaport01 = New System.Windows.Forms.TextBox()
        Me.txtzakupyPopRok = New System.Windows.Forms.TextBox()
        Me.Label70 = New System.Windows.Forms.Label()
        Me.dtpU9 = New System.Windows.Forms.DateTimePicker()
        Me.dtpU8 = New System.Windows.Forms.DateTimePicker()
        Me.dtpU6 = New System.Windows.Forms.DateTimePicker()
        Me.dtpU7 = New System.Windows.Forms.DateTimePicker()
        Me.dtpU5 = New System.Windows.Forms.DateTimePicker()
        Me.dtpU4 = New System.Windows.Forms.DateTimePicker()
        Me.dtpU3 = New System.Windows.Forms.DateTimePicker()
        Me.dtpU2 = New System.Windows.Forms.DateTimePicker()
        Me.dtpU1 = New System.Windows.Forms.DateTimePicker()
        Me.Label61 = New System.Windows.Forms.Label()
        Me.chkU9 = New System.Windows.Forms.CheckBox()
        Me.chkU8 = New System.Windows.Forms.CheckBox()
        Me.chkU7 = New System.Windows.Forms.CheckBox()
        Me.chkU6 = New System.Windows.Forms.CheckBox()
        Me.chkU5 = New System.Windows.Forms.CheckBox()
        Me.chkU4 = New System.Windows.Forms.CheckBox()
        Me.chkU3 = New System.Windows.Forms.CheckBox()
        Me.chkU2 = New System.Windows.Forms.CheckBox()
        Me.chkU1 = New System.Windows.Forms.CheckBox()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.lblNazwa2_UDod = New System.Windows.Forms.Label()
        Me.lblNazwa1_UDod = New System.Windows.Forms.Label()
        Me.Label57 = New System.Windows.Forms.Label()
        Me.lblTofi_UDod = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa_UDod = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblIdent_UDod = New System.Windows.Forms.Label()
        Me.Label72 = New System.Windows.Forms.Label()
        Me.Label73 = New System.Windows.Forms.Label()
        Me.Label94 = New System.Windows.Forms.Label()
        Me.Label98 = New System.Windows.Forms.Label()
        Me.Charakterystyka = New System.Windows.Forms.TabPage()
        Me.cmdWybierzPodbranze = New System.Windows.Forms.Button()
        Me.cmdWybierzBranze = New System.Windows.Forms.Button()
        Me.txtLiczbaPracZdalnie = New System.Windows.Forms.TextBox()
        Me.Label78 = New System.Windows.Forms.Label()
        Me.txtLiczbaZatrudnionych = New System.Windows.Forms.TextBox()
        Me.Label79 = New System.Windows.Forms.Label()
        Me.txtPodbranza = New System.Windows.Forms.TextBox()
        Me.Label67 = New System.Windows.Forms.Label()
        Me.txtBranza = New System.Windows.Forms.TextBox()
        Me.Label75 = New System.Windows.Forms.Label()
        Me.chbRodzT = New System.Windows.Forms.CheckBox()
        Me.chbRodzA = New System.Windows.Forms.CheckBox()
        Me.chbRodzL = New System.Windows.Forms.CheckBox()
        Me.dtpDataWeryfikKrytSegmentacji = New System.Windows.Forms.DateTimePicker()
        Me.Label66 = New System.Windows.Forms.Label()
        Me.chkZaintKorzystaZNajmu = New System.Windows.Forms.CheckBox()
        Me.chkZaintRozliczaStrony = New System.Windows.Forms.CheckBox()
        Me.chkAktKorzystaZNajmu = New System.Windows.Forms.CheckBox()
        Me.chkAktRozliczaStrony = New System.Windows.Forms.CheckBox()
        Me.chkAktZakupMatEksp = New System.Windows.Forms.CheckBox()
        Me.Label63 = New System.Windows.Forms.Label()
        Me.Panel14 = New System.Windows.Forms.Panel()
        Me.Label58 = New System.Windows.Forms.Label()
        Me.Panel10 = New System.Windows.Forms.Panel()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.cbxSposobWyboruDostawcy = New System.Windows.Forms.ComboBox()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.lblPotencjalDruku = New System.Windows.Forms.Label()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.txtLiczbaWydrukow = New System.Windows.Forms.TextBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.txtLiczbaUrzadzen = New System.Windows.Forms.TextBox()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.chkZainstalowanyMonitor = New System.Windows.Forms.CheckBox()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.txtUwagiCharakterystyka = New System.Windows.Forms.TextBox()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.txtOpisPotencjalu = New System.Windows.Forms.TextBox()
        Me.lblTelefonKom_Charakterystyka = New System.Windows.Forms.Label()
        Me.Label64 = New System.Windows.Forms.Label()
        Me.lblTelefon_Charakterystyka = New System.Windows.Forms.Label()
        Me.Label74 = New System.Windows.Forms.Label()
        Me.lblStanowisko_Charakterystyka = New System.Windows.Forms.Label()
        Me.Label84 = New System.Windows.Forms.Label()
        Me.txtWykazUrzadzen = New System.Windows.Forms.TextBox()
        Me.lblOfertaPlik = New System.Windows.Forms.Label()
        Me.Label59 = New System.Windows.Forms.Label()
        Me.txtTZlozenia = New System.Windows.Forms.TextBox()
        Me.Label45 = New System.Windows.Forms.Label()
        Me.lblNazwa1_Charakterystyka = New System.Windows.Forms.Label()
        Me.Label55 = New System.Windows.Forms.Label()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.Label52 = New System.Windows.Forms.Label()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.lblNazwaHandlowa_Charakterystyka = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.lblTofi_Charakterystyka = New System.Windows.Forms.Label()
        Me.lblNazwa2_Charakterystyka = New System.Windows.Forms.Label()
        Me.lblIdent_Charakterystyka = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label60 = New System.Windows.Forms.Label()
        Me.cmdPKontakt = New System.Windows.Forms.Button()
        Me.txtPKontakt = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Opiekun = New System.Windows.Forms.TabPage()
        Me.PanelSprzedaz = New System.Windows.Forms.Panel()
        Me.txtPlanKrotkookresowy = New System.Windows.Forms.TextBox()
        Me.txtPlanDlugookresowy = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.txtSprzedazUwagi = New System.Windows.Forms.TextBox()
        Me.lblTelefonKom_Sprzedaz = New System.Windows.Forms.Label()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.Label62 = New System.Windows.Forms.Label()
        Me.lblTelefon_Sprzedaz = New System.Windows.Forms.Label()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.lblStanowisko_Sprzedaz = New System.Windows.Forms.Label()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.cmdSprzedazOKonT = New System.Windows.Forms.Button()
        Me.cmdSprzedazOpiekun = New System.Windows.Forms.Button()
        Me.cmdSprzedazNKonT = New System.Windows.Forms.Button()
        Me.txtSprzedazNKonT = New System.Windows.Forms.TextBox()
        Me.txtSprzedazOKonT = New System.Windows.Forms.TextBox()
        Me.txtSprzedazOpiekun = New System.Windows.Forms.TextBox()
        Me.SplitContainer6 = New System.Windows.Forms.SplitContainer()
        Me.Panel9 = New System.Windows.Forms.Panel()
        Me.dagCoDalej = New System.Windows.Forms.DataGridView()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.dagKontakty = New System.Windows.Forms.DataGridView()
        Me.Label56 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cmdHarmonogramWizyt = New System.Windows.Forms.Button()
        Me.Label48 = New System.Windows.Forms.Label()
        Me.Label77 = New System.Windows.Forms.Label()
        Me.Label87 = New System.Windows.Forms.Label()
        Me.Panel11 = New System.Windows.Forms.Panel()
        Me.cbxSprzedazNKontR = New System.Windows.Forms.ComboBox()
        Me.cbxSprzedazOKontR = New System.Windows.Forms.ComboBox()
        Me.Panel12 = New System.Windows.Forms.Panel()
        Me.lblOstTransak = New System.Windows.Forms.Label()
        Me.Label83 = New System.Windows.Forms.Label()
        Me.cbxSprzedazNKonD = New System.Windows.Forms.ComboBox()
        Me.lblosoba_Sprzedaz = New System.Windows.Forms.Label()
        Me.Label65 = New System.Windows.Forms.Label()
        Me.lbladres_Sprzedaz = New System.Windows.Forms.Label()
        Me.cbxSprzedazOKonD = New System.Windows.Forms.ComboBox()
        Me.Label49 = New System.Windows.Forms.Label()
        Me.Label50 = New System.Windows.Forms.Label()
        Me.Label51 = New System.Windows.Forms.Label()
        Me.lblNazwa_Sprzedaz = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.lblTofi_Sprzedaz = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.lblIdent_Sprzedaz = New System.Windows.Forms.Label()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.txtNryKartPayBack = New System.Windows.Forms.TextBox()
        Me.cmdEdytujNumeryKartPayBack = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.chkKartyPayback = New System.Windows.Forms.CheckBox()
        Me.Label88 = New System.Windows.Forms.Label()
        Me.Identyfikacja = New System.Windows.Forms.TabPage()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblNryKartPayback = New System.Windows.Forms.Label()
        Me.cmdNowyNrKarty = New System.Windows.Forms.Button()
        Me.txtIDDomu = New System.Windows.Forms.TextBox()
        Me.txtRejon = New System.Windows.Forms.TextBox()
        Me.txtNrDomu = New System.Windows.Forms.TextBox()
        Me.txtMiejscowosc = New System.Windows.Forms.TextBox()
        Me.txtNazwa1 = New System.Windows.Forms.TextBox()
        Me.txtKtoWpisal = New System.Windows.Forms.TextBox()
        Me.txtUlica = New System.Windows.Forms.TextBox()
        Me.txtTofi = New System.Windows.Forms.TextBox()
        Me.txteMail = New System.Windows.Forms.TextBox()
        Me.txtFax = New System.Windows.Forms.TextBox()
        Me.txtTelefon = New System.Windows.Forms.TextBox()
        Me.txtNIP = New System.Windows.Forms.TextBox()
        Me.txtKodP = New System.Windows.Forms.TextBox()
        Me.txtIdent = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmdNrKarty = New System.Windows.Forms.Button()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.cmdKtoWpisal = New System.Windows.Forms.Button()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.lblKlientNieaktywny = New System.Windows.Forms.Label()
        Me.MenuCoDalej.SuspendLayout()
        Me.MenuKontakty.SuspendLayout()
        Me.MenuRZUKontakty.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuOsoby.SuspendLayout()
        Me.AkcjeSpec.SuspendLayout()
        Me.PanelAkcjeSpec.SuspendLayout()
        CType(Me.dagAkcjeSpec, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuAkcjeSpec.SuspendLayout()
        Me.Obroty.SuspendLayout()
        CType(Me.dagObroty, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuObroty.SuspendLayout()
        Me.ObrotyMies.SuspendLayout()
        CType(Me.dagObrotyMies, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuObrotyMies.SuspendLayout()
        Me.AnalizyZakupu.SuspendLayout()
        CType(Me.dagAnalizyZakupu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menuAnalizyZakupu.SuspendLayout()
        CType(Me.dagOsoby, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UslugiDod.SuspendLayout()
        Me.PanelUslugiDod.SuspendLayout()
        CType(Me.dagRZUKontakty, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Charakterystyka.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Opiekun.SuspendLayout()
        Me.PanelSprzedaz.SuspendLayout()
        CType(Me.SplitContainer6, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer6.Panel1.SuspendLayout()
        Me.SplitContainer6.Panel2.SuspendLayout()
        Me.SplitContainer6.SuspendLayout()
        Me.Panel9.SuspendLayout()
        CType(Me.dagCoDalej, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dagKontakty, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Identyfikacja.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnNast
        '
        Me.btnNast.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNast.ForeColor = System.Drawing.Color.Black
        Me.btnNast.Image = CType(resources.GetObject("btnNast.Image"), System.Drawing.Image)
        Me.btnNast.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnNast.Location = New System.Drawing.Point(617, 683)
        Me.btnNast.Name = "btnNast"
        Me.btnNast.Size = New System.Drawing.Size(32, 23)
        Me.btnNast.TabIndex = 524
        Me.btnNast.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnPop
        '
        Me.btnPop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPop.ForeColor = System.Drawing.Color.Black
        Me.btnPop.Image = CType(resources.GetObject("btnPop.Image"), System.Drawing.Image)
        Me.btnPop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPop.Location = New System.Drawing.Point(655, 682)
        Me.btnPop.Name = "btnPop"
        Me.btnPop.Size = New System.Drawing.Size(32, 23)
        Me.btnPop.TabIndex = 525
        Me.btnPop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdWysylaj
        '
        Me.cmdWysylaj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWysylaj.ForeColor = System.Drawing.Color.Black
        Me.cmdWysylaj.Image = CType(resources.GetObject("cmdWysylaj.Image"), System.Drawing.Image)
        Me.cmdWysylaj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWysylaj.Location = New System.Drawing.Point(693, 682)
        Me.cmdWysylaj.Name = "cmdWysylaj"
        Me.cmdWysylaj.Size = New System.Drawing.Size(80, 23)
        Me.cmdWysylaj.TabIndex = 526
        Me.cmdWysylaj.Text = "Wyœlij"
        Me.cmdWysylaj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'MenuCoDalej
        '
        Me.MenuCoDalej.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuSingleLineCoDalej, Me.menuMultiLineCoDalej, Me.ToolStripSeparator1, Me.menuOdswiezCoDalej})
        Me.MenuCoDalej.Name = "ContextMenuStrip1"
        Me.MenuCoDalej.Size = New System.Drawing.Size(187, 76)
        '
        'menuSingleLineCoDalej
        '
        Me.menuSingleLineCoDalej.Name = "menuSingleLineCoDalej"
        Me.menuSingleLineCoDalej.Size = New System.Drawing.Size(186, 22)
        Me.menuSingleLineCoDalej.Text = "Poka¿ w jednej linii"
        '
        'menuMultiLineCoDalej
        '
        Me.menuMultiLineCoDalej.Name = "menuMultiLineCoDalej"
        Me.menuMultiLineCoDalej.Size = New System.Drawing.Size(186, 22)
        Me.menuMultiLineCoDalej.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(183, 6)
        '
        'menuOdswiezCoDalej
        '
        Me.menuOdswiezCoDalej.Name = "menuOdswiezCoDalej"
        Me.menuOdswiezCoDalej.Size = New System.Drawing.Size(186, 22)
        Me.menuOdswiezCoDalej.Text = "Od¿wie¿ tabelê"
        '
        'MenuKontakty
        '
        Me.MenuKontakty.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuSingleLine, Me.menuMultiLine, Me.ToolStripSeparator2, Me.menuOdswiezKontakty})
        Me.MenuKontakty.Name = "ContextMenuStrip1"
        Me.MenuKontakty.Size = New System.Drawing.Size(187, 76)
        '
        'menuSingleLine
        '
        Me.menuSingleLine.Name = "menuSingleLine"
        Me.menuSingleLine.Size = New System.Drawing.Size(186, 22)
        Me.menuSingleLine.Text = "Poka¿ w jednej linii"
        '
        'menuMultiLine
        '
        Me.menuMultiLine.Name = "menuMultiLine"
        Me.menuMultiLine.Size = New System.Drawing.Size(186, 22)
        Me.menuMultiLine.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(183, 6)
        '
        'menuOdswiezKontakty
        '
        Me.menuOdswiezKontakty.Name = "menuOdswiezKontakty"
        Me.menuOdswiezKontakty.Size = New System.Drawing.Size(186, 22)
        Me.menuOdswiezKontakty.Text = "Odœwie¿ tabelê"
        '
        'MenuRZUKontakty
        '
        Me.MenuRZUKontakty.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuRZUSingleLine, Me.menuRZUMultiLine, Me.ToolStripSeparator3, Me.menuOdswiezRZUKontakty})
        Me.MenuRZUKontakty.Name = "ContextMenuStrip1"
        Me.MenuRZUKontakty.Size = New System.Drawing.Size(187, 76)
        '
        'menuRZUSingleLine
        '
        Me.menuRZUSingleLine.Name = "menuRZUSingleLine"
        Me.menuRZUSingleLine.Size = New System.Drawing.Size(186, 22)
        Me.menuRZUSingleLine.Text = "Poka¿ w jednej linii"
        '
        'menuRZUMultiLine
        '
        Me.menuRZUMultiLine.Name = "menuRZUMultiLine"
        Me.menuRZUMultiLine.Size = New System.Drawing.Size(186, 22)
        Me.menuRZUMultiLine.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(183, 6)
        '
        'menuOdswiezRZUKontakty
        '
        Me.menuOdswiezRZUKontakty.Name = "menuOdswiezRZUKontakty"
        Me.menuOdswiezRZUKontakty.Size = New System.Drawing.Size(186, 22)
        Me.menuOdswiezRZUKontakty.Text = "Odœwie¿ tabelê"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DodajKartêPayback, Me.UsuñKartêPayback, Me.EdytujKartêPayback})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(184, 70)
        '
        'DodajKartêPayback
        '
        Me.DodajKartêPayback.Name = "DodajKartêPayback"
        Me.DodajKartêPayback.Size = New System.Drawing.Size(183, 22)
        Me.DodajKartêPayback.Text = "Dodaj kartê Payback"
        '
        'UsuñKartêPayback
        '
        Me.UsuñKartêPayback.Name = "UsuñKartêPayback"
        Me.UsuñKartêPayback.Size = New System.Drawing.Size(183, 22)
        Me.UsuñKartêPayback.Text = "Usuñ kartê Payback"
        '
        'EdytujKartêPayback
        '
        Me.EdytujKartêPayback.Name = "EdytujKartêPayback"
        Me.EdytujKartêPayback.Size = New System.Drawing.Size(183, 22)
        Me.EdytujKartêPayback.Text = "Edytuj kartê Payback"
        '
        'cmdKontaktPromotor
        '
        Me.cmdKontaktPromotor.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdKontaktPromotor.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdKontaktPromotor.Location = New System.Drawing.Point(102, 682)
        Me.cmdKontaktPromotor.Name = "cmdKontaktPromotor"
        Me.cmdKontaktPromotor.Size = New System.Drawing.Size(120, 23)
        Me.cmdKontaktPromotor.TabIndex = 52
        Me.cmdKontaktPromotor.Text = "Kontakty handlowe"
        '
        'cmdWycofajSie
        '
        Me.cmdWycofajSie.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWycofajSie.Image = CType(resources.GetObject("cmdWycofajSie.Image"), System.Drawing.Image)
        Me.cmdWycofajSie.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWycofajSie.Location = New System.Drawing.Point(880, 682)
        Me.cmdWycofajSie.Name = "cmdWycofajSie"
        Me.cmdWycofajSie.Size = New System.Drawing.Size(104, 23)
        Me.cmdWycofajSie.TabIndex = 528
        Me.cmdWycofajSie.Text = "Wycofaj siê"
        Me.cmdWycofajSie.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPrzywroc
        '
        Me.cmdPrzywroc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPrzywroc.Image = CType(resources.GetObject("cmdPrzywroc.Image"), System.Drawing.Image)
        Me.cmdPrzywroc.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPrzywroc.Location = New System.Drawing.Point(779, 682)
        Me.cmdPrzywroc.Name = "cmdPrzywroc"
        Me.cmdPrzywroc.Size = New System.Drawing.Size(95, 23)
        Me.cmdPrzywroc.TabIndex = 527
        Me.cmdPrzywroc.Text = "Przywróæ"
        Me.cmdPrzywroc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(990, 682)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 23)
        Me.cmdPowrot.TabIndex = 529
        Me.cmdPowrot.Text = "Zapisz"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 682)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(88, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 16
        Me.PictureBox2.TabStop = False
        '
        'cmdWybierzR1
        '
        Me.cmdWybierzR1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzR1.Image = CType(resources.GetObject("cmdWybierzR1.Image"), System.Drawing.Image)
        Me.cmdWybierzR1.Location = New System.Drawing.Point(970, 142)
        Me.cmdWybierzR1.Name = "cmdWybierzR1"
        Me.cmdWybierzR1.Size = New System.Drawing.Size(32, 22)
        Me.cmdWybierzR1.TabIndex = 196
        Me.ToolTip1.SetToolTip(Me.cmdWybierzR1, "Wybierz plik z ofert¹ z dysku")
        '
        'cmdPokazR1
        '
        Me.cmdPokazR1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokazR1.Image = CType(resources.GetObject("cmdPokazR1.Image"), System.Drawing.Image)
        Me.cmdPokazR1.Location = New System.Drawing.Point(1008, 142)
        Me.cmdPokazR1.Name = "cmdPokazR1"
        Me.cmdPokazR1.Size = New System.Drawing.Size(32, 22)
        Me.cmdPokazR1.TabIndex = 197
        Me.ToolTip1.SetToolTip(Me.cmdPokazR1, "Poka¿ plik z ofert¹")
        '
        'cmdWybierzR2
        '
        Me.cmdWybierzR2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzR2.Image = CType(resources.GetObject("cmdWybierzR2.Image"), System.Drawing.Image)
        Me.cmdWybierzR2.Location = New System.Drawing.Point(970, 166)
        Me.cmdWybierzR2.Name = "cmdWybierzR2"
        Me.cmdWybierzR2.Size = New System.Drawing.Size(32, 22)
        Me.cmdWybierzR2.TabIndex = 198
        Me.ToolTip1.SetToolTip(Me.cmdWybierzR2, "Wybierz plik z ofert¹ z dysku")
        '
        'cmdPokazR2
        '
        Me.cmdPokazR2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokazR2.Image = CType(resources.GetObject("cmdPokazR2.Image"), System.Drawing.Image)
        Me.cmdPokazR2.Location = New System.Drawing.Point(1008, 166)
        Me.cmdPokazR2.Name = "cmdPokazR2"
        Me.cmdPokazR2.Size = New System.Drawing.Size(32, 22)
        Me.cmdPokazR2.TabIndex = 199
        Me.ToolTip1.SetToolTip(Me.cmdPokazR2, "Poka¿ plik z ofert¹")
        '
        'cmdWybierzR3
        '
        Me.cmdWybierzR3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzR3.Image = CType(resources.GetObject("cmdWybierzR3.Image"), System.Drawing.Image)
        Me.cmdWybierzR3.Location = New System.Drawing.Point(970, 188)
        Me.cmdWybierzR3.Name = "cmdWybierzR3"
        Me.cmdWybierzR3.Size = New System.Drawing.Size(32, 22)
        Me.cmdWybierzR3.TabIndex = 200
        Me.ToolTip1.SetToolTip(Me.cmdWybierzR3, "Wybierz plik z ofert¹ z dysku")
        '
        'cmdPokazR3
        '
        Me.cmdPokazR3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokazR3.Image = CType(resources.GetObject("cmdPokazR3.Image"), System.Drawing.Image)
        Me.cmdPokazR3.Location = New System.Drawing.Point(1008, 188)
        Me.cmdPokazR3.Name = "cmdPokazR3"
        Me.cmdPokazR3.Size = New System.Drawing.Size(32, 22)
        Me.cmdPokazR3.TabIndex = 201
        Me.ToolTip1.SetToolTip(Me.cmdPokazR3, "Poka¿ plik z ofert¹")
        '
        'cmdWybierzR4
        '
        Me.cmdWybierzR4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzR4.Image = CType(resources.GetObject("cmdWybierzR4.Image"), System.Drawing.Image)
        Me.cmdWybierzR4.Location = New System.Drawing.Point(970, 210)
        Me.cmdWybierzR4.Name = "cmdWybierzR4"
        Me.cmdWybierzR4.Size = New System.Drawing.Size(32, 22)
        Me.cmdWybierzR4.TabIndex = 202
        Me.ToolTip1.SetToolTip(Me.cmdWybierzR4, "Wybierz plik z ofert¹ z dysku")
        '
        'cmdPokazR4
        '
        Me.cmdPokazR4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokazR4.Image = CType(resources.GetObject("cmdPokazR4.Image"), System.Drawing.Image)
        Me.cmdPokazR4.Location = New System.Drawing.Point(1008, 210)
        Me.cmdPokazR4.Name = "cmdPokazR4"
        Me.cmdPokazR4.Size = New System.Drawing.Size(32, 22)
        Me.cmdPokazR4.TabIndex = 203
        Me.ToolTip1.SetToolTip(Me.cmdPokazR4, "Poka¿ plik z ofert¹")
        '
        'cmdWybierzR5
        '
        Me.cmdWybierzR5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzR5.Image = CType(resources.GetObject("cmdWybierzR5.Image"), System.Drawing.Image)
        Me.cmdWybierzR5.Location = New System.Drawing.Point(970, 232)
        Me.cmdWybierzR5.Name = "cmdWybierzR5"
        Me.cmdWybierzR5.Size = New System.Drawing.Size(32, 22)
        Me.cmdWybierzR5.TabIndex = 204
        Me.ToolTip1.SetToolTip(Me.cmdWybierzR5, "Wybierz plik z ofert¹ z dysku")
        '
        'cmdPokazR5
        '
        Me.cmdPokazR5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokazR5.Image = CType(resources.GetObject("cmdPokazR5.Image"), System.Drawing.Image)
        Me.cmdPokazR5.Location = New System.Drawing.Point(1008, 232)
        Me.cmdPokazR5.Name = "cmdPokazR5"
        Me.cmdPokazR5.Size = New System.Drawing.Size(32, 22)
        Me.cmdPokazR5.TabIndex = 205
        Me.ToolTip1.SetToolTip(Me.cmdPokazR5, "Poka¿ plik z ofert¹")
        '
        'cmdWybierzR6
        '
        Me.cmdWybierzR6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzR6.Image = CType(resources.GetObject("cmdWybierzR6.Image"), System.Drawing.Image)
        Me.cmdWybierzR6.Location = New System.Drawing.Point(970, 254)
        Me.cmdWybierzR6.Name = "cmdWybierzR6"
        Me.cmdWybierzR6.Size = New System.Drawing.Size(32, 22)
        Me.cmdWybierzR6.TabIndex = 206
        Me.ToolTip1.SetToolTip(Me.cmdWybierzR6, "Wybierz plik z ofert¹ z dysku")
        '
        'cmdPokazR6
        '
        Me.cmdPokazR6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokazR6.Image = CType(resources.GetObject("cmdPokazR6.Image"), System.Drawing.Image)
        Me.cmdPokazR6.Location = New System.Drawing.Point(1008, 254)
        Me.cmdPokazR6.Name = "cmdPokazR6"
        Me.cmdPokazR6.Size = New System.Drawing.Size(32, 22)
        Me.cmdPokazR6.TabIndex = 207
        Me.ToolTip1.SetToolTip(Me.cmdPokazR6, "Poka¿ plik z ofert¹")
        '
        'cmdWybierzR7
        '
        Me.cmdWybierzR7.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzR7.Image = CType(resources.GetObject("cmdWybierzR7.Image"), System.Drawing.Image)
        Me.cmdWybierzR7.Location = New System.Drawing.Point(970, 276)
        Me.cmdWybierzR7.Name = "cmdWybierzR7"
        Me.cmdWybierzR7.Size = New System.Drawing.Size(32, 22)
        Me.cmdWybierzR7.TabIndex = 208
        Me.ToolTip1.SetToolTip(Me.cmdWybierzR7, "Wybierz plik z ofert¹ z dysku")
        '
        'cmdPokazR7
        '
        Me.cmdPokazR7.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokazR7.Image = CType(resources.GetObject("cmdPokazR7.Image"), System.Drawing.Image)
        Me.cmdPokazR7.Location = New System.Drawing.Point(1008, 276)
        Me.cmdPokazR7.Name = "cmdPokazR7"
        Me.cmdPokazR7.Size = New System.Drawing.Size(32, 22)
        Me.cmdPokazR7.TabIndex = 209
        Me.ToolTip1.SetToolTip(Me.cmdPokazR7, "Poka¿ plik z ofert¹")
        '
        'cmdWybierzR8
        '
        Me.cmdWybierzR8.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzR8.Image = CType(resources.GetObject("cmdWybierzR8.Image"), System.Drawing.Image)
        Me.cmdWybierzR8.Location = New System.Drawing.Point(970, 298)
        Me.cmdWybierzR8.Name = "cmdWybierzR8"
        Me.cmdWybierzR8.Size = New System.Drawing.Size(32, 22)
        Me.cmdWybierzR8.TabIndex = 210
        Me.ToolTip1.SetToolTip(Me.cmdWybierzR8, "Wybierz plik z ofert¹ z dysku")
        Me.cmdWybierzR8.Visible = False
        '
        'cmdPokazR8
        '
        Me.cmdPokazR8.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokazR8.Image = CType(resources.GetObject("cmdPokazR8.Image"), System.Drawing.Image)
        Me.cmdPokazR8.Location = New System.Drawing.Point(1008, 298)
        Me.cmdPokazR8.Name = "cmdPokazR8"
        Me.cmdPokazR8.Size = New System.Drawing.Size(32, 22)
        Me.cmdPokazR8.TabIndex = 211
        Me.ToolTip1.SetToolTip(Me.cmdPokazR8, "Poka¿ plik z ofert¹")
        Me.cmdPokazR8.Visible = False
        '
        'cmdWybierzR9
        '
        Me.cmdWybierzR9.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzR9.Image = CType(resources.GetObject("cmdWybierzR9.Image"), System.Drawing.Image)
        Me.cmdWybierzR9.Location = New System.Drawing.Point(970, 318)
        Me.cmdWybierzR9.Name = "cmdWybierzR9"
        Me.cmdWybierzR9.Size = New System.Drawing.Size(32, 22)
        Me.cmdWybierzR9.TabIndex = 212
        Me.ToolTip1.SetToolTip(Me.cmdWybierzR9, "Wybierz plik z ofert¹ z dysku")
        Me.cmdWybierzR9.Visible = False
        '
        'cmdPokazR9
        '
        Me.cmdPokazR9.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPokazR9.Image = CType(resources.GetObject("cmdPokazR9.Image"), System.Drawing.Image)
        Me.cmdPokazR9.Location = New System.Drawing.Point(1008, 318)
        Me.cmdPokazR9.Name = "cmdPokazR9"
        Me.cmdPokazR9.Size = New System.Drawing.Size(32, 22)
        Me.cmdPokazR9.TabIndex = 213
        Me.ToolTip1.SetToolTip(Me.cmdPokazR9, "Poka¿ plik z ofert¹")
        Me.cmdPokazR9.Visible = False
        '
        'cmdWybierzPlik
        '
        Me.cmdWybierzPlik.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzPlik.Image = CType(resources.GetObject("cmdWybierzPlik.Image"), System.Drawing.Image)
        Me.cmdWybierzPlik.Location = New System.Drawing.Point(550, 303)
        Me.cmdWybierzPlik.Name = "cmdWybierzPlik"
        Me.cmdWybierzPlik.Size = New System.Drawing.Size(32, 22)
        Me.cmdWybierzPlik.TabIndex = 117
        Me.ToolTip1.SetToolTip(Me.cmdWybierzPlik, "Wybierz plik z ofert¹ z dysku")
        '
        'cmdPoka¿Plik
        '
        Me.cmdPoka¿Plik.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPoka¿Plik.Image = CType(resources.GetObject("cmdPoka¿Plik.Image"), System.Drawing.Image)
        Me.cmdPoka¿Plik.Location = New System.Drawing.Point(588, 303)
        Me.cmdPoka¿Plik.Name = "cmdPoka¿Plik"
        Me.cmdPoka¿Plik.Size = New System.Drawing.Size(32, 22)
        Me.cmdPoka¿Plik.TabIndex = 118
        Me.ToolTip1.SetToolTip(Me.cmdPoka¿Plik, "Poka¿ plik z ofert¹")
        '
        'cmdCoDalej
        '
        Me.cmdCoDalej.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdCoDalej.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCoDalej.Location = New System.Drawing.Point(229, 682)
        Me.cmdCoDalej.Name = "cmdCoDalej"
        Me.cmdCoDalej.Size = New System.Drawing.Size(85, 23)
        Me.cmdCoDalej.TabIndex = 530
        Me.cmdCoDalej.Text = "Co Dalej"
        '
        'MenuOsoby
        '
        Me.MenuOsoby.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuOsobyEdytuj, Me.menuOsobyDodaj, Me.menuOsobyUsun, Me.ToolStripSeparator5, Me.MenuOsobySingleL, Me.menuOsobyMultiL, Me.ToolStripSeparator4, Me.menuOsobyOdswiez})
        Me.MenuOsoby.Name = "ContextMenuStrip1"
        Me.MenuOsoby.Size = New System.Drawing.Size(187, 148)
        '
        'menuOsobyEdytuj
        '
        Me.menuOsobyEdytuj.Name = "menuOsobyEdytuj"
        Me.menuOsobyEdytuj.Size = New System.Drawing.Size(186, 22)
        Me.menuOsobyEdytuj.Text = "Edytuj"
        '
        'menuOsobyDodaj
        '
        Me.menuOsobyDodaj.Name = "menuOsobyDodaj"
        Me.menuOsobyDodaj.Size = New System.Drawing.Size(186, 22)
        Me.menuOsobyDodaj.Text = "Dodaj"
        '
        'menuOsobyUsun
        '
        Me.menuOsobyUsun.Name = "menuOsobyUsun"
        Me.menuOsobyUsun.Size = New System.Drawing.Size(186, 22)
        Me.menuOsobyUsun.Text = "Usuñ"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(183, 6)
        '
        'MenuOsobySingleL
        '
        Me.MenuOsobySingleL.Name = "MenuOsobySingleL"
        Me.MenuOsobySingleL.Size = New System.Drawing.Size(186, 22)
        Me.MenuOsobySingleL.Text = "Poka¿ w jednej linii"
        '
        'menuOsobyMultiL
        '
        Me.menuOsobyMultiL.Name = "menuOsobyMultiL"
        Me.menuOsobyMultiL.Size = New System.Drawing.Size(186, 22)
        Me.menuOsobyMultiL.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(183, 6)
        '
        'menuOsobyOdswiez
        '
        Me.menuOsobyOdswiez.Name = "menuOsobyOdswiez"
        Me.menuOsobyOdswiez.Size = New System.Drawing.Size(186, 22)
        Me.menuOsobyOdswiez.Text = "Odœwie¿ tabelê"
        '
        'AkcjeSpec
        '
        Me.AkcjeSpec.BackColor = System.Drawing.SystemColors.Control
        Me.AkcjeSpec.Controls.Add(Me.PanelAkcjeSpec)
        Me.AkcjeSpec.Location = New System.Drawing.Point(4, 22)
        Me.AkcjeSpec.Name = "AkcjeSpec"
        Me.AkcjeSpec.Size = New System.Drawing.Size(1054, 645)
        Me.AkcjeSpec.TabIndex = 10
        Me.AkcjeSpec.Text = "Akcje marketingowe"
        '
        'PanelAkcjeSpec
        '
        Me.PanelAkcjeSpec.Controls.Add(Me.lblTelefonKom_Akcje)
        Me.PanelAkcjeSpec.Controls.Add(Me.Label133)
        Me.PanelAkcjeSpec.Controls.Add(Me.lblTelefon_Akcje)
        Me.PanelAkcjeSpec.Controls.Add(Me.Label135)
        Me.PanelAkcjeSpec.Controls.Add(Me.lblStanowisko_Akcje)
        Me.PanelAkcjeSpec.Controls.Add(Me.Label148)
        Me.PanelAkcjeSpec.Controls.Add(Me.cmdWszystkoAkcjeSpec)
        Me.PanelAkcjeSpec.Controls.Add(Me.lblTofi_Akcje)
        Me.PanelAkcjeSpec.Controls.Add(Me.dagAkcjeSpec)
        Me.PanelAkcjeSpec.Controls.Add(Me.stbFiltryAkcjeSpec)
        Me.PanelAkcjeSpec.Controls.Add(Me.stbAkcjeSpec)
        Me.PanelAkcjeSpec.Controls.Add(Me.lblNazwa2_Akcje)
        Me.PanelAkcjeSpec.Controls.Add(Me.Label86)
        Me.PanelAkcjeSpec.Controls.Add(Me.lblNazwa1_Akcje)
        Me.PanelAkcjeSpec.Controls.Add(Me.lblNazwaHandlowa_Akcje)
        Me.PanelAkcjeSpec.Controls.Add(Me.Label92)
        Me.PanelAkcjeSpec.Controls.Add(Me.lblIdent_Akcje)
        Me.PanelAkcjeSpec.Controls.Add(Me.Label95)
        Me.PanelAkcjeSpec.Controls.Add(Me.Label96)
        Me.PanelAkcjeSpec.Controls.Add(Me.Label97)
        Me.PanelAkcjeSpec.Controls.Add(Me.Panel8)
        Me.PanelAkcjeSpec.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelAkcjeSpec.Location = New System.Drawing.Point(0, 0)
        Me.PanelAkcjeSpec.Name = "PanelAkcjeSpec"
        Me.PanelAkcjeSpec.Size = New System.Drawing.Size(1054, 645)
        Me.PanelAkcjeSpec.TabIndex = 212
        '
        'lblTelefonKom_Akcje
        '
        Me.lblTelefonKom_Akcje.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefonKom_Akcje.Location = New System.Drawing.Point(940, 56)
        Me.lblTelefonKom_Akcje.Name = "lblTelefonKom_Akcje"
        Me.lblTelefonKom_Akcje.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefonKom_Akcje.TabIndex = 248
        '
        'Label133
        '
        Me.Label133.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label133.ForeColor = System.Drawing.Color.Navy
        Me.Label133.Location = New System.Drawing.Point(891, 56)
        Me.Label133.Name = "Label133"
        Me.Label133.Size = New System.Drawing.Size(112, 16)
        Me.Label133.TabIndex = 249
        Me.Label133.Text = "Tel. kom. . . . . . . . . . . . . . . "
        '
        'lblTelefon_Akcje
        '
        Me.lblTelefon_Akcje.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefon_Akcje.Location = New System.Drawing.Point(673, 56)
        Me.lblTelefon_Akcje.Name = "lblTelefon_Akcje"
        Me.lblTelefon_Akcje.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefon_Akcje.TabIndex = 246
        '
        'Label135
        '
        Me.Label135.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label135.ForeColor = System.Drawing.Color.Navy
        Me.Label135.Location = New System.Drawing.Point(624, 56)
        Me.Label135.Name = "Label135"
        Me.Label135.Size = New System.Drawing.Size(112, 16)
        Me.Label135.TabIndex = 247
        Me.Label135.Text = "Telefon . . . . . . . . . . . . . . "
        '
        'lblStanowisko_Akcje
        '
        Me.lblStanowisko_Akcje.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStanowisko_Akcje.Location = New System.Drawing.Point(404, 56)
        Me.lblStanowisko_Akcje.Name = "lblStanowisko_Akcje"
        Me.lblStanowisko_Akcje.Size = New System.Drawing.Size(200, 16)
        Me.lblStanowisko_Akcje.TabIndex = 244
        '
        'Label148
        '
        Me.Label148.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label148.ForeColor = System.Drawing.Color.Navy
        Me.Label148.Location = New System.Drawing.Point(333, 56)
        Me.Label148.Name = "Label148"
        Me.Label148.Size = New System.Drawing.Size(112, 16)
        Me.Label148.TabIndex = 245
        Me.Label148.Text = "Stanowisko . . . . . . . . . . . . . . "
        '
        'cmdWszystkoAkcjeSpec
        '
        Me.cmdWszystkoAkcjeSpec.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystkoAkcjeSpec.Location = New System.Drawing.Point(509, 87)
        Me.cmdWszystkoAkcjeSpec.Name = "cmdWszystkoAkcjeSpec"
        Me.cmdWszystkoAkcjeSpec.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystkoAkcjeSpec.TabIndex = 210
        '
        'lblTofi_Akcje
        '
        Me.lblTofi_Akcje.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTofi_Akcje.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTofi_Akcje.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTofi_Akcje.ForeColor = System.Drawing.Color.Purple
        Me.lblTofi_Akcje.Location = New System.Drawing.Point(305, 8)
        Me.lblTofi_Akcje.Name = "lblTofi_Akcje"
        Me.lblTofi_Akcje.Size = New System.Drawing.Size(737, 16)
        Me.lblTofi_Akcje.TabIndex = 108
        '
        'dagAkcjeSpec
        '
        Me.dagAkcjeSpec.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagAkcjeSpec.ContextMenuStrip = Me.menuAkcjeSpec
        Me.dagAkcjeSpec.Location = New System.Drawing.Point(11, 113)
        Me.dagAkcjeSpec.Name = "dagAkcjeSpec"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagAkcjeSpec.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dagAkcjeSpec.Size = New System.Drawing.Size(1031, 496)
        Me.dagAkcjeSpec.TabIndex = 209
        '
        'menuAkcjeSpec
        '
        Me.menuAkcjeSpec.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuAkcjeSpecSingleL, Me.menuAkcjeSpecMultiL, Me.ToolStripSeparator11, Me.menuAkcjeSpecOdswiez})
        Me.menuAkcjeSpec.Name = "ContextMenuStrip1"
        Me.menuAkcjeSpec.Size = New System.Drawing.Size(187, 76)
        '
        'menuAkcjeSpecSingleL
        '
        Me.menuAkcjeSpecSingleL.Name = "menuAkcjeSpecSingleL"
        Me.menuAkcjeSpecSingleL.Size = New System.Drawing.Size(186, 22)
        Me.menuAkcjeSpecSingleL.Text = "Poka¿ w jednej linii"
        '
        'menuAkcjeSpecMultiL
        '
        Me.menuAkcjeSpecMultiL.Name = "menuAkcjeSpecMultiL"
        Me.menuAkcjeSpecMultiL.Size = New System.Drawing.Size(186, 22)
        Me.menuAkcjeSpecMultiL.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator11
        '
        Me.ToolStripSeparator11.Name = "ToolStripSeparator11"
        Me.ToolStripSeparator11.Size = New System.Drawing.Size(183, 6)
        '
        'menuAkcjeSpecOdswiez
        '
        Me.menuAkcjeSpecOdswiez.Name = "menuAkcjeSpecOdswiez"
        Me.menuAkcjeSpecOdswiez.Size = New System.Drawing.Size(186, 22)
        Me.menuAkcjeSpecOdswiez.Text = "Odœwie¿ tabelê"
        '
        'stbFiltryAkcjeSpec
        '
        Me.stbFiltryAkcjeSpec.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltryAkcjeSpec.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltryAkcjeSpec.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltryAkcjeSpec.Location = New System.Drawing.Point(11, 87)
        Me.stbFiltryAkcjeSpec.Name = "stbFiltryAkcjeSpec"
        Me.stbFiltryAkcjeSpec.ShowPanels = True
        Me.stbFiltryAkcjeSpec.Size = New System.Drawing.Size(1031, 22)
        Me.stbFiltryAkcjeSpec.TabIndex = 211
        Me.stbFiltryAkcjeSpec.Text = "StatusBar1"
        '
        'stbAkcjeSpec
        '
        Me.stbAkcjeSpec.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbAkcjeSpec.Dock = System.Windows.Forms.DockStyle.None
        Me.stbAkcjeSpec.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbAkcjeSpec.Location = New System.Drawing.Point(11, 611)
        Me.stbAkcjeSpec.Name = "stbAkcjeSpec"
        Me.stbAkcjeSpec.ShowPanels = True
        Me.stbAkcjeSpec.Size = New System.Drawing.Size(1031, 22)
        Me.stbAkcjeSpec.TabIndex = 208
        Me.stbAkcjeSpec.Text = "stbAkcjeSpec"
        '
        'lblNazwa2_Akcje
        '
        Me.lblNazwa2_Akcje.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa2_Akcje.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa2_Akcje.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa2_Akcje.Location = New System.Drawing.Point(112, 56)
        Me.lblNazwa2_Akcje.Name = "lblNazwa2_Akcje"
        Me.lblNazwa2_Akcje.Size = New System.Drawing.Size(200, 16)
        Me.lblNazwa2_Akcje.TabIndex = 107
        '
        'Label86
        '
        Me.Label86.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label86.ForeColor = System.Drawing.Color.Navy
        Me.Label86.Location = New System.Drawing.Point(8, 56)
        Me.Label86.Name = "Label86"
        Me.Label86.Size = New System.Drawing.Size(112, 16)
        Me.Label86.TabIndex = 115
        Me.Label86.Text = "Osoba kontaktowa . . . . . . . . . . . . . . "
        '
        'lblNazwa1_Akcje
        '
        Me.lblNazwa1_Akcje.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa1_Akcje.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa1_Akcje.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa1_Akcje.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa1_Akcje.Location = New System.Drawing.Point(112, 40)
        Me.lblNazwa1_Akcje.Name = "lblNazwa1_Akcje"
        Me.lblNazwa1_Akcje.Size = New System.Drawing.Size(930, 16)
        Me.lblNazwa1_Akcje.TabIndex = 106
        '
        'lblNazwaHandlowa_Akcje
        '
        Me.lblNazwaHandlowa_Akcje.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwaHandlowa_Akcje.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwaHandlowa_Akcje.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwaHandlowa_Akcje.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwaHandlowa_Akcje.Location = New System.Drawing.Point(112, 24)
        Me.lblNazwaHandlowa_Akcje.Name = "lblNazwaHandlowa_Akcje"
        Me.lblNazwaHandlowa_Akcje.Size = New System.Drawing.Size(930, 16)
        Me.lblNazwaHandlowa_Akcje.TabIndex = 109
        '
        'Label92
        '
        Me.Label92.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label92.ForeColor = System.Drawing.Color.Navy
        Me.Label92.Location = New System.Drawing.Point(8, 40)
        Me.Label92.Name = "Label92"
        Me.Label92.Size = New System.Drawing.Size(112, 16)
        Me.Label92.TabIndex = 110
        Me.Label92.Text = "Adres . . . . . . . . . . . . . . ."
        '
        'lblIdent_Akcje
        '
        Me.lblIdent_Akcje.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent_Akcje.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdent_Akcje.ForeColor = System.Drawing.Color.Purple
        Me.lblIdent_Akcje.Location = New System.Drawing.Point(112, 8)
        Me.lblIdent_Akcje.Name = "lblIdent_Akcje"
        Me.lblIdent_Akcje.Size = New System.Drawing.Size(112, 16)
        Me.lblIdent_Akcje.TabIndex = 105
        '
        'Label95
        '
        Me.Label95.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label95.ForeColor = System.Drawing.Color.Navy
        Me.Label95.Location = New System.Drawing.Point(251, 8)
        Me.Label95.Name = "Label95"
        Me.Label95.Size = New System.Drawing.Size(112, 16)
        Me.Label95.TabIndex = 103
        Me.Label95.Text = "Nr TOFI . . . . . "
        '
        'Label96
        '
        Me.Label96.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label96.ForeColor = System.Drawing.Color.Navy
        Me.Label96.Location = New System.Drawing.Point(8, 24)
        Me.Label96.Name = "Label96"
        Me.Label96.Size = New System.Drawing.Size(112, 16)
        Me.Label96.TabIndex = 102
        Me.Label96.Text = "Nazwa klienta . . . . . . . . ."
        '
        'Label97
        '
        Me.Label97.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label97.ForeColor = System.Drawing.Color.Navy
        Me.Label97.Location = New System.Drawing.Point(8, 8)
        Me.Label97.Name = "Label97"
        Me.Label97.Size = New System.Drawing.Size(112, 16)
        Me.Label97.TabIndex = 101
        Me.Label97.Text = "Identyfikator . . . . . . . . ."
        '
        'Panel8
        '
        Me.Panel8.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel8.Location = New System.Drawing.Point(0, 80)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(1054, 1)
        Me.Panel8.TabIndex = 104
        '
        'Obroty
        '
        Me.Obroty.BackColor = System.Drawing.SystemColors.Control
        Me.Obroty.Controls.Add(Me.lblTelefonKom_Obroty)
        Me.Obroty.Controls.Add(Me.Label127)
        Me.Obroty.Controls.Add(Me.lblTelefon_Obroty)
        Me.Obroty.Controls.Add(Me.Label129)
        Me.Obroty.Controls.Add(Me.lblStanowisko_Obroty)
        Me.Obroty.Controls.Add(Me.Label131)
        Me.Obroty.Controls.Add(Me.cmdWszystkoObroty)
        Me.Obroty.Controls.Add(Me.stbFiltryObroty)
        Me.Obroty.Controls.Add(Me.dagObroty)
        Me.Obroty.Controls.Add(Me.stbObroty)
        Me.Obroty.Controls.Add(Me.lblNazwa2_Obroty)
        Me.Obroty.Controls.Add(Me.Label68)
        Me.Obroty.Controls.Add(Me.cmdEdytujObroty)
        Me.Obroty.Controls.Add(Me.cmdusuñObroty)
        Me.Obroty.Controls.Add(Me.cmdDodajObroty)
        Me.Obroty.Controls.Add(Me.lblNazwa1_Obroty)
        Me.Obroty.Controls.Add(Me.lblNazwaHandlowa_Obroty)
        Me.Obroty.Controls.Add(Me.Label85)
        Me.Obroty.Controls.Add(Me.lblTofi_Obroty)
        Me.Obroty.Controls.Add(Me.lblIdent_Obroty)
        Me.Obroty.Controls.Add(Me.Panel7)
        Me.Obroty.Controls.Add(Me.Label89)
        Me.Obroty.Controls.Add(Me.Label90)
        Me.Obroty.Controls.Add(Me.Label91)
        Me.Obroty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Obroty.ForeColor = System.Drawing.Color.Purple
        Me.Obroty.Location = New System.Drawing.Point(4, 22)
        Me.Obroty.Name = "Obroty"
        Me.Obroty.Size = New System.Drawing.Size(1054, 645)
        Me.Obroty.TabIndex = 9
        Me.Obroty.Text = "Obroty w ost.mies."
        '
        'lblTelefonKom_Obroty
        '
        Me.lblTelefonKom_Obroty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefonKom_Obroty.Location = New System.Drawing.Point(940, 56)
        Me.lblTelefonKom_Obroty.Name = "lblTelefonKom_Obroty"
        Me.lblTelefonKom_Obroty.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefonKom_Obroty.TabIndex = 362
        '
        'Label127
        '
        Me.Label127.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label127.ForeColor = System.Drawing.Color.Navy
        Me.Label127.Location = New System.Drawing.Point(891, 56)
        Me.Label127.Name = "Label127"
        Me.Label127.Size = New System.Drawing.Size(112, 16)
        Me.Label127.TabIndex = 363
        Me.Label127.Text = "Tel. kom. . . . . . . . . . . . . . . "
        '
        'lblTelefon_Obroty
        '
        Me.lblTelefon_Obroty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefon_Obroty.Location = New System.Drawing.Point(673, 56)
        Me.lblTelefon_Obroty.Name = "lblTelefon_Obroty"
        Me.lblTelefon_Obroty.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefon_Obroty.TabIndex = 360
        '
        'Label129
        '
        Me.Label129.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label129.ForeColor = System.Drawing.Color.Navy
        Me.Label129.Location = New System.Drawing.Point(624, 56)
        Me.Label129.Name = "Label129"
        Me.Label129.Size = New System.Drawing.Size(112, 16)
        Me.Label129.TabIndex = 361
        Me.Label129.Text = "Telefon . . . . . . . . . . . . . . "
        '
        'lblStanowisko_Obroty
        '
        Me.lblStanowisko_Obroty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStanowisko_Obroty.Location = New System.Drawing.Point(404, 56)
        Me.lblStanowisko_Obroty.Name = "lblStanowisko_Obroty"
        Me.lblStanowisko_Obroty.Size = New System.Drawing.Size(200, 16)
        Me.lblStanowisko_Obroty.TabIndex = 358
        '
        'Label131
        '
        Me.Label131.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label131.ForeColor = System.Drawing.Color.Navy
        Me.Label131.Location = New System.Drawing.Point(333, 56)
        Me.Label131.Name = "Label131"
        Me.Label131.Size = New System.Drawing.Size(112, 16)
        Me.Label131.TabIndex = 359
        Me.Label131.Text = "Stanowisko . . . . . . . . . . . . . . "
        '
        'cmdWszystkoObroty
        '
        Me.cmdWszystkoObroty.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystkoObroty.Location = New System.Drawing.Point(618, 87)
        Me.cmdWszystkoObroty.Name = "cmdWszystkoObroty"
        Me.cmdWszystkoObroty.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystkoObroty.TabIndex = 356
        '
        'stbFiltryObroty
        '
        Me.stbFiltryObroty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltryObroty.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltryObroty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltryObroty.Location = New System.Drawing.Point(11, 87)
        Me.stbFiltryObroty.Name = "stbFiltryObroty"
        Me.stbFiltryObroty.ShowPanels = True
        Me.stbFiltryObroty.Size = New System.Drawing.Size(1031, 22)
        Me.stbFiltryObroty.TabIndex = 357
        Me.stbFiltryObroty.Text = "StatusBar1"
        '
        'dagObroty
        '
        Me.dagObroty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagObroty.ContextMenuStrip = Me.menuObroty
        Me.dagObroty.Location = New System.Drawing.Point(11, 113)
        Me.dagObroty.Name = "dagObroty"
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagObroty.RowHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.dagObroty.Size = New System.Drawing.Size(1031, 470)
        Me.dagObroty.TabIndex = 355
        '
        'menuObroty
        '
        Me.menuObroty.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuObrotyEdytuj, Me.menuObrotyDodaj, Me.menuObrotyUsun, Me.ToolStripSeparator9, Me.menuObrotySingleL, Me.menuObrotyMultiL, Me.ToolStripSeparator10, Me.menuObrotyOdswiez})
        Me.menuObroty.Name = "ContextMenuStrip1"
        Me.menuObroty.Size = New System.Drawing.Size(187, 148)
        '
        'menuObrotyEdytuj
        '
        Me.menuObrotyEdytuj.Enabled = False
        Me.menuObrotyEdytuj.Name = "menuObrotyEdytuj"
        Me.menuObrotyEdytuj.Size = New System.Drawing.Size(186, 22)
        Me.menuObrotyEdytuj.Text = "Edytuj"
        '
        'menuObrotyDodaj
        '
        Me.menuObrotyDodaj.Enabled = False
        Me.menuObrotyDodaj.Name = "menuObrotyDodaj"
        Me.menuObrotyDodaj.Size = New System.Drawing.Size(186, 22)
        Me.menuObrotyDodaj.Text = "Dodaj"
        '
        'menuObrotyUsun
        '
        Me.menuObrotyUsun.Enabled = False
        Me.menuObrotyUsun.Name = "menuObrotyUsun"
        Me.menuObrotyUsun.Size = New System.Drawing.Size(186, 22)
        Me.menuObrotyUsun.Text = "Usuñ"
        '
        'ToolStripSeparator9
        '
        Me.ToolStripSeparator9.Name = "ToolStripSeparator9"
        Me.ToolStripSeparator9.Size = New System.Drawing.Size(183, 6)
        '
        'menuObrotySingleL
        '
        Me.menuObrotySingleL.Name = "menuObrotySingleL"
        Me.menuObrotySingleL.Size = New System.Drawing.Size(186, 22)
        Me.menuObrotySingleL.Text = "Poka¿ w jednej linii"
        '
        'menuObrotyMultiL
        '
        Me.menuObrotyMultiL.Name = "menuObrotyMultiL"
        Me.menuObrotyMultiL.Size = New System.Drawing.Size(186, 22)
        Me.menuObrotyMultiL.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator10
        '
        Me.ToolStripSeparator10.Name = "ToolStripSeparator10"
        Me.ToolStripSeparator10.Size = New System.Drawing.Size(183, 6)
        '
        'menuObrotyOdswiez
        '
        Me.menuObrotyOdswiez.Name = "menuObrotyOdswiez"
        Me.menuObrotyOdswiez.Size = New System.Drawing.Size(186, 22)
        Me.menuObrotyOdswiez.Text = "Odœwie¿ tabelê"
        '
        'stbObroty
        '
        Me.stbObroty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbObroty.Dock = System.Windows.Forms.DockStyle.None
        Me.stbObroty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbObroty.Location = New System.Drawing.Point(11, 585)
        Me.stbObroty.Name = "stbObroty"
        Me.stbObroty.ShowPanels = True
        Me.stbObroty.Size = New System.Drawing.Size(1031, 22)
        Me.stbObroty.TabIndex = 354
        Me.stbObroty.Text = "StatusBar1"
        '
        'lblNazwa2_Obroty
        '
        Me.lblNazwa2_Obroty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa2_Obroty.Location = New System.Drawing.Point(112, 56)
        Me.lblNazwa2_Obroty.Name = "lblNazwa2_Obroty"
        Me.lblNazwa2_Obroty.Size = New System.Drawing.Size(200, 16)
        Me.lblNazwa2_Obroty.TabIndex = 70
        '
        'Label68
        '
        Me.Label68.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label68.ForeColor = System.Drawing.Color.Navy
        Me.Label68.Location = New System.Drawing.Point(8, 56)
        Me.Label68.Name = "Label68"
        Me.Label68.Size = New System.Drawing.Size(112, 16)
        Me.Label68.TabIndex = 100
        Me.Label68.Text = "Osoba kontaktowa . . . . . . . . . . . . . . "
        '
        'cmdEdytujObroty
        '
        Me.cmdEdytujObroty.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytujObroty.Enabled = False
        Me.cmdEdytujObroty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytujObroty.ForeColor = System.Drawing.Color.Black
        Me.cmdEdytujObroty.Image = CType(resources.GetObject("cmdEdytujObroty.Image"), System.Drawing.Image)
        Me.cmdEdytujObroty.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytujObroty.Location = New System.Drawing.Point(962, 613)
        Me.cmdEdytujObroty.Name = "cmdEdytujObroty"
        Me.cmdEdytujObroty.Size = New System.Drawing.Size(80, 23)
        Me.cmdEdytujObroty.TabIndex = 353
        Me.cmdEdytujObroty.Text = "Edytuj"
        Me.cmdEdytujObroty.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdusuñObroty
        '
        Me.cmdusuñObroty.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdusuñObroty.Enabled = False
        Me.cmdusuñObroty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdusuñObroty.ForeColor = System.Drawing.Color.Black
        Me.cmdusuñObroty.Image = CType(resources.GetObject("cmdusuñObroty.Image"), System.Drawing.Image)
        Me.cmdusuñObroty.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdusuñObroty.Location = New System.Drawing.Point(874, 613)
        Me.cmdusuñObroty.Name = "cmdusuñObroty"
        Me.cmdusuñObroty.Size = New System.Drawing.Size(80, 23)
        Me.cmdusuñObroty.TabIndex = 352
        Me.cmdusuñObroty.Text = "Usuñ"
        Me.cmdusuñObroty.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdDodajObroty
        '
        Me.cmdDodajObroty.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodajObroty.Enabled = False
        Me.cmdDodajObroty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodajObroty.ForeColor = System.Drawing.Color.Black
        Me.cmdDodajObroty.Image = CType(resources.GetObject("cmdDodajObroty.Image"), System.Drawing.Image)
        Me.cmdDodajObroty.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodajObroty.Location = New System.Drawing.Point(785, 613)
        Me.cmdDodajObroty.Name = "cmdDodajObroty"
        Me.cmdDodajObroty.Size = New System.Drawing.Size(80, 23)
        Me.cmdDodajObroty.TabIndex = 351
        Me.cmdDodajObroty.Text = "Dodaj"
        Me.cmdDodajObroty.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblNazwa1_Obroty
        '
        Me.lblNazwa1_Obroty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa1_Obroty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa1_Obroty.Location = New System.Drawing.Point(112, 40)
        Me.lblNazwa1_Obroty.Name = "lblNazwa1_Obroty"
        Me.lblNazwa1_Obroty.Size = New System.Drawing.Size(930, 16)
        Me.lblNazwa1_Obroty.TabIndex = 69
        '
        'lblNazwaHandlowa_Obroty
        '
        Me.lblNazwaHandlowa_Obroty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwaHandlowa_Obroty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwaHandlowa_Obroty.Location = New System.Drawing.Point(112, 24)
        Me.lblNazwaHandlowa_Obroty.Name = "lblNazwaHandlowa_Obroty"
        Me.lblNazwaHandlowa_Obroty.Size = New System.Drawing.Size(930, 16)
        Me.lblNazwaHandlowa_Obroty.TabIndex = 72
        '
        'Label85
        '
        Me.Label85.ForeColor = System.Drawing.Color.Navy
        Me.Label85.Location = New System.Drawing.Point(8, 40)
        Me.Label85.Name = "Label85"
        Me.Label85.Size = New System.Drawing.Size(112, 16)
        Me.Label85.TabIndex = 73
        Me.Label85.Text = "Adres . . . . . . . . . . . . . . ."
        '
        'lblTofi_Obroty
        '
        Me.lblTofi_Obroty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTofi_Obroty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTofi_Obroty.Location = New System.Drawing.Point(305, 8)
        Me.lblTofi_Obroty.Name = "lblTofi_Obroty"
        Me.lblTofi_Obroty.Size = New System.Drawing.Size(737, 16)
        Me.lblTofi_Obroty.TabIndex = 71
        '
        'lblIdent_Obroty
        '
        Me.lblIdent_Obroty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent_Obroty.Location = New System.Drawing.Point(112, 8)
        Me.lblIdent_Obroty.Name = "lblIdent_Obroty"
        Me.lblIdent_Obroty.Size = New System.Drawing.Size(112, 16)
        Me.lblIdent_Obroty.TabIndex = 68
        '
        'Panel7
        '
        Me.Panel7.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel7.Location = New System.Drawing.Point(0, 80)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(1054, 1)
        Me.Panel7.TabIndex = 67
        '
        'Label89
        '
        Me.Label89.ForeColor = System.Drawing.Color.Navy
        Me.Label89.Location = New System.Drawing.Point(256, 8)
        Me.Label89.Name = "Label89"
        Me.Label89.Size = New System.Drawing.Size(112, 16)
        Me.Label89.TabIndex = 66
        Me.Label89.Text = "Nr TOFI . . . . . "
        '
        'Label90
        '
        Me.Label90.ForeColor = System.Drawing.Color.Navy
        Me.Label90.Location = New System.Drawing.Point(8, 24)
        Me.Label90.Name = "Label90"
        Me.Label90.Size = New System.Drawing.Size(112, 16)
        Me.Label90.TabIndex = 65
        Me.Label90.Text = "Nazwa klienta . . . . . . . . ."
        '
        'Label91
        '
        Me.Label91.ForeColor = System.Drawing.Color.Navy
        Me.Label91.Location = New System.Drawing.Point(8, 8)
        Me.Label91.Name = "Label91"
        Me.Label91.Size = New System.Drawing.Size(112, 16)
        Me.Label91.TabIndex = 64
        Me.Label91.Text = "Identyfikator . . . . . . . . ."
        '
        'ObrotyMies
        '
        Me.ObrotyMies.BackColor = System.Drawing.SystemColors.Control
        Me.ObrotyMies.Controls.Add(Me.lblTelefonKom_ObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.Label121)
        Me.ObrotyMies.Controls.Add(Me.lblTelefon_ObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.Label123)
        Me.ObrotyMies.Controls.Add(Me.lblStanowisko_ObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.Label125)
        Me.ObrotyMies.Controls.Add(Me.cmdWszystkoObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.stbFiltryObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.dagObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.stbObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.lblNazwa2_ObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.Label69)
        Me.ObrotyMies.Controls.Add(Me.cmdEdytujObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.cmdUsuñObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.cmdDodajObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.lblNazwa1_ObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.lblNazwaHandlowa_ObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.Label76)
        Me.ObrotyMies.Controls.Add(Me.lblTofi_ObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.lblIdent_ObrotyMies)
        Me.ObrotyMies.Controls.Add(Me.Panel6)
        Me.ObrotyMies.Controls.Add(Me.Label80)
        Me.ObrotyMies.Controls.Add(Me.Label81)
        Me.ObrotyMies.Controls.Add(Me.Label82)
        Me.ObrotyMies.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.ObrotyMies.ForeColor = System.Drawing.Color.Purple
        Me.ObrotyMies.Location = New System.Drawing.Point(4, 22)
        Me.ObrotyMies.Name = "ObrotyMies"
        Me.ObrotyMies.Size = New System.Drawing.Size(1054, 645)
        Me.ObrotyMies.TabIndex = 8
        Me.ObrotyMies.Text = "Obroty miesiêczne."
        '
        'lblTelefonKom_ObrotyMies
        '
        Me.lblTelefonKom_ObrotyMies.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefonKom_ObrotyMies.Location = New System.Drawing.Point(940, 56)
        Me.lblTelefonKom_ObrotyMies.Name = "lblTelefonKom_ObrotyMies"
        Me.lblTelefonKom_ObrotyMies.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefonKom_ObrotyMies.TabIndex = 312
        '
        'Label121
        '
        Me.Label121.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label121.ForeColor = System.Drawing.Color.Navy
        Me.Label121.Location = New System.Drawing.Point(891, 56)
        Me.Label121.Name = "Label121"
        Me.Label121.Size = New System.Drawing.Size(112, 16)
        Me.Label121.TabIndex = 313
        Me.Label121.Text = "Tel. kom. . . . . . . . . . . . . . . "
        '
        'lblTelefon_ObrotyMies
        '
        Me.lblTelefon_ObrotyMies.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefon_ObrotyMies.Location = New System.Drawing.Point(673, 56)
        Me.lblTelefon_ObrotyMies.Name = "lblTelefon_ObrotyMies"
        Me.lblTelefon_ObrotyMies.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefon_ObrotyMies.TabIndex = 310
        '
        'Label123
        '
        Me.Label123.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label123.ForeColor = System.Drawing.Color.Navy
        Me.Label123.Location = New System.Drawing.Point(624, 56)
        Me.Label123.Name = "Label123"
        Me.Label123.Size = New System.Drawing.Size(112, 16)
        Me.Label123.TabIndex = 311
        Me.Label123.Text = "Telefon . . . . . . . . . . . . . . "
        '
        'lblStanowisko_ObrotyMies
        '
        Me.lblStanowisko_ObrotyMies.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStanowisko_ObrotyMies.Location = New System.Drawing.Point(404, 56)
        Me.lblStanowisko_ObrotyMies.Name = "lblStanowisko_ObrotyMies"
        Me.lblStanowisko_ObrotyMies.Size = New System.Drawing.Size(200, 16)
        Me.lblStanowisko_ObrotyMies.TabIndex = 308
        '
        'Label125
        '
        Me.Label125.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label125.ForeColor = System.Drawing.Color.Navy
        Me.Label125.Location = New System.Drawing.Point(333, 56)
        Me.Label125.Name = "Label125"
        Me.Label125.Size = New System.Drawing.Size(112, 16)
        Me.Label125.TabIndex = 309
        Me.Label125.Text = "Stanowisko . . . . . . . . . . . . . . "
        '
        'cmdWszystkoObrotyMies
        '
        Me.cmdWszystkoObrotyMies.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystkoObrotyMies.Location = New System.Drawing.Point(618, 87)
        Me.cmdWszystkoObrotyMies.Name = "cmdWszystkoObrotyMies"
        Me.cmdWszystkoObrotyMies.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystkoObrotyMies.TabIndex = 306
        '
        'stbFiltryObrotyMies
        '
        Me.stbFiltryObrotyMies.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltryObrotyMies.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltryObrotyMies.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltryObrotyMies.Location = New System.Drawing.Point(11, 87)
        Me.stbFiltryObrotyMies.Name = "stbFiltryObrotyMies"
        Me.stbFiltryObrotyMies.ShowPanels = True
        Me.stbFiltryObrotyMies.Size = New System.Drawing.Size(1031, 22)
        Me.stbFiltryObrotyMies.TabIndex = 307
        Me.stbFiltryObrotyMies.Text = "StatusBar1"
        '
        'dagObrotyMies
        '
        Me.dagObrotyMies.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagObrotyMies.ContextMenuStrip = Me.menuObrotyMies
        Me.dagObrotyMies.Location = New System.Drawing.Point(11, 113)
        Me.dagObrotyMies.Name = "dagObrotyMies"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagObrotyMies.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dagObrotyMies.Size = New System.Drawing.Size(1031, 470)
        Me.dagObrotyMies.TabIndex = 305
        '
        'menuObrotyMies
        '
        Me.menuObrotyMies.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuObrotyMiesEdytuj, Me.menuObrotyMiesDodaj, Me.menuObrotyMiesUsun, Me.ToolStripSeparator7, Me.menuObrotyMiesSingleL, Me.menuObrotyMiesMultiL, Me.ToolStripSeparator8, Me.menuObrotyMiesOdswiez})
        Me.menuObrotyMies.Name = "ContextMenuStrip1"
        Me.menuObrotyMies.Size = New System.Drawing.Size(187, 148)
        '
        'menuObrotyMiesEdytuj
        '
        Me.menuObrotyMiesEdytuj.Enabled = False
        Me.menuObrotyMiesEdytuj.Name = "menuObrotyMiesEdytuj"
        Me.menuObrotyMiesEdytuj.Size = New System.Drawing.Size(186, 22)
        Me.menuObrotyMiesEdytuj.Text = "Edytuj"
        '
        'menuObrotyMiesDodaj
        '
        Me.menuObrotyMiesDodaj.Enabled = False
        Me.menuObrotyMiesDodaj.Name = "menuObrotyMiesDodaj"
        Me.menuObrotyMiesDodaj.Size = New System.Drawing.Size(186, 22)
        Me.menuObrotyMiesDodaj.Text = "Dodaj"
        '
        'menuObrotyMiesUsun
        '
        Me.menuObrotyMiesUsun.Enabled = False
        Me.menuObrotyMiesUsun.Name = "menuObrotyMiesUsun"
        Me.menuObrotyMiesUsun.Size = New System.Drawing.Size(186, 22)
        Me.menuObrotyMiesUsun.Text = "Usuñ"
        '
        'ToolStripSeparator7
        '
        Me.ToolStripSeparator7.Name = "ToolStripSeparator7"
        Me.ToolStripSeparator7.Size = New System.Drawing.Size(183, 6)
        '
        'menuObrotyMiesSingleL
        '
        Me.menuObrotyMiesSingleL.Name = "menuObrotyMiesSingleL"
        Me.menuObrotyMiesSingleL.Size = New System.Drawing.Size(186, 22)
        Me.menuObrotyMiesSingleL.Text = "Poka¿ w jednej linii"
        '
        'menuObrotyMiesMultiL
        '
        Me.menuObrotyMiesMultiL.Name = "menuObrotyMiesMultiL"
        Me.menuObrotyMiesMultiL.Size = New System.Drawing.Size(186, 22)
        Me.menuObrotyMiesMultiL.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator8
        '
        Me.ToolStripSeparator8.Name = "ToolStripSeparator8"
        Me.ToolStripSeparator8.Size = New System.Drawing.Size(183, 6)
        '
        'menuObrotyMiesOdswiez
        '
        Me.menuObrotyMiesOdswiez.Name = "menuObrotyMiesOdswiez"
        Me.menuObrotyMiesOdswiez.Size = New System.Drawing.Size(186, 22)
        Me.menuObrotyMiesOdswiez.Text = "Odœwie¿ tabelê"
        '
        'stbObrotyMies
        '
        Me.stbObrotyMies.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbObrotyMies.Dock = System.Windows.Forms.DockStyle.None
        Me.stbObrotyMies.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbObrotyMies.Location = New System.Drawing.Point(11, 585)
        Me.stbObrotyMies.Name = "stbObrotyMies"
        Me.stbObrotyMies.ShowPanels = True
        Me.stbObrotyMies.Size = New System.Drawing.Size(1031, 22)
        Me.stbObrotyMies.TabIndex = 304
        Me.stbObrotyMies.Text = "StatusBar1"
        '
        'lblNazwa2_ObrotyMies
        '
        Me.lblNazwa2_ObrotyMies.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa2_ObrotyMies.Location = New System.Drawing.Point(112, 56)
        Me.lblNazwa2_ObrotyMies.Name = "lblNazwa2_ObrotyMies"
        Me.lblNazwa2_ObrotyMies.Size = New System.Drawing.Size(200, 16)
        Me.lblNazwa2_ObrotyMies.TabIndex = 70
        '
        'Label69
        '
        Me.Label69.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label69.ForeColor = System.Drawing.Color.Navy
        Me.Label69.Location = New System.Drawing.Point(8, 56)
        Me.Label69.Name = "Label69"
        Me.Label69.Size = New System.Drawing.Size(112, 16)
        Me.Label69.TabIndex = 100
        Me.Label69.Text = "Osoba kontaktowa . . . . . . . . . . . . . . "
        '
        'cmdEdytujObrotyMies
        '
        Me.cmdEdytujObrotyMies.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytujObrotyMies.Enabled = False
        Me.cmdEdytujObrotyMies.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytujObrotyMies.ForeColor = System.Drawing.Color.Black
        Me.cmdEdytujObrotyMies.Image = CType(resources.GetObject("cmdEdytujObrotyMies.Image"), System.Drawing.Image)
        Me.cmdEdytujObrotyMies.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytujObrotyMies.Location = New System.Drawing.Point(960, 613)
        Me.cmdEdytujObrotyMies.Name = "cmdEdytujObrotyMies"
        Me.cmdEdytujObrotyMies.Size = New System.Drawing.Size(80, 23)
        Me.cmdEdytujObrotyMies.TabIndex = 303
        Me.cmdEdytujObrotyMies.Text = "Edytuj"
        Me.cmdEdytujObrotyMies.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdUsuñObrotyMies
        '
        Me.cmdUsuñObrotyMies.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUsuñObrotyMies.Enabled = False
        Me.cmdUsuñObrotyMies.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdUsuñObrotyMies.ForeColor = System.Drawing.Color.Black
        Me.cmdUsuñObrotyMies.Image = CType(resources.GetObject("cmdUsuñObrotyMies.Image"), System.Drawing.Image)
        Me.cmdUsuñObrotyMies.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUsuñObrotyMies.Location = New System.Drawing.Point(873, 613)
        Me.cmdUsuñObrotyMies.Name = "cmdUsuñObrotyMies"
        Me.cmdUsuñObrotyMies.Size = New System.Drawing.Size(80, 23)
        Me.cmdUsuñObrotyMies.TabIndex = 302
        Me.cmdUsuñObrotyMies.Text = "Usuñ"
        Me.cmdUsuñObrotyMies.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdDodajObrotyMies
        '
        Me.cmdDodajObrotyMies.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodajObrotyMies.Enabled = False
        Me.cmdDodajObrotyMies.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodajObrotyMies.ForeColor = System.Drawing.Color.Black
        Me.cmdDodajObrotyMies.Image = CType(resources.GetObject("cmdDodajObrotyMies.Image"), System.Drawing.Image)
        Me.cmdDodajObrotyMies.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodajObrotyMies.Location = New System.Drawing.Point(785, 613)
        Me.cmdDodajObrotyMies.Name = "cmdDodajObrotyMies"
        Me.cmdDodajObrotyMies.Size = New System.Drawing.Size(80, 23)
        Me.cmdDodajObrotyMies.TabIndex = 301
        Me.cmdDodajObrotyMies.Text = "Dodaj"
        Me.cmdDodajObrotyMies.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblNazwa1_ObrotyMies
        '
        Me.lblNazwa1_ObrotyMies.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa1_ObrotyMies.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa1_ObrotyMies.Location = New System.Drawing.Point(112, 40)
        Me.lblNazwa1_ObrotyMies.Name = "lblNazwa1_ObrotyMies"
        Me.lblNazwa1_ObrotyMies.Size = New System.Drawing.Size(928, 16)
        Me.lblNazwa1_ObrotyMies.TabIndex = 69
        '
        'lblNazwaHandlowa_ObrotyMies
        '
        Me.lblNazwaHandlowa_ObrotyMies.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwaHandlowa_ObrotyMies.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwaHandlowa_ObrotyMies.Location = New System.Drawing.Point(112, 24)
        Me.lblNazwaHandlowa_ObrotyMies.Name = "lblNazwaHandlowa_ObrotyMies"
        Me.lblNazwaHandlowa_ObrotyMies.Size = New System.Drawing.Size(928, 16)
        Me.lblNazwaHandlowa_ObrotyMies.TabIndex = 72
        '
        'Label76
        '
        Me.Label76.ForeColor = System.Drawing.Color.Navy
        Me.Label76.Location = New System.Drawing.Point(8, 40)
        Me.Label76.Name = "Label76"
        Me.Label76.Size = New System.Drawing.Size(112, 16)
        Me.Label76.TabIndex = 73
        Me.Label76.Text = "Adres . . . . . . . . . . . . . . ."
        '
        'lblTofi_ObrotyMies
        '
        Me.lblTofi_ObrotyMies.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTofi_ObrotyMies.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTofi_ObrotyMies.Location = New System.Drawing.Point(305, 8)
        Me.lblTofi_ObrotyMies.Name = "lblTofi_ObrotyMies"
        Me.lblTofi_ObrotyMies.Size = New System.Drawing.Size(735, 16)
        Me.lblTofi_ObrotyMies.TabIndex = 71
        '
        'lblIdent_ObrotyMies
        '
        Me.lblIdent_ObrotyMies.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent_ObrotyMies.Location = New System.Drawing.Point(112, 8)
        Me.lblIdent_ObrotyMies.Name = "lblIdent_ObrotyMies"
        Me.lblIdent_ObrotyMies.Size = New System.Drawing.Size(112, 16)
        Me.lblIdent_ObrotyMies.TabIndex = 68
        '
        'Panel6
        '
        Me.Panel6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel6.Location = New System.Drawing.Point(0, 80)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(1054, 1)
        Me.Panel6.TabIndex = 67
        '
        'Label80
        '
        Me.Label80.ForeColor = System.Drawing.Color.Navy
        Me.Label80.Location = New System.Drawing.Point(256, 8)
        Me.Label80.Name = "Label80"
        Me.Label80.Size = New System.Drawing.Size(112, 16)
        Me.Label80.TabIndex = 66
        Me.Label80.Text = "Nr TOFI . . . . . "
        '
        'Label81
        '
        Me.Label81.ForeColor = System.Drawing.Color.Navy
        Me.Label81.Location = New System.Drawing.Point(8, 24)
        Me.Label81.Name = "Label81"
        Me.Label81.Size = New System.Drawing.Size(112, 16)
        Me.Label81.TabIndex = 65
        Me.Label81.Text = "Nazwa klienta . . . . . . . . ."
        '
        'Label82
        '
        Me.Label82.ForeColor = System.Drawing.Color.Navy
        Me.Label82.Location = New System.Drawing.Point(8, 8)
        Me.Label82.Name = "Label82"
        Me.Label82.Size = New System.Drawing.Size(112, 16)
        Me.Label82.TabIndex = 64
        Me.Label82.Text = "Identyfikator . . . . . . . . ."
        '
        'AnalizyZakupu
        '
        Me.AnalizyZakupu.BackColor = System.Drawing.SystemColors.Control
        Me.AnalizyZakupu.Controls.Add(Me.lblTelefonKom_Analizy)
        Me.AnalizyZakupu.Controls.Add(Me.Label115)
        Me.AnalizyZakupu.Controls.Add(Me.lblTelefon_Analizy)
        Me.AnalizyZakupu.Controls.Add(Me.Label117)
        Me.AnalizyZakupu.Controls.Add(Me.lblStanowisko_Analizy)
        Me.AnalizyZakupu.Controls.Add(Me.Label119)
        Me.AnalizyZakupu.Controls.Add(Me.cmdWszystkoAnalizyZakupu)
        Me.AnalizyZakupu.Controls.Add(Me.stbFiltryAnalizyZakupu)
        Me.AnalizyZakupu.Controls.Add(Me.dagAnalizyZakupu)
        Me.AnalizyZakupu.Controls.Add(Me.stbAnalizyZakupu)
        Me.AnalizyZakupu.Controls.Add(Me.lblNazwa2_Analizy)
        Me.AnalizyZakupu.Controls.Add(Me.Label93)
        Me.AnalizyZakupu.Controls.Add(Me.lblNazwa1_Analizy)
        Me.AnalizyZakupu.Controls.Add(Me.lblNazwaHandlowa_Analizy)
        Me.AnalizyZakupu.Controls.Add(Me.Label99)
        Me.AnalizyZakupu.Controls.Add(Me.lblTOFI_Analizy)
        Me.AnalizyZakupu.Controls.Add(Me.lblIdent_Analizy)
        Me.AnalizyZakupu.Controls.Add(Me.Panel13)
        Me.AnalizyZakupu.Controls.Add(Me.Label102)
        Me.AnalizyZakupu.Controls.Add(Me.Label103)
        Me.AnalizyZakupu.Controls.Add(Me.Label104)
        Me.AnalizyZakupu.Location = New System.Drawing.Point(4, 22)
        Me.AnalizyZakupu.Name = "AnalizyZakupu"
        Me.AnalizyZakupu.Size = New System.Drawing.Size(1054, 645)
        Me.AnalizyZakupu.TabIndex = 14
        Me.AnalizyZakupu.Text = "Analizy zakupu"
        '
        'lblTelefonKom_Analizy
        '
        Me.lblTelefonKom_Analizy.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefonKom_Analizy.Location = New System.Drawing.Point(940, 56)
        Me.lblTelefonKom_Analizy.Name = "lblTelefonKom_Analizy"
        Me.lblTelefonKom_Analizy.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefonKom_Analizy.TabIndex = 262
        '
        'Label115
        '
        Me.Label115.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label115.ForeColor = System.Drawing.Color.Navy
        Me.Label115.Location = New System.Drawing.Point(891, 56)
        Me.Label115.Name = "Label115"
        Me.Label115.Size = New System.Drawing.Size(112, 16)
        Me.Label115.TabIndex = 263
        Me.Label115.Text = "Tel. kom. . . . . . . . . . . . . . . "
        '
        'lblTelefon_Analizy
        '
        Me.lblTelefon_Analizy.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefon_Analizy.Location = New System.Drawing.Point(673, 56)
        Me.lblTelefon_Analizy.Name = "lblTelefon_Analizy"
        Me.lblTelefon_Analizy.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefon_Analizy.TabIndex = 260
        '
        'Label117
        '
        Me.Label117.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label117.ForeColor = System.Drawing.Color.Navy
        Me.Label117.Location = New System.Drawing.Point(624, 56)
        Me.Label117.Name = "Label117"
        Me.Label117.Size = New System.Drawing.Size(112, 16)
        Me.Label117.TabIndex = 261
        Me.Label117.Text = "Telefon . . . . . . . . . . . . . . "
        '
        'lblStanowisko_Analizy
        '
        Me.lblStanowisko_Analizy.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStanowisko_Analizy.Location = New System.Drawing.Point(404, 56)
        Me.lblStanowisko_Analizy.Name = "lblStanowisko_Analizy"
        Me.lblStanowisko_Analizy.Size = New System.Drawing.Size(200, 16)
        Me.lblStanowisko_Analizy.TabIndex = 258
        '
        'Label119
        '
        Me.Label119.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label119.ForeColor = System.Drawing.Color.Navy
        Me.Label119.Location = New System.Drawing.Point(333, 56)
        Me.Label119.Name = "Label119"
        Me.Label119.Size = New System.Drawing.Size(112, 16)
        Me.Label119.TabIndex = 259
        Me.Label119.Text = "Stanowisko . . . . . . . . . . . . . . "
        '
        'cmdWszystkoAnalizyZakupu
        '
        Me.cmdWszystkoAnalizyZakupu.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystkoAnalizyZakupu.Location = New System.Drawing.Point(618, 86)
        Me.cmdWszystkoAnalizyZakupu.Name = "cmdWszystkoAnalizyZakupu"
        Me.cmdWszystkoAnalizyZakupu.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystkoAnalizyZakupu.TabIndex = 256
        '
        'stbFiltryAnalizyZakupu
        '
        Me.stbFiltryAnalizyZakupu.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltryAnalizyZakupu.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltryAnalizyZakupu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltryAnalizyZakupu.Location = New System.Drawing.Point(11, 86)
        Me.stbFiltryAnalizyZakupu.Name = "stbFiltryAnalizyZakupu"
        Me.stbFiltryAnalizyZakupu.ShowPanels = True
        Me.stbFiltryAnalizyZakupu.Size = New System.Drawing.Size(1031, 22)
        Me.stbFiltryAnalizyZakupu.TabIndex = 257
        Me.stbFiltryAnalizyZakupu.Text = "StatusBar1"
        '
        'dagAnalizyZakupu
        '
        Me.dagAnalizyZakupu.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagAnalizyZakupu.ContextMenuStrip = Me.menuAnalizyZakupu
        Me.dagAnalizyZakupu.Location = New System.Drawing.Point(11, 112)
        Me.dagAnalizyZakupu.Name = "dagAnalizyZakupu"
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagAnalizyZakupu.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dagAnalizyZakupu.Size = New System.Drawing.Size(1031, 496)
        Me.dagAnalizyZakupu.TabIndex = 255
        '
        'menuAnalizyZakupu
        '
        Me.menuAnalizyZakupu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuAnalizyZakupuSingleL, Me.menuAnalizyZakupuMultiL, Me.ToolStripSeparator6, Me.menuAnalizyZakupuOdswiez})
        Me.menuAnalizyZakupu.Name = "ContextMenuStrip1"
        Me.menuAnalizyZakupu.Size = New System.Drawing.Size(187, 76)
        '
        'menuAnalizyZakupuSingleL
        '
        Me.menuAnalizyZakupuSingleL.Name = "menuAnalizyZakupuSingleL"
        Me.menuAnalizyZakupuSingleL.Size = New System.Drawing.Size(186, 22)
        Me.menuAnalizyZakupuSingleL.Text = "Poka¿ w jednej linii"
        '
        'menuAnalizyZakupuMultiL
        '
        Me.menuAnalizyZakupuMultiL.Name = "menuAnalizyZakupuMultiL"
        Me.menuAnalizyZakupuMultiL.Size = New System.Drawing.Size(186, 22)
        Me.menuAnalizyZakupuMultiL.Text = "Poka¿ w wielu liniach"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(183, 6)
        '
        'menuAnalizyZakupuOdswiez
        '
        Me.menuAnalizyZakupuOdswiez.Name = "menuAnalizyZakupuOdswiez"
        Me.menuAnalizyZakupuOdswiez.Size = New System.Drawing.Size(186, 22)
        Me.menuAnalizyZakupuOdswiez.Text = "Odœwie¿ tabelê"
        '
        'stbAnalizyZakupu
        '
        Me.stbAnalizyZakupu.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbAnalizyZakupu.Dock = System.Windows.Forms.DockStyle.None
        Me.stbAnalizyZakupu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbAnalizyZakupu.Location = New System.Drawing.Point(11, 611)
        Me.stbAnalizyZakupu.Name = "stbAnalizyZakupu"
        Me.stbAnalizyZakupu.ShowPanels = True
        Me.stbAnalizyZakupu.Size = New System.Drawing.Size(1031, 22)
        Me.stbAnalizyZakupu.TabIndex = 254
        Me.stbAnalizyZakupu.Text = "StatusBar1"
        '
        'lblNazwa2_Analizy
        '
        Me.lblNazwa2_Analizy.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa2_Analizy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa2_Analizy.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa2_Analizy.Location = New System.Drawing.Point(112, 55)
        Me.lblNazwa2_Analizy.Name = "lblNazwa2_Analizy"
        Me.lblNazwa2_Analizy.Size = New System.Drawing.Size(200, 16)
        Me.lblNazwa2_Analizy.TabIndex = 107
        '
        'Label93
        '
        Me.Label93.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label93.ForeColor = System.Drawing.Color.Navy
        Me.Label93.Location = New System.Drawing.Point(8, 55)
        Me.Label93.Name = "Label93"
        Me.Label93.Size = New System.Drawing.Size(112, 16)
        Me.Label93.TabIndex = 115
        Me.Label93.Text = "Osoba kontaktowa . . . . . . . . . . . . . . "
        '
        'lblNazwa1_Analizy
        '
        Me.lblNazwa1_Analizy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa1_Analizy.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa1_Analizy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa1_Analizy.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa1_Analizy.Location = New System.Drawing.Point(112, 39)
        Me.lblNazwa1_Analizy.Name = "lblNazwa1_Analizy"
        Me.lblNazwa1_Analizy.Size = New System.Drawing.Size(930, 16)
        Me.lblNazwa1_Analizy.TabIndex = 106
        '
        'lblNazwaHandlowa_Analizy
        '
        Me.lblNazwaHandlowa_Analizy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwaHandlowa_Analizy.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwaHandlowa_Analizy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwaHandlowa_Analizy.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwaHandlowa_Analizy.Location = New System.Drawing.Point(112, 23)
        Me.lblNazwaHandlowa_Analizy.Name = "lblNazwaHandlowa_Analizy"
        Me.lblNazwaHandlowa_Analizy.Size = New System.Drawing.Size(930, 16)
        Me.lblNazwaHandlowa_Analizy.TabIndex = 109
        '
        'Label99
        '
        Me.Label99.ForeColor = System.Drawing.Color.Navy
        Me.Label99.Location = New System.Drawing.Point(8, 39)
        Me.Label99.Name = "Label99"
        Me.Label99.Size = New System.Drawing.Size(112, 16)
        Me.Label99.TabIndex = 110
        Me.Label99.Text = "Adres . . . . . . . . . . . . . . ."
        '
        'lblTOFI_Analizy
        '
        Me.lblTOFI_Analizy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTOFI_Analizy.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTOFI_Analizy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTOFI_Analizy.ForeColor = System.Drawing.Color.Purple
        Me.lblTOFI_Analizy.Location = New System.Drawing.Point(305, 7)
        Me.lblTOFI_Analizy.Name = "lblTOFI_Analizy"
        Me.lblTOFI_Analizy.Size = New System.Drawing.Size(737, 16)
        Me.lblTOFI_Analizy.TabIndex = 108
        '
        'lblIdent_Analizy
        '
        Me.lblIdent_Analizy.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent_Analizy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdent_Analizy.ForeColor = System.Drawing.Color.Purple
        Me.lblIdent_Analizy.Location = New System.Drawing.Point(112, 7)
        Me.lblIdent_Analizy.Name = "lblIdent_Analizy"
        Me.lblIdent_Analizy.Size = New System.Drawing.Size(112, 16)
        Me.lblIdent_Analizy.TabIndex = 105
        '
        'Panel13
        '
        Me.Panel13.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel13.Location = New System.Drawing.Point(0, 79)
        Me.Panel13.Name = "Panel13"
        Me.Panel13.Size = New System.Drawing.Size(1054, 1)
        Me.Panel13.TabIndex = 104
        '
        'Label102
        '
        Me.Label102.ForeColor = System.Drawing.Color.Navy
        Me.Label102.Location = New System.Drawing.Point(256, 7)
        Me.Label102.Name = "Label102"
        Me.Label102.Size = New System.Drawing.Size(112, 16)
        Me.Label102.TabIndex = 103
        Me.Label102.Text = "Nr TOFI . . . . . "
        '
        'Label103
        '
        Me.Label103.ForeColor = System.Drawing.Color.Navy
        Me.Label103.Location = New System.Drawing.Point(8, 23)
        Me.Label103.Name = "Label103"
        Me.Label103.Size = New System.Drawing.Size(112, 16)
        Me.Label103.TabIndex = 102
        Me.Label103.Text = "Nazwa klienta . . . . . . . . ."
        '
        'Label104
        '
        Me.Label104.ForeColor = System.Drawing.Color.Navy
        Me.Label104.Location = New System.Drawing.Point(8, 7)
        Me.Label104.Name = "Label104"
        Me.Label104.Size = New System.Drawing.Size(112, 16)
        Me.Label104.TabIndex = 101
        Me.Label104.Text = "Identyfikator . . . . . . . . ."
        '
        'cmdWszystkoOsoby
        '
        Me.cmdWszystkoOsoby.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystkoOsoby.Location = New System.Drawing.Point(618, 248)
        Me.cmdWszystkoOsoby.Name = "cmdWszystkoOsoby"
        Me.cmdWszystkoOsoby.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystkoOsoby.TabIndex = 206
        '
        'stbFiltryOsoby
        '
        Me.stbFiltryOsoby.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltryOsoby.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltryOsoby.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltryOsoby.Location = New System.Drawing.Point(11, 248)
        Me.stbFiltryOsoby.Name = "stbFiltryOsoby"
        Me.stbFiltryOsoby.ShowPanels = True
        Me.stbFiltryOsoby.Size = New System.Drawing.Size(1031, 22)
        Me.stbFiltryOsoby.TabIndex = 207
        Me.stbFiltryOsoby.Text = "StatusBar1"
        '
        'dagOsoby
        '
        Me.dagOsoby.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagOsoby.ContextMenuStrip = Me.MenuOsoby
        Me.dagOsoby.Location = New System.Drawing.Point(11, 274)
        Me.dagOsoby.Name = "dagOsoby"
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagOsoby.RowHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.dagOsoby.Size = New System.Drawing.Size(1031, 311)
        Me.dagOsoby.TabIndex = 205
        '
        'stbOsoby
        '
        Me.stbOsoby.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbOsoby.Dock = System.Windows.Forms.DockStyle.None
        Me.stbOsoby.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbOsoby.Location = New System.Drawing.Point(11, 585)
        Me.stbOsoby.Name = "stbOsoby"
        Me.stbOsoby.ShowPanels = True
        Me.stbOsoby.Size = New System.Drawing.Size(1031, 22)
        Me.stbOsoby.TabIndex = 204
        Me.stbOsoby.Text = "StatusBar1"
        '
        'cmdEdytujOsoby
        '
        Me.cmdEdytujOsoby.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdEdytujOsoby.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdEdytujOsoby.ForeColor = System.Drawing.Color.Black
        Me.cmdEdytujOsoby.Image = CType(resources.GetObject("cmdEdytujOsoby.Image"), System.Drawing.Image)
        Me.cmdEdytujOsoby.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdytujOsoby.Location = New System.Drawing.Point(962, 613)
        Me.cmdEdytujOsoby.Name = "cmdEdytujOsoby"
        Me.cmdEdytujOsoby.Size = New System.Drawing.Size(80, 23)
        Me.cmdEdytujOsoby.TabIndex = 203
        Me.cmdEdytujOsoby.Text = "Edytuj"
        Me.cmdEdytujOsoby.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdUsuñOsoby
        '
        Me.cmdUsuñOsoby.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUsuñOsoby.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdUsuñOsoby.ForeColor = System.Drawing.Color.Black
        Me.cmdUsuñOsoby.Image = CType(resources.GetObject("cmdUsuñOsoby.Image"), System.Drawing.Image)
        Me.cmdUsuñOsoby.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUsuñOsoby.Location = New System.Drawing.Point(874, 613)
        Me.cmdUsuñOsoby.Name = "cmdUsuñOsoby"
        Me.cmdUsuñOsoby.Size = New System.Drawing.Size(80, 23)
        Me.cmdUsuñOsoby.TabIndex = 202
        Me.cmdUsuñOsoby.Text = "Usuñ"
        Me.cmdUsuñOsoby.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdDodajOsoby
        '
        Me.cmdDodajOsoby.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDodajOsoby.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdDodajOsoby.ForeColor = System.Drawing.Color.Black
        Me.cmdDodajOsoby.Image = CType(resources.GetObject("cmdDodajOsoby.Image"), System.Drawing.Image)
        Me.cmdDodajOsoby.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDodajOsoby.Location = New System.Drawing.Point(785, 613)
        Me.cmdDodajOsoby.Name = "cmdDodajOsoby"
        Me.cmdDodajOsoby.Size = New System.Drawing.Size(80, 23)
        Me.cmdDodajOsoby.TabIndex = 201
        Me.cmdDodajOsoby.Text = "Dodaj"
        Me.cmdDodajOsoby.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'UslugiDod
        '
        Me.UslugiDod.BackColor = System.Drawing.SystemColors.Control
        Me.UslugiDod.Controls.Add(Me.PanelUslugiDod)
        Me.UslugiDod.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.UslugiDod.Location = New System.Drawing.Point(4, 22)
        Me.UslugiDod.Name = "UslugiDod"
        Me.UslugiDod.Size = New System.Drawing.Size(1054, 645)
        Me.UslugiDod.TabIndex = 5
        Me.UslugiDod.Text = "Rozszerzony zakres us³ug"
        '
        'PanelUslugiDod
        '
        Me.PanelUslugiDod.Controls.Add(Me.lblTelefonKom_UDod)
        Me.PanelUslugiDod.Controls.Add(Me.Label105)
        Me.PanelUslugiDod.Controls.Add(Me.lblTelefon_UDod)
        Me.PanelUslugiDod.Controls.Add(Me.Label111)
        Me.PanelUslugiDod.Controls.Add(Me.lblStanowisko_UDod)
        Me.PanelUslugiDod.Controls.Add(Me.Label113)
        Me.PanelUslugiDod.Controls.Add(Me.txtWartGran)
        Me.PanelUslugiDod.Controls.Add(Me.Label153)
        Me.PanelUslugiDod.Controls.Add(Me.dagRZUKontakty)
        Me.PanelUslugiDod.Controls.Add(Me.stbRZUKontakty)
        Me.PanelUslugiDod.Controls.Add(Me.cmdFi)
        Me.PanelUslugiDod.Controls.Add(Me.cmdWszystkoRZUKontakty)
        Me.PanelUslugiDod.Controls.Add(Me.stbFiltryRZUKontakty)
        Me.PanelUslugiDod.Controls.Add(Me.dtpDataOb)
        Me.PanelUslugiDod.Controls.Add(Me.Label109)
        Me.PanelUslugiDod.Controls.Add(Me.Label71)
        Me.PanelUslugiDod.Controls.Add(Me.cmdPokazR9)
        Me.PanelUslugiDod.Controls.Add(Me.cmdWybierzR9)
        Me.PanelUslugiDod.Controls.Add(Me.cmdPokazR8)
        Me.PanelUslugiDod.Controls.Add(Me.cmdWybierzR8)
        Me.PanelUslugiDod.Controls.Add(Me.cmdPokazR7)
        Me.PanelUslugiDod.Controls.Add(Me.cmdWybierzR7)
        Me.PanelUslugiDod.Controls.Add(Me.cmdPokazR6)
        Me.PanelUslugiDod.Controls.Add(Me.cmdWybierzR6)
        Me.PanelUslugiDod.Controls.Add(Me.cmdPokazR5)
        Me.PanelUslugiDod.Controls.Add(Me.cmdWybierzR5)
        Me.PanelUslugiDod.Controls.Add(Me.cmdPokazR4)
        Me.PanelUslugiDod.Controls.Add(Me.cmdWybierzR4)
        Me.PanelUslugiDod.Controls.Add(Me.cmdPokazR3)
        Me.PanelUslugiDod.Controls.Add(Me.cmdWybierzR3)
        Me.PanelUslugiDod.Controls.Add(Me.cmdPokazR2)
        Me.PanelUslugiDod.Controls.Add(Me.cmdWybierzR2)
        Me.PanelUslugiDod.Controls.Add(Me.cmdPokazR1)
        Me.PanelUslugiDod.Controls.Add(Me.cmdWybierzR1)
        Me.PanelUslugiDod.Controls.Add(Me.txtRaport09)
        Me.PanelUslugiDod.Controls.Add(Me.txtRaport08)
        Me.PanelUslugiDod.Controls.Add(Me.txtRaport07)
        Me.PanelUslugiDod.Controls.Add(Me.txtRaport06)
        Me.PanelUslugiDod.Controls.Add(Me.txtRaport05)
        Me.PanelUslugiDod.Controls.Add(Me.txtRaport04)
        Me.PanelUslugiDod.Controls.Add(Me.txtRaport03)
        Me.PanelUslugiDod.Controls.Add(Me.txtRaport02)
        Me.PanelUslugiDod.Controls.Add(Me.txtRaport01)
        Me.PanelUslugiDod.Controls.Add(Me.txtzakupyPopRok)
        Me.PanelUslugiDod.Controls.Add(Me.Label70)
        Me.PanelUslugiDod.Controls.Add(Me.dtpU9)
        Me.PanelUslugiDod.Controls.Add(Me.dtpU8)
        Me.PanelUslugiDod.Controls.Add(Me.dtpU6)
        Me.PanelUslugiDod.Controls.Add(Me.dtpU7)
        Me.PanelUslugiDod.Controls.Add(Me.dtpU5)
        Me.PanelUslugiDod.Controls.Add(Me.dtpU4)
        Me.PanelUslugiDod.Controls.Add(Me.dtpU3)
        Me.PanelUslugiDod.Controls.Add(Me.dtpU2)
        Me.PanelUslugiDod.Controls.Add(Me.dtpU1)
        Me.PanelUslugiDod.Controls.Add(Me.Label61)
        Me.PanelUslugiDod.Controls.Add(Me.chkU9)
        Me.PanelUslugiDod.Controls.Add(Me.chkU8)
        Me.PanelUslugiDod.Controls.Add(Me.chkU7)
        Me.PanelUslugiDod.Controls.Add(Me.chkU6)
        Me.PanelUslugiDod.Controls.Add(Me.chkU5)
        Me.PanelUslugiDod.Controls.Add(Me.chkU4)
        Me.PanelUslugiDod.Controls.Add(Me.chkU3)
        Me.PanelUslugiDod.Controls.Add(Me.chkU2)
        Me.PanelUslugiDod.Controls.Add(Me.chkU1)
        Me.PanelUslugiDod.Controls.Add(Me.Label25)
        Me.PanelUslugiDod.Controls.Add(Me.Label24)
        Me.PanelUslugiDod.Controls.Add(Me.lblNazwa2_UDod)
        Me.PanelUslugiDod.Controls.Add(Me.lblNazwa1_UDod)
        Me.PanelUslugiDod.Controls.Add(Me.Label57)
        Me.PanelUslugiDod.Controls.Add(Me.lblTofi_UDod)
        Me.PanelUslugiDod.Controls.Add(Me.lblNazwaHandlowa_UDod)
        Me.PanelUslugiDod.Controls.Add(Me.Panel2)
        Me.PanelUslugiDod.Controls.Add(Me.lblIdent_UDod)
        Me.PanelUslugiDod.Controls.Add(Me.Label72)
        Me.PanelUslugiDod.Controls.Add(Me.Label73)
        Me.PanelUslugiDod.Controls.Add(Me.Label94)
        Me.PanelUslugiDod.Controls.Add(Me.Label98)
        Me.PanelUslugiDod.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelUslugiDod.Location = New System.Drawing.Point(0, 0)
        Me.PanelUslugiDod.Name = "PanelUslugiDod"
        Me.PanelUslugiDod.Size = New System.Drawing.Size(1054, 645)
        Me.PanelUslugiDod.TabIndex = 359
        '
        'lblTelefonKom_UDod
        '
        Me.lblTelefonKom_UDod.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefonKom_UDod.Location = New System.Drawing.Point(940, 56)
        Me.lblTelefonKom_UDod.Name = "lblTelefonKom_UDod"
        Me.lblTelefonKom_UDod.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefonKom_UDod.TabIndex = 248
        '
        'Label105
        '
        Me.Label105.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label105.ForeColor = System.Drawing.Color.Navy
        Me.Label105.Location = New System.Drawing.Point(891, 56)
        Me.Label105.Name = "Label105"
        Me.Label105.Size = New System.Drawing.Size(112, 16)
        Me.Label105.TabIndex = 249
        Me.Label105.Text = "Tel. kom. . . . . . . . . . . . . . . "
        '
        'lblTelefon_UDod
        '
        Me.lblTelefon_UDod.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefon_UDod.Location = New System.Drawing.Point(673, 56)
        Me.lblTelefon_UDod.Name = "lblTelefon_UDod"
        Me.lblTelefon_UDod.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefon_UDod.TabIndex = 246
        '
        'Label111
        '
        Me.Label111.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label111.ForeColor = System.Drawing.Color.Navy
        Me.Label111.Location = New System.Drawing.Point(624, 56)
        Me.Label111.Name = "Label111"
        Me.Label111.Size = New System.Drawing.Size(112, 16)
        Me.Label111.TabIndex = 247
        Me.Label111.Text = "Telefon . . . . . . . . . . . . . . "
        '
        'lblStanowisko_UDod
        '
        Me.lblStanowisko_UDod.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStanowisko_UDod.Location = New System.Drawing.Point(404, 56)
        Me.lblStanowisko_UDod.Name = "lblStanowisko_UDod"
        Me.lblStanowisko_UDod.Size = New System.Drawing.Size(200, 16)
        Me.lblStanowisko_UDod.TabIndex = 244
        '
        'Label113
        '
        Me.Label113.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label113.ForeColor = System.Drawing.Color.Navy
        Me.Label113.Location = New System.Drawing.Point(333, 56)
        Me.Label113.Name = "Label113"
        Me.Label113.Size = New System.Drawing.Size(112, 16)
        Me.Label113.TabIndex = 245
        Me.Label113.Text = "Stanowisko . . . . . . . . . . . . . . "
        '
        'txtWartGran
        '
        Me.txtWartGran.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtWartGran.ForeColor = System.Drawing.Color.Purple
        Me.txtWartGran.Location = New System.Drawing.Point(664, 89)
        Me.txtWartGran.Name = "txtWartGran"
        Me.txtWartGran.ReadOnly = True
        Me.txtWartGran.Size = New System.Drawing.Size(60, 20)
        Me.txtWartGran.TabIndex = 222
        Me.txtWartGran.Text = "1"
        '
        'Label153
        '
        Me.Label153.ForeColor = System.Drawing.Color.Navy
        Me.Label153.Location = New System.Drawing.Point(478, 93)
        Me.Label153.Name = "Label153"
        Me.Label153.Size = New System.Drawing.Size(271, 16)
        Me.Label153.TabIndex = 223
        Me.Label153.Text = "Wartoœæ Graniczna za poprzedni rok . . . . . . . . ."
        '
        'dagRZUKontakty
        '
        Me.dagRZUKontakty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagRZUKontakty.ContextMenuStrip = Me.MenuRZUKontakty
        Me.dagRZUKontakty.Location = New System.Drawing.Point(0, 387)
        Me.dagRZUKontakty.Name = "dagRZUKontakty"
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagRZUKontakty.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dagRZUKontakty.Size = New System.Drawing.Size(1054, 229)
        Me.dagRZUKontakty.TabIndex = 221
        '
        'stbRZUKontakty
        '
        Me.stbRZUKontakty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbRZUKontakty.Dock = System.Windows.Forms.DockStyle.None
        Me.stbRZUKontakty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbRZUKontakty.Location = New System.Drawing.Point(0, 620)
        Me.stbRZUKontakty.Name = "stbRZUKontakty"
        Me.stbRZUKontakty.ShowPanels = True
        Me.stbRZUKontakty.Size = New System.Drawing.Size(1054, 22)
        Me.stbRZUKontakty.TabIndex = 220
        Me.stbRZUKontakty.Text = "StatusBar1"
        '
        'cmdFi
        '
        Me.cmdFi.Image = CType(resources.GetObject("cmdFi.Image"), System.Drawing.Image)
        Me.cmdFi.Location = New System.Drawing.Point(586, 365)
        Me.cmdFi.Name = "cmdFi"
        Me.cmdFi.Size = New System.Drawing.Size(16, 20)
        Me.cmdFi.TabIndex = 219
        Me.cmdFi.UseVisualStyleBackColor = True
        Me.cmdFi.Visible = False
        '
        'cmdWszystkoRZUKontakty
        '
        Me.cmdWszystkoRZUKontakty.Image = Global.KartotekaKlientow.My.Resources.Resources.close_16
        Me.cmdWszystkoRZUKontakty.Location = New System.Drawing.Point(607, 363)
        Me.cmdWszystkoRZUKontakty.Name = "cmdWszystkoRZUKontakty"
        Me.cmdWszystkoRZUKontakty.Size = New System.Drawing.Size(26, 22)
        Me.cmdWszystkoRZUKontakty.TabIndex = 217
        '
        'stbFiltryRZUKontakty
        '
        Me.stbFiltryRZUKontakty.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.stbFiltryRZUKontakty.Dock = System.Windows.Forms.DockStyle.None
        Me.stbFiltryRZUKontakty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.stbFiltryRZUKontakty.Location = New System.Drawing.Point(0, 363)
        Me.stbFiltryRZUKontakty.Name = "stbFiltryRZUKontakty"
        Me.stbFiltryRZUKontakty.ShowPanels = True
        Me.stbFiltryRZUKontakty.Size = New System.Drawing.Size(1054, 22)
        Me.stbFiltryRZUKontakty.TabIndex = 218
        Me.stbFiltryRZUKontakty.Text = "StatusBar1"
        '
        'dtpDataOb
        '
        Me.dtpDataOb.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpDataOb.CustomFormat = "yyyy-MM-dd"
        Me.dtpDataOb.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDataOb.Location = New System.Drawing.Point(926, 90)
        Me.dtpDataOb.Name = "dtpDataOb"
        Me.dtpDataOb.Size = New System.Drawing.Size(100, 20)
        Me.dtpDataOb.TabIndex = 215
        Me.dtpDataOb.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'Label109
        '
        Me.Label109.ForeColor = System.Drawing.Color.Navy
        Me.Label109.Location = New System.Drawing.Point(783, 93)
        Me.Label109.Name = "Label109"
        Me.Label109.Size = New System.Drawing.Size(147, 16)
        Me.Label109.TabIndex = 216
        Me.Label109.Text = "Data obowi¹zywania RZU . . . . . . . ."
        '
        'Label71
        '
        Me.Label71.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label71.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label71.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label71.ForeColor = System.Drawing.Color.White
        Me.Label71.Location = New System.Drawing.Point(0, 344)
        Me.Label71.Name = "Label71"
        Me.Label71.Size = New System.Drawing.Size(1054, 16)
        Me.Label71.TabIndex = 214
        Me.Label71.Text = "Kontakty handlowe"
        Me.Label71.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtRaport09
        '
        Me.txtRaport09.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRaport09.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtRaport09.ForeColor = System.Drawing.Color.Purple
        Me.txtRaport09.Location = New System.Drawing.Point(367, 319)
        Me.txtRaport09.Name = "txtRaport09"
        Me.txtRaport09.Size = New System.Drawing.Size(597, 20)
        Me.txtRaport09.TabIndex = 194
        Me.txtRaport09.Visible = False
        '
        'txtRaport08
        '
        Me.txtRaport08.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRaport08.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtRaport08.ForeColor = System.Drawing.Color.Purple
        Me.txtRaport08.Location = New System.Drawing.Point(367, 297)
        Me.txtRaport08.Name = "txtRaport08"
        Me.txtRaport08.Size = New System.Drawing.Size(597, 20)
        Me.txtRaport08.TabIndex = 193
        Me.txtRaport08.Visible = False
        '
        'txtRaport07
        '
        Me.txtRaport07.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRaport07.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtRaport07.ForeColor = System.Drawing.Color.Purple
        Me.txtRaport07.Location = New System.Drawing.Point(367, 275)
        Me.txtRaport07.Name = "txtRaport07"
        Me.txtRaport07.Size = New System.Drawing.Size(597, 20)
        Me.txtRaport07.TabIndex = 192
        '
        'txtRaport06
        '
        Me.txtRaport06.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRaport06.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtRaport06.ForeColor = System.Drawing.Color.Purple
        Me.txtRaport06.Location = New System.Drawing.Point(367, 253)
        Me.txtRaport06.Name = "txtRaport06"
        Me.txtRaport06.Size = New System.Drawing.Size(597, 20)
        Me.txtRaport06.TabIndex = 191
        '
        'txtRaport05
        '
        Me.txtRaport05.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRaport05.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtRaport05.ForeColor = System.Drawing.Color.Purple
        Me.txtRaport05.Location = New System.Drawing.Point(367, 231)
        Me.txtRaport05.Name = "txtRaport05"
        Me.txtRaport05.Size = New System.Drawing.Size(597, 20)
        Me.txtRaport05.TabIndex = 190
        '
        'txtRaport04
        '
        Me.txtRaport04.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRaport04.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtRaport04.ForeColor = System.Drawing.Color.Purple
        Me.txtRaport04.Location = New System.Drawing.Point(367, 209)
        Me.txtRaport04.Name = "txtRaport04"
        Me.txtRaport04.Size = New System.Drawing.Size(597, 20)
        Me.txtRaport04.TabIndex = 189
        '
        'txtRaport03
        '
        Me.txtRaport03.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRaport03.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtRaport03.ForeColor = System.Drawing.Color.Purple
        Me.txtRaport03.Location = New System.Drawing.Point(367, 187)
        Me.txtRaport03.Name = "txtRaport03"
        Me.txtRaport03.Size = New System.Drawing.Size(597, 20)
        Me.txtRaport03.TabIndex = 188
        '
        'txtRaport02
        '
        Me.txtRaport02.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRaport02.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtRaport02.ForeColor = System.Drawing.Color.Purple
        Me.txtRaport02.Location = New System.Drawing.Point(367, 165)
        Me.txtRaport02.Name = "txtRaport02"
        Me.txtRaport02.Size = New System.Drawing.Size(597, 20)
        Me.txtRaport02.TabIndex = 187
        '
        'txtRaport01
        '
        Me.txtRaport01.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRaport01.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtRaport01.ForeColor = System.Drawing.Color.Purple
        Me.txtRaport01.Location = New System.Drawing.Point(367, 143)
        Me.txtRaport01.Name = "txtRaport01"
        Me.txtRaport01.Size = New System.Drawing.Size(597, 20)
        Me.txtRaport01.TabIndex = 186
        '
        'txtzakupyPopRok
        '
        Me.txtzakupyPopRok.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtzakupyPopRok.ForeColor = System.Drawing.Color.Purple
        Me.txtzakupyPopRok.Location = New System.Drawing.Point(367, 90)
        Me.txtzakupyPopRok.Name = "txtzakupyPopRok"
        Me.txtzakupyPopRok.Size = New System.Drawing.Size(60, 20)
        Me.txtzakupyPopRok.TabIndex = 111
        Me.txtzakupyPopRok.Text = "1"
        '
        'Label70
        '
        Me.Label70.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label70.ForeColor = System.Drawing.Color.Navy
        Me.Label70.Location = New System.Drawing.Point(367, 128)
        Me.Label70.Name = "Label70"
        Me.Label70.Size = New System.Drawing.Size(416, 16)
        Me.Label70.TabIndex = 185
        Me.Label70.Text = "Plik z Raportem realizacji."
        '
        'dtpU9
        '
        Me.dtpU9.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpU9.CustomFormat = "yyyy-MM-dd"
        Me.dtpU9.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpU9.Location = New System.Drawing.Point(260, 318)
        Me.dtpU9.Name = "dtpU9"
        Me.dtpU9.Size = New System.Drawing.Size(100, 20)
        Me.dtpU9.TabIndex = 184
        Me.dtpU9.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        Me.dtpU9.Visible = False
        '
        'dtpU8
        '
        Me.dtpU8.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpU8.CustomFormat = "yyyy-MM-dd"
        Me.dtpU8.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpU8.Location = New System.Drawing.Point(260, 296)
        Me.dtpU8.Name = "dtpU8"
        Me.dtpU8.Size = New System.Drawing.Size(100, 20)
        Me.dtpU8.TabIndex = 183
        Me.dtpU8.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        Me.dtpU8.Visible = False
        '
        'dtpU6
        '
        Me.dtpU6.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpU6.CustomFormat = "yyyy-MM-dd"
        Me.dtpU6.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpU6.Location = New System.Drawing.Point(260, 252)
        Me.dtpU6.Name = "dtpU6"
        Me.dtpU6.Size = New System.Drawing.Size(100, 20)
        Me.dtpU6.TabIndex = 181
        Me.dtpU6.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'dtpU7
        '
        Me.dtpU7.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpU7.CustomFormat = "yyyy-MM-dd"
        Me.dtpU7.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpU7.Location = New System.Drawing.Point(260, 275)
        Me.dtpU7.Name = "dtpU7"
        Me.dtpU7.Size = New System.Drawing.Size(100, 20)
        Me.dtpU7.TabIndex = 182
        Me.dtpU7.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'dtpU5
        '
        Me.dtpU5.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpU5.CustomFormat = "yyyy-MM-dd"
        Me.dtpU5.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpU5.Location = New System.Drawing.Point(260, 231)
        Me.dtpU5.Name = "dtpU5"
        Me.dtpU5.Size = New System.Drawing.Size(100, 20)
        Me.dtpU5.TabIndex = 180
        Me.dtpU5.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'dtpU4
        '
        Me.dtpU4.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpU4.CustomFormat = "yyyy-MM-dd"
        Me.dtpU4.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpU4.Location = New System.Drawing.Point(260, 209)
        Me.dtpU4.Name = "dtpU4"
        Me.dtpU4.Size = New System.Drawing.Size(100, 20)
        Me.dtpU4.TabIndex = 179
        Me.dtpU4.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'dtpU3
        '
        Me.dtpU3.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpU3.CustomFormat = "yyyy-MM-dd"
        Me.dtpU3.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpU3.Location = New System.Drawing.Point(260, 186)
        Me.dtpU3.Name = "dtpU3"
        Me.dtpU3.Size = New System.Drawing.Size(100, 20)
        Me.dtpU3.TabIndex = 178
        Me.dtpU3.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'dtpU2
        '
        Me.dtpU2.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpU2.CustomFormat = "yyyy-MM-dd"
        Me.dtpU2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpU2.Location = New System.Drawing.Point(260, 164)
        Me.dtpU2.Name = "dtpU2"
        Me.dtpU2.Size = New System.Drawing.Size(100, 20)
        Me.dtpU2.TabIndex = 177
        Me.dtpU2.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'dtpU1
        '
        Me.dtpU1.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpU1.CustomFormat = "yyyy-MM-dd"
        Me.dtpU1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpU1.Location = New System.Drawing.Point(260, 142)
        Me.dtpU1.Name = "dtpU1"
        Me.dtpU1.Size = New System.Drawing.Size(100, 20)
        Me.dtpU1.TabIndex = 176
        Me.dtpU1.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'Label61
        '
        Me.Label61.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label61.ForeColor = System.Drawing.Color.Navy
        Me.Label61.Location = New System.Drawing.Point(260, 113)
        Me.Label61.Name = "Label61"
        Me.Label61.Size = New System.Drawing.Size(85, 31)
        Me.Label61.TabIndex = 123
        Me.Label61.Text = "Uzgodniony termin realizacji :"
        '
        'chkU9
        '
        Me.chkU9.ForeColor = System.Drawing.Color.Navy
        Me.chkU9.Location = New System.Drawing.Point(11, 323)
        Me.chkU9.Name = "chkU9"
        Me.chkU9.Size = New System.Drawing.Size(300, 18)
        Me.chkU9.TabIndex = 122
        Me.chkU9.Text = " . . . . . . . . . . . . . . . . . . . . ." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.chkU9.Visible = False
        '
        'chkU8
        '
        Me.chkU8.ForeColor = System.Drawing.Color.Navy
        Me.chkU8.Location = New System.Drawing.Point(11, 301)
        Me.chkU8.Name = "chkU8"
        Me.chkU8.Size = New System.Drawing.Size(300, 18)
        Me.chkU8.TabIndex = 121
        Me.chkU8.Text = " . . . . . . . . . . . . . " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.chkU8.Visible = False
        '
        'chkU7
        '
        Me.chkU7.ForeColor = System.Drawing.Color.Navy
        Me.chkU7.Location = New System.Drawing.Point(11, 279)
        Me.chkU7.Name = "chkU7"
        Me.chkU7.Size = New System.Drawing.Size(300, 18)
        Me.chkU7.TabIndex = 120
        Me.chkU7.Text = "NAPRAWA DRUKAREK LASEROWYCH . . . . . . . . . . . . . . . ."
        '
        'chkU6
        '
        Me.chkU6.ForeColor = System.Drawing.Color.Navy
        Me.chkU6.Location = New System.Drawing.Point(11, 257)
        Me.chkU6.Name = "chkU6"
        Me.chkU6.Size = New System.Drawing.Size(300, 18)
        Me.chkU6.TabIndex = 119
        Me.chkU6.Text = "CZYSZCZENIE DRUKAREK LASEROWYCH . . . . . . . . . . . . . . . . . . . . . ."
        '
        'chkU5
        '
        Me.chkU5.ForeColor = System.Drawing.Color.Navy
        Me.chkU5.Location = New System.Drawing.Point(11, 235)
        Me.chkU5.Name = "chkU5"
        Me.chkU5.Size = New System.Drawing.Size(300, 18)
        Me.chkU5.TabIndex = 118
        Me.chkU5.Text = "PRZEGL¥D TECHNICZNY . . . . . . . . . . . . . . . . . . . ." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'chkU4
        '
        Me.chkU4.ForeColor = System.Drawing.Color.Navy
        Me.chkU4.Location = New System.Drawing.Point(11, 213)
        Me.chkU4.Name = "chkU4"
        Me.chkU4.Size = New System.Drawing.Size(300, 18)
        Me.chkU4.TabIndex = 117
        Me.chkU4.Text = "DOBOR URZ¥DZENIA DRUKUJ¥CEGO . . . . . . . . . . . . . "
        '
        'chkU3
        '
        Me.chkU3.ForeColor = System.Drawing.Color.Navy
        Me.chkU3.Location = New System.Drawing.Point(11, 191)
        Me.chkU3.Name = "chkU3"
        Me.chkU3.Size = New System.Drawing.Size(300, 18)
        Me.chkU3.TabIndex = 116
        Me.chkU3.Text = "AUDYT DRUKU . . . . . . . . . . . . . . . . . . . . . . . . . . . . . ."
        '
        'chkU2
        '
        Me.chkU2.ForeColor = System.Drawing.Color.Navy
        Me.chkU2.Location = New System.Drawing.Point(11, 169)
        Me.chkU2.Name = "chkU2"
        Me.chkU2.Size = New System.Drawing.Size(300, 18)
        Me.chkU2.TabIndex = 115
        Me.chkU2.Text = "WSPARCIE INFORMATYCZNE . . . . . . . . . . . . . . . . . . . . . . . . ." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'chkU1
        '
        Me.chkU1.ForeColor = System.Drawing.Color.Navy
        Me.chkU1.Location = New System.Drawing.Point(11, 147)
        Me.chkU1.Name = "chkU1"
        Me.chkU1.Size = New System.Drawing.Size(300, 18)
        Me.chkU1.TabIndex = 114
        Me.chkU1.Text = "PLATFORMA DRUKU . . . . . . . . . . . . . . . . . . . . . . . . . . "
        '
        'Label25
        '
        Me.Label25.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label25.ForeColor = System.Drawing.Color.Navy
        Me.Label25.Location = New System.Drawing.Point(8, 128)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(360, 16)
        Me.Label25.TabIndex = 113
        Me.Label25.Text = "Dostêpne Us³ugi Dodatkowe :"
        '
        'Label24
        '
        Me.Label24.ForeColor = System.Drawing.Color.Navy
        Me.Label24.Location = New System.Drawing.Point(8, 93)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(397, 16)
        Me.Label24.TabIndex = 112
        Me.Label24.Text = "Sumaryczna wartoœæ zakupów w firmie PRYZMAT w roku poprzednim . . . . . . . . ."
        '
        'lblNazwa2_UDod
        '
        Me.lblNazwa2_UDod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa2_UDod.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa2_UDod.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa2_UDod.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa2_UDod.Location = New System.Drawing.Point(112, 40)
        Me.lblNazwa2_UDod.Name = "lblNazwa2_UDod"
        Me.lblNazwa2_UDod.Size = New System.Drawing.Size(928, 16)
        Me.lblNazwa2_UDod.TabIndex = 105
        '
        'lblNazwa1_UDod
        '
        Me.lblNazwa1_UDod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa1_UDod.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa1_UDod.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa1_UDod.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa1_UDod.Location = New System.Drawing.Point(112, 24)
        Me.lblNazwa1_UDod.Name = "lblNazwa1_UDod"
        Me.lblNazwa1_UDod.Size = New System.Drawing.Size(928, 16)
        Me.lblNazwa1_UDod.TabIndex = 108
        '
        'Label57
        '
        Me.Label57.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label57.ForeColor = System.Drawing.Color.Navy
        Me.Label57.Location = New System.Drawing.Point(8, 40)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(112, 16)
        Me.Label57.TabIndex = 109
        Me.Label57.Text = "Adres . . . . . . . . . . . . . . "
        '
        'lblTofi_UDod
        '
        Me.lblTofi_UDod.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTofi_UDod.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTofi_UDod.ForeColor = System.Drawing.Color.Purple
        Me.lblTofi_UDod.Location = New System.Drawing.Point(305, 8)
        Me.lblTofi_UDod.Name = "lblTofi_UDod"
        Me.lblTofi_UDod.Size = New System.Drawing.Size(735, 16)
        Me.lblTofi_UDod.TabIndex = 107
        '
        'lblNazwaHandlowa_UDod
        '
        Me.lblNazwaHandlowa_UDod.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwaHandlowa_UDod.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwaHandlowa_UDod.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwaHandlowa_UDod.Location = New System.Drawing.Point(112, 56)
        Me.lblNazwaHandlowa_UDod.Name = "lblNazwaHandlowa_UDod"
        Me.lblNazwaHandlowa_UDod.Size = New System.Drawing.Size(200, 16)
        Me.lblNazwaHandlowa_UDod.TabIndex = 106
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(0, 80)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1054, 1)
        Me.Panel2.TabIndex = 103
        '
        'lblIdent_UDod
        '
        Me.lblIdent_UDod.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent_UDod.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdent_UDod.ForeColor = System.Drawing.Color.Purple
        Me.lblIdent_UDod.Location = New System.Drawing.Point(112, 8)
        Me.lblIdent_UDod.Name = "lblIdent_UDod"
        Me.lblIdent_UDod.Size = New System.Drawing.Size(112, 16)
        Me.lblIdent_UDod.TabIndex = 104
        '
        'Label72
        '
        Me.Label72.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label72.ForeColor = System.Drawing.Color.Navy
        Me.Label72.Location = New System.Drawing.Point(256, 8)
        Me.Label72.Name = "Label72"
        Me.Label72.Size = New System.Drawing.Size(112, 16)
        Me.Label72.TabIndex = 102
        Me.Label72.Text = "Nr TOFI . . . . . "
        '
        'Label73
        '
        Me.Label73.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label73.ForeColor = System.Drawing.Color.Navy
        Me.Label73.Location = New System.Drawing.Point(8, 24)
        Me.Label73.Name = "Label73"
        Me.Label73.Size = New System.Drawing.Size(112, 16)
        Me.Label73.TabIndex = 101
        Me.Label73.Text = "Nazwa klienta . . . . . . . . ."
        '
        'Label94
        '
        Me.Label94.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label94.ForeColor = System.Drawing.Color.Navy
        Me.Label94.Location = New System.Drawing.Point(8, 8)
        Me.Label94.Name = "Label94"
        Me.Label94.Size = New System.Drawing.Size(112, 16)
        Me.Label94.TabIndex = 100
        Me.Label94.Text = "Identyfikator . . . . . . . . ."
        '
        'Label98
        '
        Me.Label98.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label98.ForeColor = System.Drawing.Color.Navy
        Me.Label98.Location = New System.Drawing.Point(8, 56)
        Me.Label98.Name = "Label98"
        Me.Label98.Size = New System.Drawing.Size(112, 16)
        Me.Label98.TabIndex = 110
        Me.Label98.Text = "Osoba kontaktowa . . . . . . . . . . . . . . "
        '
        'Charakterystyka
        '
        Me.Charakterystyka.BackColor = System.Drawing.SystemColors.Control
        Me.Charakterystyka.Controls.Add(Me.cmdWybierzPodbranze)
        Me.Charakterystyka.Controls.Add(Me.cmdWybierzBranze)
        Me.Charakterystyka.Controls.Add(Me.txtLiczbaPracZdalnie)
        Me.Charakterystyka.Controls.Add(Me.Label78)
        Me.Charakterystyka.Controls.Add(Me.txtLiczbaZatrudnionych)
        Me.Charakterystyka.Controls.Add(Me.Label79)
        Me.Charakterystyka.Controls.Add(Me.txtPodbranza)
        Me.Charakterystyka.Controls.Add(Me.Label67)
        Me.Charakterystyka.Controls.Add(Me.txtBranza)
        Me.Charakterystyka.Controls.Add(Me.Label75)
        Me.Charakterystyka.Controls.Add(Me.chbRodzT)
        Me.Charakterystyka.Controls.Add(Me.chbRodzA)
        Me.Charakterystyka.Controls.Add(Me.chbRodzL)
        Me.Charakterystyka.Controls.Add(Me.dtpDataWeryfikKrytSegmentacji)
        Me.Charakterystyka.Controls.Add(Me.Label66)
        Me.Charakterystyka.Controls.Add(Me.chkZaintKorzystaZNajmu)
        Me.Charakterystyka.Controls.Add(Me.chkZaintRozliczaStrony)
        Me.Charakterystyka.Controls.Add(Me.chkAktKorzystaZNajmu)
        Me.Charakterystyka.Controls.Add(Me.chkAktRozliczaStrony)
        Me.Charakterystyka.Controls.Add(Me.chkAktZakupMatEksp)
        Me.Charakterystyka.Controls.Add(Me.Label63)
        Me.Charakterystyka.Controls.Add(Me.Panel14)
        Me.Charakterystyka.Controls.Add(Me.Label58)
        Me.Charakterystyka.Controls.Add(Me.Panel10)
        Me.Charakterystyka.Controls.Add(Me.Label53)
        Me.Charakterystyka.Controls.Add(Me.Label47)
        Me.Charakterystyka.Controls.Add(Me.Label42)
        Me.Charakterystyka.Controls.Add(Me.Panel4)
        Me.Charakterystyka.Controls.Add(Me.Panel5)
        Me.Charakterystyka.Controls.Add(Me.Label40)
        Me.Charakterystyka.Controls.Add(Me.cbxSposobWyboruDostawcy)
        Me.Charakterystyka.Controls.Add(Me.Label38)
        Me.Charakterystyka.Controls.Add(Me.Label37)
        Me.Charakterystyka.Controls.Add(Me.lblPotencjalDruku)
        Me.Charakterystyka.Controls.Add(Me.Label36)
        Me.Charakterystyka.Controls.Add(Me.Label35)
        Me.Charakterystyka.Controls.Add(Me.txtLiczbaWydrukow)
        Me.Charakterystyka.Controls.Add(Me.Label29)
        Me.Charakterystyka.Controls.Add(Me.txtLiczbaUrzadzen)
        Me.Charakterystyka.Controls.Add(Me.Label28)
        Me.Charakterystyka.Controls.Add(Me.Label20)
        Me.Charakterystyka.Controls.Add(Me.chkZainstalowanyMonitor)
        Me.Charakterystyka.Controls.Add(Me.SplitContainer1)
        Me.Charakterystyka.Controls.Add(Me.lblTelefonKom_Charakterystyka)
        Me.Charakterystyka.Controls.Add(Me.Label64)
        Me.Charakterystyka.Controls.Add(Me.lblTelefon_Charakterystyka)
        Me.Charakterystyka.Controls.Add(Me.Label74)
        Me.Charakterystyka.Controls.Add(Me.lblStanowisko_Charakterystyka)
        Me.Charakterystyka.Controls.Add(Me.Label84)
        Me.Charakterystyka.Controls.Add(Me.txtWykazUrzadzen)
        Me.Charakterystyka.Controls.Add(Me.cmdPoka¿Plik)
        Me.Charakterystyka.Controls.Add(Me.cmdWybierzPlik)
        Me.Charakterystyka.Controls.Add(Me.lblOfertaPlik)
        Me.Charakterystyka.Controls.Add(Me.Label59)
        Me.Charakterystyka.Controls.Add(Me.txtTZlozenia)
        Me.Charakterystyka.Controls.Add(Me.Label45)
        Me.Charakterystyka.Controls.Add(Me.lblNazwa1_Charakterystyka)
        Me.Charakterystyka.Controls.Add(Me.Label55)
        Me.Charakterystyka.Controls.Add(Me.Label54)
        Me.Charakterystyka.Controls.Add(Me.Label52)
        Me.Charakterystyka.Controls.Add(Me.Label34)
        Me.Charakterystyka.Controls.Add(Me.Label27)
        Me.Charakterystyka.Controls.Add(Me.lblNazwaHandlowa_Charakterystyka)
        Me.Charakterystyka.Controls.Add(Me.Label19)
        Me.Charakterystyka.Controls.Add(Me.lblTofi_Charakterystyka)
        Me.Charakterystyka.Controls.Add(Me.lblNazwa2_Charakterystyka)
        Me.Charakterystyka.Controls.Add(Me.lblIdent_Charakterystyka)
        Me.Charakterystyka.Controls.Add(Me.Panel1)
        Me.Charakterystyka.Controls.Add(Me.Label15)
        Me.Charakterystyka.Controls.Add(Me.Label17)
        Me.Charakterystyka.Controls.Add(Me.Label18)
        Me.Charakterystyka.Controls.Add(Me.Label60)
        Me.Charakterystyka.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Charakterystyka.Location = New System.Drawing.Point(4, 22)
        Me.Charakterystyka.Name = "Charakterystyka"
        Me.Charakterystyka.Size = New System.Drawing.Size(1054, 645)
        Me.Charakterystyka.TabIndex = 4
        Me.Charakterystyka.Text = "Charakterystyka"
        '
        'cmdWybierzPodbranze
        '
        Me.cmdWybierzPodbranze.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzPodbranze.Image = CType(resources.GetObject("cmdWybierzPodbranze.Image"), System.Drawing.Image)
        Me.cmdWybierzPodbranze.Location = New System.Drawing.Point(360, 106)
        Me.cmdWybierzPodbranze.Name = "cmdWybierzPodbranze"
        Me.cmdWybierzPodbranze.Size = New System.Drawing.Size(32, 23)
        Me.cmdWybierzPodbranze.TabIndex = 284
        Me.cmdWybierzPodbranze.Text = "2"
        '
        'cmdWybierzBranze
        '
        Me.cmdWybierzBranze.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierzBranze.Image = CType(resources.GetObject("cmdWybierzBranze.Image"), System.Drawing.Image)
        Me.cmdWybierzBranze.Location = New System.Drawing.Point(360, 84)
        Me.cmdWybierzBranze.Name = "cmdWybierzBranze"
        Me.cmdWybierzBranze.Size = New System.Drawing.Size(32, 23)
        Me.cmdWybierzBranze.TabIndex = 283
        Me.cmdWybierzBranze.Text = "2"
        '
        'txtLiczbaPracZdalnie
        '
        Me.txtLiczbaPracZdalnie.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLiczbaPracZdalnie.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtLiczbaPracZdalnie.ForeColor = System.Drawing.Color.Purple
        Me.txtLiczbaPracZdalnie.Location = New System.Drawing.Point(550, 107)
        Me.txtLiczbaPracZdalnie.Name = "txtLiczbaPracZdalnie"
        Me.txtLiczbaPracZdalnie.Size = New System.Drawing.Size(70, 20)
        Me.txtLiczbaPracZdalnie.TabIndex = 282
        '
        'Label78
        '
        Me.Label78.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label78.ForeColor = System.Drawing.Color.Navy
        Me.Label78.Location = New System.Drawing.Point(411, 110)
        Me.Label78.Name = "Label78"
        Me.Label78.Size = New System.Drawing.Size(183, 16)
        Me.Label78.TabIndex = 281
        Me.Label78.Text = "Liczba pracuj¹cych zdalnie . . . . . . ."
        '
        'txtLiczbaZatrudnionych
        '
        Me.txtLiczbaZatrudnionych.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLiczbaZatrudnionych.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtLiczbaZatrudnionych.ForeColor = System.Drawing.Color.Purple
        Me.txtLiczbaZatrudnionych.Location = New System.Drawing.Point(550, 85)
        Me.txtLiczbaZatrudnionych.Name = "txtLiczbaZatrudnionych"
        Me.txtLiczbaZatrudnionych.Size = New System.Drawing.Size(70, 20)
        Me.txtLiczbaZatrudnionych.TabIndex = 280
        '
        'Label79
        '
        Me.Label79.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label79.ForeColor = System.Drawing.Color.Navy
        Me.Label79.Location = New System.Drawing.Point(411, 88)
        Me.Label79.Name = "Label79"
        Me.Label79.Size = New System.Drawing.Size(183, 16)
        Me.Label79.TabIndex = 279
        Me.Label79.Text = "Liczba zatrudnionych . . . . . . ."
        '
        'txtPodbranza
        '
        Me.txtPodbranza.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPodbranza.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtPodbranza.ForeColor = System.Drawing.Color.Purple
        Me.txtPodbranza.Location = New System.Drawing.Point(109, 107)
        Me.txtPodbranza.Name = "txtPodbranza"
        Me.txtPodbranza.Size = New System.Drawing.Size(283, 20)
        Me.txtPodbranza.TabIndex = 278
        '
        'Label67
        '
        Me.Label67.ForeColor = System.Drawing.Color.Navy
        Me.Label67.Location = New System.Drawing.Point(11, 110)
        Me.Label67.Name = "Label67"
        Me.Label67.Size = New System.Drawing.Size(183, 16)
        Me.Label67.TabIndex = 277
        Me.Label67.Text = "Podbran¿a . . . . . . ."
        '
        'txtBranza
        '
        Me.txtBranza.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBranza.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtBranza.ForeColor = System.Drawing.Color.Purple
        Me.txtBranza.Location = New System.Drawing.Point(109, 85)
        Me.txtBranza.Name = "txtBranza"
        Me.txtBranza.Size = New System.Drawing.Size(283, 20)
        Me.txtBranza.TabIndex = 276
        '
        'Label75
        '
        Me.Label75.ForeColor = System.Drawing.Color.Navy
        Me.Label75.Location = New System.Drawing.Point(11, 88)
        Me.Label75.Name = "Label75"
        Me.Label75.Size = New System.Drawing.Size(183, 16)
        Me.Label75.TabIndex = 275
        Me.Label75.Text = "Bran¿a . . . . . . . . . . ."
        '
        'chbRodzT
        '
        Me.chbRodzT.Location = New System.Drawing.Point(235, 138)
        Me.chbRodzT.Name = "chbRodzT"
        Me.chbRodzT.Size = New System.Drawing.Size(24, 16)
        Me.chbRodzT.TabIndex = 107
        '
        'chbRodzA
        '
        Me.chbRodzA.Location = New System.Drawing.Point(175, 138)
        Me.chbRodzA.Name = "chbRodzA"
        Me.chbRodzA.Size = New System.Drawing.Size(24, 16)
        Me.chbRodzA.TabIndex = 106
        '
        'chbRodzL
        '
        Me.chbRodzL.Location = New System.Drawing.Point(125, 138)
        Me.chbRodzL.Name = "chbRodzL"
        Me.chbRodzL.Size = New System.Drawing.Size(24, 16)
        Me.chbRodzL.TabIndex = 105
        '
        'dtpDataWeryfikKrytSegmentacji
        '
        Me.dtpDataWeryfikKrytSegmentacji.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDataWeryfikKrytSegmentacji.CalendarForeColor = System.Drawing.Color.Navy
        Me.dtpDataWeryfikKrytSegmentacji.CustomFormat = "yyyy-MM-dd"
        Me.dtpDataWeryfikKrytSegmentacji.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDataWeryfikKrytSegmentacji.Location = New System.Drawing.Point(850, 325)
        Me.dtpDataWeryfikKrytSegmentacji.Name = "dtpDataWeryfikKrytSegmentacji"
        Me.dtpDataWeryfikKrytSegmentacji.Size = New System.Drawing.Size(100, 20)
        Me.dtpDataWeryfikKrytSegmentacji.TabIndex = 274
        Me.dtpDataWeryfikKrytSegmentacji.Value = New Date(2014, 4, 8, 0, 0, 0, 0)
        '
        'Label66
        '
        Me.Label66.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label66.ForeColor = System.Drawing.Color.Navy
        Me.Label66.Location = New System.Drawing.Point(637, 329)
        Me.Label66.Name = "Label66"
        Me.Label66.Size = New System.Drawing.Size(207, 16)
        Me.Label66.TabIndex = 273
        Me.Label66.Text = "Data weryfikacji kryteriów segmentacji. . . . . . ."
        '
        'chkZaintKorzystaZNajmu
        '
        Me.chkZaintKorzystaZNajmu.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkZaintKorzystaZNajmu.Location = New System.Drawing.Point(909, 283)
        Me.chkZaintKorzystaZNajmu.Name = "chkZaintKorzystaZNajmu"
        Me.chkZaintKorzystaZNajmu.Size = New System.Drawing.Size(24, 16)
        Me.chkZaintKorzystaZNajmu.TabIndex = 272
        '
        'chkZaintRozliczaStrony
        '
        Me.chkZaintRozliczaStrony.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkZaintRozliczaStrony.Location = New System.Drawing.Point(829, 283)
        Me.chkZaintRozliczaStrony.Name = "chkZaintRozliczaStrony"
        Me.chkZaintRozliczaStrony.Size = New System.Drawing.Size(24, 16)
        Me.chkZaintRozliczaStrony.TabIndex = 271
        '
        'chkAktKorzystaZNajmu
        '
        Me.chkAktKorzystaZNajmu.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkAktKorzystaZNajmu.Location = New System.Drawing.Point(909, 262)
        Me.chkAktKorzystaZNajmu.Name = "chkAktKorzystaZNajmu"
        Me.chkAktKorzystaZNajmu.Size = New System.Drawing.Size(24, 16)
        Me.chkAktKorzystaZNajmu.TabIndex = 270
        '
        'chkAktRozliczaStrony
        '
        Me.chkAktRozliczaStrony.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkAktRozliczaStrony.Location = New System.Drawing.Point(829, 262)
        Me.chkAktRozliczaStrony.Name = "chkAktRozliczaStrony"
        Me.chkAktRozliczaStrony.Size = New System.Drawing.Size(24, 16)
        Me.chkAktRozliczaStrony.TabIndex = 269
        '
        'chkAktZakupMatEksp
        '
        Me.chkAktZakupMatEksp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkAktZakupMatEksp.Location = New System.Drawing.Point(756, 262)
        Me.chkAktZakupMatEksp.Name = "chkAktZakupMatEksp"
        Me.chkAktZakupMatEksp.Size = New System.Drawing.Size(24, 16)
        Me.chkAktZakupMatEksp.TabIndex = 268
        '
        'Label63
        '
        Me.Label63.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label63.ForeColor = System.Drawing.Color.Navy
        Me.Label63.Location = New System.Drawing.Point(637, 283)
        Me.Label63.Name = "Label63"
        Me.Label63.Size = New System.Drawing.Size(99, 16)
        Me.Label63.TabIndex = 267
        Me.Label63.Text = "Zainteresowanie . . . . . . . . ."
        '
        'Panel14
        '
        Me.Panel14.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel14.Location = New System.Drawing.Point(637, 279)
        Me.Panel14.Name = "Panel14"
        Me.Panel14.Size = New System.Drawing.Size(403, 1)
        Me.Panel14.TabIndex = 266
        '
        'Label58
        '
        Me.Label58.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label58.ForeColor = System.Drawing.Color.Navy
        Me.Label58.Location = New System.Drawing.Point(637, 262)
        Me.Label58.Name = "Label58"
        Me.Label58.Size = New System.Drawing.Size(99, 16)
        Me.Label58.TabIndex = 265
        Me.Label58.Text = "Stan obecny . . . . . . . . ."
        '
        'Panel10
        '
        Me.Panel10.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel10.Location = New System.Drawing.Point(876, 229)
        Me.Panel10.Name = "Panel10"
        Me.Panel10.Size = New System.Drawing.Size(1, 68)
        Me.Panel10.TabIndex = 264
        '
        'Label53
        '
        Me.Label53.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label53.ForeColor = System.Drawing.Color.Navy
        Me.Label53.Location = New System.Drawing.Point(886, 225)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(64, 30)
        Me.Label53.TabIndex = 263
        Me.Label53.Text = "Najem"
        Me.Label53.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label47
        '
        Me.Label47.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label47.ForeColor = System.Drawing.Color.Navy
        Me.Label47.Location = New System.Drawing.Point(806, 225)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(64, 30)
        Me.Label47.TabIndex = 262
        Me.Label47.Text = "Rozliczenie w stronach"
        Me.Label47.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label42
        '
        Me.Label42.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label42.ForeColor = System.Drawing.Color.Navy
        Me.Label42.Location = New System.Drawing.Point(726, 225)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(64, 30)
        Me.Label42.TabIndex = 261
        Me.Label42.Text = "Zakup mat. eksp."
        Me.Label42.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel4
        '
        Me.Panel4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Location = New System.Drawing.Point(796, 229)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(1, 68)
        Me.Panel4.TabIndex = 260
        '
        'Panel5
        '
        Me.Panel5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Location = New System.Drawing.Point(637, 258)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(403, 1)
        Me.Panel5.TabIndex = 259
        '
        'Label40
        '
        Me.Label40.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label40.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label40.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label40.ForeColor = System.Drawing.Color.White
        Me.Label40.Location = New System.Drawing.Point(637, 207)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(403, 16)
        Me.Label40.TabIndex = 258
        Me.Label40.Text = "Forma rozliczania druku"
        Me.Label40.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbxSposobWyboruDostawcy
        '
        Me.cbxSposobWyboruDostawcy.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbxSposobWyboruDostawcy.ForeColor = System.Drawing.Color.Purple
        Me.cbxSposobWyboruDostawcy.Location = New System.Drawing.Point(796, 180)
        Me.cbxSposobWyboruDostawcy.Name = "cbxSposobWyboruDostawcy"
        Me.cbxSposobWyboruDostawcy.Size = New System.Drawing.Size(244, 22)
        Me.cbxSposobWyboruDostawcy.TabIndex = 256
        '
        'Label38
        '
        Me.Label38.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label38.ForeColor = System.Drawing.Color.Navy
        Me.Label38.Location = New System.Drawing.Point(637, 184)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(207, 16)
        Me.Label38.TabIndex = 257
        Me.Label38.Text = "Sposób wyboru dostawcy . . . . . . . . ."
        '
        'Label37
        '
        Me.Label37.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label37.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label37.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label37.ForeColor = System.Drawing.Color.White
        Me.Label37.Location = New System.Drawing.Point(637, 159)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(403, 16)
        Me.Label37.TabIndex = 255
        Me.Label37.Text = "Sposób wyboru dostawcy"
        Me.Label37.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblPotencjalDruku
        '
        Me.lblPotencjalDruku.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPotencjalDruku.Font = New System.Drawing.Font("Arial", 24.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblPotencjalDruku.ForeColor = System.Drawing.Color.Navy
        Me.lblPotencjalDruku.Location = New System.Drawing.Point(964, 113)
        Me.lblPotencjalDruku.Name = "lblPotencjalDruku"
        Me.lblPotencjalDruku.Size = New System.Drawing.Size(76, 37)
        Me.lblPotencjalDruku.TabIndex = 254
        Me.lblPotencjalDruku.Text = "W5"
        Me.lblPotencjalDruku.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label36
        '
        Me.Label36.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label36.ForeColor = System.Drawing.Color.Navy
        Me.Label36.Location = New System.Drawing.Point(876, 113)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(51, 16)
        Me.Label36.TabIndex = 253
        Me.Label36.Text = "szt."
        '
        'Label35
        '
        Me.Label35.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label35.ForeColor = System.Drawing.Color.Navy
        Me.Label35.Location = New System.Drawing.Point(876, 134)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(51, 16)
        Me.Label35.TabIndex = 252
        Me.Label35.Text = "tys. A4"
        '
        'txtLiczbaWydrukow
        '
        Me.txtLiczbaWydrukow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLiczbaWydrukow.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtLiczbaWydrukow.ForeColor = System.Drawing.Color.Purple
        Me.txtLiczbaWydrukow.Location = New System.Drawing.Point(759, 131)
        Me.txtLiczbaWydrukow.Name = "txtLiczbaWydrukow"
        Me.txtLiczbaWydrukow.Size = New System.Drawing.Size(111, 20)
        Me.txtLiczbaWydrukow.TabIndex = 251
        '
        'Label29
        '
        Me.Label29.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label29.ForeColor = System.Drawing.Color.Navy
        Me.Label29.Location = New System.Drawing.Point(637, 134)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(207, 16)
        Me.Label29.TabIndex = 250
        Me.Label29.Text = "Liczba wydruków / rok . . . . . . ."
        '
        'txtLiczbaUrzadzen
        '
        Me.txtLiczbaUrzadzen.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLiczbaUrzadzen.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtLiczbaUrzadzen.ForeColor = System.Drawing.Color.Purple
        Me.txtLiczbaUrzadzen.Location = New System.Drawing.Point(759, 109)
        Me.txtLiczbaUrzadzen.Name = "txtLiczbaUrzadzen"
        Me.txtLiczbaUrzadzen.Size = New System.Drawing.Size(111, 20)
        Me.txtLiczbaUrzadzen.TabIndex = 249
        '
        'Label28
        '
        Me.Label28.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label28.ForeColor = System.Drawing.Color.Navy
        Me.Label28.Location = New System.Drawing.Point(637, 112)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(207, 16)
        Me.Label28.TabIndex = 248
        Me.Label28.Text = "Liczba urz¹dzeñ . . . . . . ."
        '
        'Label20
        '
        Me.Label20.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label20.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.White
        Me.Label20.Location = New System.Drawing.Point(637, 87)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(403, 16)
        Me.Label20.TabIndex = 247
        Me.Label20.Text = "Potencja³ druku"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'chkZainstalowanyMonitor
        '
        Me.chkZainstalowanyMonitor.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkZainstalowanyMonitor.ForeColor = System.Drawing.Color.Navy
        Me.chkZainstalowanyMonitor.Location = New System.Drawing.Point(414, 139)
        Me.chkZainstalowanyMonitor.Name = "chkZainstalowanyMonitor"
        Me.chkZainstalowanyMonitor.Size = New System.Drawing.Size(150, 16)
        Me.chkZainstalowanyMonitor.TabIndex = 246
        Me.chkZainstalowanyMonitor.Text = "Zainstalowany monitor . . . "
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 353)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label23)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtUwagiCharakterystyka)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label43)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtOpisPotencjalu)
        Me.SplitContainer1.Size = New System.Drawing.Size(1054, 289)
        Me.SplitContainer1.SplitterDistance = 527
        Me.SplitContainer1.TabIndex = 244
        '
        'Label23
        '
        Me.Label23.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label23.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label23.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label23.ForeColor = System.Drawing.Color.White
        Me.Label23.Location = New System.Drawing.Point(0, 0)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(524, 16)
        Me.Label23.TabIndex = 165
        Me.Label23.Text = "Uwagi"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtUwagiCharakterystyka
        '
        Me.txtUwagiCharakterystyka.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUwagiCharakterystyka.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtUwagiCharakterystyka.ForeColor = System.Drawing.Color.Purple
        Me.txtUwagiCharakterystyka.Location = New System.Drawing.Point(0, 18)
        Me.txtUwagiCharakterystyka.Multiline = True
        Me.txtUwagiCharakterystyka.Name = "txtUwagiCharakterystyka"
        Me.txtUwagiCharakterystyka.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtUwagiCharakterystyka.Size = New System.Drawing.Size(523, 271)
        Me.txtUwagiCharakterystyka.TabIndex = 112
        '
        'Label43
        '
        Me.Label43.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label43.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label43.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label43.ForeColor = System.Drawing.Color.White
        Me.Label43.Location = New System.Drawing.Point(0, 0)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(520, 16)
        Me.Label43.TabIndex = 164
        Me.Label43.Text = "Opis potencja³u"
        Me.Label43.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtOpisPotencjalu
        '
        Me.txtOpisPotencjalu.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOpisPotencjalu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtOpisPotencjalu.ForeColor = System.Drawing.Color.Purple
        Me.txtOpisPotencjalu.Location = New System.Drawing.Point(0, 18)
        Me.txtOpisPotencjalu.Multiline = True
        Me.txtOpisPotencjalu.Name = "txtOpisPotencjalu"
        Me.txtOpisPotencjalu.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtOpisPotencjalu.Size = New System.Drawing.Size(520, 271)
        Me.txtOpisPotencjalu.TabIndex = 165
        Me.txtOpisPotencjalu.Text = "162"
        '
        'lblTelefonKom_Charakterystyka
        '
        Me.lblTelefonKom_Charakterystyka.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefonKom_Charakterystyka.Location = New System.Drawing.Point(940, 56)
        Me.lblTelefonKom_Charakterystyka.Name = "lblTelefonKom_Charakterystyka"
        Me.lblTelefonKom_Charakterystyka.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefonKom_Charakterystyka.TabIndex = 242
        '
        'Label64
        '
        Me.Label64.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label64.ForeColor = System.Drawing.Color.Navy
        Me.Label64.Location = New System.Drawing.Point(891, 56)
        Me.Label64.Name = "Label64"
        Me.Label64.Size = New System.Drawing.Size(112, 16)
        Me.Label64.TabIndex = 243
        Me.Label64.Text = "Tel. kom. . . . . . . . . . . . . . . "
        '
        'lblTelefon_Charakterystyka
        '
        Me.lblTelefon_Charakterystyka.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefon_Charakterystyka.Location = New System.Drawing.Point(673, 56)
        Me.lblTelefon_Charakterystyka.Name = "lblTelefon_Charakterystyka"
        Me.lblTelefon_Charakterystyka.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefon_Charakterystyka.TabIndex = 240
        '
        'Label74
        '
        Me.Label74.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label74.ForeColor = System.Drawing.Color.Navy
        Me.Label74.Location = New System.Drawing.Point(624, 55)
        Me.Label74.Name = "Label74"
        Me.Label74.Size = New System.Drawing.Size(112, 16)
        Me.Label74.TabIndex = 241
        Me.Label74.Text = "Telefon . . . . . . . . . . . . . . "
        '
        'lblStanowisko_Charakterystyka
        '
        Me.lblStanowisko_Charakterystyka.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStanowisko_Charakterystyka.Location = New System.Drawing.Point(404, 56)
        Me.lblStanowisko_Charakterystyka.Name = "lblStanowisko_Charakterystyka"
        Me.lblStanowisko_Charakterystyka.Size = New System.Drawing.Size(200, 16)
        Me.lblStanowisko_Charakterystyka.TabIndex = 238
        '
        'Label84
        '
        Me.Label84.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label84.ForeColor = System.Drawing.Color.Navy
        Me.Label84.Location = New System.Drawing.Point(333, 56)
        Me.Label84.Name = "Label84"
        Me.Label84.Size = New System.Drawing.Size(112, 16)
        Me.Label84.TabIndex = 239
        Me.Label84.Text = "Stanowisko . . . . . . . . . . . . . . "
        '
        'txtWykazUrzadzen
        '
        Me.txtWykazUrzadzen.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtWykazUrzadzen.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtWykazUrzadzen.ForeColor = System.Drawing.Color.Purple
        Me.txtWykazUrzadzen.Location = New System.Drawing.Point(109, 160)
        Me.txtWykazUrzadzen.Multiline = True
        Me.txtWykazUrzadzen.Name = "txtWykazUrzadzen"
        Me.txtWykazUrzadzen.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtWykazUrzadzen.Size = New System.Drawing.Size(511, 140)
        Me.txtWykazUrzadzen.TabIndex = 108
        '
        'lblOfertaPlik
        '
        Me.lblOfertaPlik.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblOfertaPlik.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOfertaPlik.ForeColor = System.Drawing.Color.Purple
        Me.lblOfertaPlik.Location = New System.Drawing.Point(109, 303)
        Me.lblOfertaPlik.Name = "lblOfertaPlik"
        Me.lblOfertaPlik.Size = New System.Drawing.Size(435, 17)
        Me.lblOfertaPlik.TabIndex = 116
        Me.lblOfertaPlik.Text = " "
        '
        'Label59
        '
        Me.Label59.ForeColor = System.Drawing.Color.Navy
        Me.Label59.Location = New System.Drawing.Point(8, 303)
        Me.Label59.Name = "Label59"
        Me.Label59.Size = New System.Drawing.Size(139, 16)
        Me.Label59.TabIndex = 115
        Me.Label59.Text = "Plik z ofert¹ . . . . . . . . . . . . "
        '
        'txtTZlozenia
        '
        Me.txtTZlozenia.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtTZlozenia.ForeColor = System.Drawing.Color.Purple
        Me.txtTZlozenia.Location = New System.Drawing.Point(178, 327)
        Me.txtTZlozenia.Name = "txtTZlozenia"
        Me.txtTZlozenia.Size = New System.Drawing.Size(46, 20)
        Me.txtTZlozenia.TabIndex = 114
        Me.txtTZlozenia.Text = "RRTT"
        '
        'Label45
        '
        Me.Label45.ForeColor = System.Drawing.Color.Navy
        Me.Label45.Location = New System.Drawing.Point(8, 330)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(207, 16)
        Me.Label45.TabIndex = 113
        Me.Label45.Text = "Nr tygodnia weryfikacji oferty . . . . . . . . ."
        '
        'lblNazwa1_Charakterystyka
        '
        Me.lblNazwa1_Charakterystyka.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa1_Charakterystyka.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa1_Charakterystyka.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa1_Charakterystyka.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa1_Charakterystyka.Location = New System.Drawing.Point(112, 40)
        Me.lblNazwa1_Charakterystyka.Name = "lblNazwa1_Charakterystyka"
        Me.lblNazwa1_Charakterystyka.Size = New System.Drawing.Size(928, 16)
        Me.lblNazwa1_Charakterystyka.TabIndex = 59
        '
        'Label55
        '
        Me.Label55.Font = New System.Drawing.Font("Arial Black", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label55.ForeColor = System.Drawing.Color.Navy
        Me.Label55.Location = New System.Drawing.Point(209, 136)
        Me.Label55.Name = "Label55"
        Me.Label55.Size = New System.Drawing.Size(30, 16)
        Me.Label55.TabIndex = 91
        Me.Label55.Text = "A3"
        '
        'Label54
        '
        Me.Label54.Font = New System.Drawing.Font("Arial Black", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label54.ForeColor = System.Drawing.Color.Navy
        Me.Label54.Location = New System.Drawing.Point(159, 136)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(30, 16)
        Me.Label54.TabIndex = 90
        Me.Label54.Text = "A"
        '
        'Label52
        '
        Me.Label52.Font = New System.Drawing.Font("Arial Black", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label52.ForeColor = System.Drawing.Color.Navy
        Me.Label52.Location = New System.Drawing.Point(109, 136)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(20, 16)
        Me.Label52.TabIndex = 89
        Me.Label52.Text = "L"
        '
        'Label34
        '
        Me.Label34.ForeColor = System.Drawing.Color.Navy
        Me.Label34.Location = New System.Drawing.Point(8, 138)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(122, 16)
        Me.Label34.TabIndex = 72
        Me.Label34.Text = "Rodzaj sprzêtu . . . . . . . . ."
        '
        'Label27
        '
        Me.Label27.ForeColor = System.Drawing.Color.Navy
        Me.Label27.Location = New System.Drawing.Point(8, 163)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(112, 16)
        Me.Label27.TabIndex = 69
        Me.Label27.Text = "Wykaz urz¹dzeñ . . . . . . . . ."
        '
        'lblNazwaHandlowa_Charakterystyka
        '
        Me.lblNazwaHandlowa_Charakterystyka.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwaHandlowa_Charakterystyka.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwaHandlowa_Charakterystyka.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwaHandlowa_Charakterystyka.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwaHandlowa_Charakterystyka.Location = New System.Drawing.Point(112, 24)
        Me.lblNazwaHandlowa_Charakterystyka.Name = "lblNazwaHandlowa_Charakterystyka"
        Me.lblNazwaHandlowa_Charakterystyka.Size = New System.Drawing.Size(928, 16)
        Me.lblNazwaHandlowa_Charakterystyka.TabIndex = 62
        '
        'Label19
        '
        Me.Label19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.Navy
        Me.Label19.Location = New System.Drawing.Point(8, 40)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(112, 16)
        Me.Label19.TabIndex = 63
        Me.Label19.Text = "Adres . . . . . . . . . . . . . . "
        '
        'lblTofi_Charakterystyka
        '
        Me.lblTofi_Charakterystyka.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTofi_Charakterystyka.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTofi_Charakterystyka.ForeColor = System.Drawing.Color.Purple
        Me.lblTofi_Charakterystyka.Location = New System.Drawing.Point(305, 8)
        Me.lblTofi_Charakterystyka.Name = "lblTofi_Charakterystyka"
        Me.lblTofi_Charakterystyka.Size = New System.Drawing.Size(735, 16)
        Me.lblTofi_Charakterystyka.TabIndex = 61
        '
        'lblNazwa2_Charakterystyka
        '
        Me.lblNazwa2_Charakterystyka.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa2_Charakterystyka.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblNazwa2_Charakterystyka.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwa2_Charakterystyka.Location = New System.Drawing.Point(112, 56)
        Me.lblNazwa2_Charakterystyka.Name = "lblNazwa2_Charakterystyka"
        Me.lblNazwa2_Charakterystyka.Size = New System.Drawing.Size(200, 16)
        Me.lblNazwa2_Charakterystyka.TabIndex = 60
        '
        'lblIdent_Charakterystyka
        '
        Me.lblIdent_Charakterystyka.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent_Charakterystyka.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblIdent_Charakterystyka.ForeColor = System.Drawing.Color.Purple
        Me.lblIdent_Charakterystyka.Location = New System.Drawing.Point(112, 8)
        Me.lblIdent_Charakterystyka.Name = "lblIdent_Charakterystyka"
        Me.lblIdent_Charakterystyka.Size = New System.Drawing.Size(112, 16)
        Me.lblIdent_Charakterystyka.TabIndex = 58
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Location = New System.Drawing.Point(0, 80)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1054, 1)
        Me.Panel1.TabIndex = 57
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Navy
        Me.Label15.Location = New System.Drawing.Point(256, 8)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(112, 16)
        Me.Label15.TabIndex = 56
        Me.Label15.Text = "Nr TOFI . . . . . "
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Navy
        Me.Label17.Location = New System.Drawing.Point(8, 24)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(112, 16)
        Me.Label17.TabIndex = 55
        Me.Label17.Text = "Nazwa klienta . . . . . . . . ."
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Navy
        Me.Label18.Location = New System.Drawing.Point(8, 8)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(112, 16)
        Me.Label18.TabIndex = 54
        Me.Label18.Text = "Identyfikator . . . . . . . . ."
        '
        'Label60
        '
        Me.Label60.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label60.ForeColor = System.Drawing.Color.Navy
        Me.Label60.Location = New System.Drawing.Point(8, 56)
        Me.Label60.Name = "Label60"
        Me.Label60.Size = New System.Drawing.Size(112, 16)
        Me.Label60.TabIndex = 99
        Me.Label60.Text = "Osoba kontaktowa . . . . . . . . . . . . . . "
        '
        'cmdPKontakt
        '
        Me.cmdPKontakt.Image = CType(resources.GetObject("cmdPKontakt.Image"), System.Drawing.Image)
        Me.cmdPKontakt.Location = New System.Drawing.Point(964, 56)
        Me.cmdPKontakt.Name = "cmdPKontakt"
        Me.cmdPKontakt.Size = New System.Drawing.Size(24, 20)
        Me.cmdPKontakt.TabIndex = 22
        '
        'txtPKontakt
        '
        Me.txtPKontakt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtPKontakt.ForeColor = System.Drawing.Color.Purple
        Me.txtPKontakt.Location = New System.Drawing.Point(876, 56)
        Me.txtPKontakt.Name = "txtPKontakt"
        Me.txtPKontakt.Size = New System.Drawing.Size(96, 20)
        Me.txtPKontakt.TabIndex = 101
        '
        'Label21
        '
        Me.Label21.ForeColor = System.Drawing.Color.Navy
        Me.Label21.Location = New System.Drawing.Point(772, 58)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(112, 16)
        Me.Label21.TabIndex = 65
        Me.Label21.Text = "Pierwszy kontakt . . . . . . . . ."
        '
        'Opiekun
        '
        Me.Opiekun.BackColor = System.Drawing.SystemColors.Control
        Me.Opiekun.Controls.Add(Me.PanelSprzedaz)
        Me.Opiekun.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Opiekun.ForeColor = System.Drawing.Color.Purple
        Me.Opiekun.Location = New System.Drawing.Point(4, 22)
        Me.Opiekun.Name = "Opiekun"
        Me.Opiekun.Size = New System.Drawing.Size(1054, 645)
        Me.Opiekun.TabIndex = 1
        Me.Opiekun.Text = "Opiekun"
        '
        'PanelSprzedaz
        '
        Me.PanelSprzedaz.Controls.Add(Me.txtPlanKrotkookresowy)
        Me.PanelSprzedaz.Controls.Add(Me.txtPlanDlugookresowy)
        Me.PanelSprzedaz.Controls.Add(Me.Label1)
        Me.PanelSprzedaz.Controls.Add(Me.Label22)
        Me.PanelSprzedaz.Controls.Add(Me.txtSprzedazUwagi)
        Me.PanelSprzedaz.Controls.Add(Me.lblTelefonKom_Sprzedaz)
        Me.PanelSprzedaz.Controls.Add(Me.Label44)
        Me.PanelSprzedaz.Controls.Add(Me.Label62)
        Me.PanelSprzedaz.Controls.Add(Me.lblTelefon_Sprzedaz)
        Me.PanelSprzedaz.Controls.Add(Me.Label41)
        Me.PanelSprzedaz.Controls.Add(Me.lblStanowisko_Sprzedaz)
        Me.PanelSprzedaz.Controls.Add(Me.Label39)
        Me.PanelSprzedaz.Controls.Add(Me.cmdSprzedazOKonT)
        Me.PanelSprzedaz.Controls.Add(Me.cmdSprzedazOpiekun)
        Me.PanelSprzedaz.Controls.Add(Me.cmdSprzedazNKonT)
        Me.PanelSprzedaz.Controls.Add(Me.txtSprzedazNKonT)
        Me.PanelSprzedaz.Controls.Add(Me.txtSprzedazOKonT)
        Me.PanelSprzedaz.Controls.Add(Me.txtSprzedazOpiekun)
        Me.PanelSprzedaz.Controls.Add(Me.SplitContainer6)
        Me.PanelSprzedaz.Controls.Add(Me.Label11)
        Me.PanelSprzedaz.Controls.Add(Me.cmdHarmonogramWizyt)
        Me.PanelSprzedaz.Controls.Add(Me.Label48)
        Me.PanelSprzedaz.Controls.Add(Me.Label77)
        Me.PanelSprzedaz.Controls.Add(Me.Label87)
        Me.PanelSprzedaz.Controls.Add(Me.Panel11)
        Me.PanelSprzedaz.Controls.Add(Me.cbxSprzedazNKontR)
        Me.PanelSprzedaz.Controls.Add(Me.cbxSprzedazOKontR)
        Me.PanelSprzedaz.Controls.Add(Me.Panel12)
        Me.PanelSprzedaz.Controls.Add(Me.lblOstTransak)
        Me.PanelSprzedaz.Controls.Add(Me.Label83)
        Me.PanelSprzedaz.Controls.Add(Me.cbxSprzedazNKonD)
        Me.PanelSprzedaz.Controls.Add(Me.lblosoba_Sprzedaz)
        Me.PanelSprzedaz.Controls.Add(Me.Label65)
        Me.PanelSprzedaz.Controls.Add(Me.lbladres_Sprzedaz)
        Me.PanelSprzedaz.Controls.Add(Me.cbxSprzedazOKonD)
        Me.PanelSprzedaz.Controls.Add(Me.Label49)
        Me.PanelSprzedaz.Controls.Add(Me.Label50)
        Me.PanelSprzedaz.Controls.Add(Me.Label51)
        Me.PanelSprzedaz.Controls.Add(Me.lblNazwa_Sprzedaz)
        Me.PanelSprzedaz.Controls.Add(Me.Label13)
        Me.PanelSprzedaz.Controls.Add(Me.lblTofi_Sprzedaz)
        Me.PanelSprzedaz.Controls.Add(Me.Panel3)
        Me.PanelSprzedaz.Controls.Add(Me.lblIdent_Sprzedaz)
        Me.PanelSprzedaz.Controls.Add(Me.Label30)
        Me.PanelSprzedaz.Controls.Add(Me.Label31)
        Me.PanelSprzedaz.Controls.Add(Me.Label32)
        Me.PanelSprzedaz.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelSprzedaz.Location = New System.Drawing.Point(0, 0)
        Me.PanelSprzedaz.Name = "PanelSprzedaz"
        Me.PanelSprzedaz.Size = New System.Drawing.Size(1054, 645)
        Me.PanelSprzedaz.TabIndex = 359
        '
        'txtPlanKrotkookresowy
        '
        Me.txtPlanKrotkookresowy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtPlanKrotkookresowy.ForeColor = System.Drawing.Color.Purple
        Me.txtPlanKrotkookresowy.Location = New System.Drawing.Point(469, 166)
        Me.txtPlanKrotkookresowy.Name = "txtPlanKrotkookresowy"
        Me.txtPlanKrotkookresowy.Size = New System.Drawing.Size(81, 20)
        Me.txtPlanKrotkookresowy.TabIndex = 239
        Me.txtPlanKrotkookresowy.Text = "8"
        '
        'txtPlanDlugookresowy
        '
        Me.txtPlanDlugookresowy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtPlanDlugookresowy.ForeColor = System.Drawing.Color.Purple
        Me.txtPlanDlugookresowy.Location = New System.Drawing.Point(469, 142)
        Me.txtPlanDlugookresowy.Name = "txtPlanDlugookresowy"
        Me.txtPlanDlugookresowy.Size = New System.Drawing.Size(81, 20)
        Me.txtPlanDlugookresowy.TabIndex = 238
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(352, 169)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(134, 16)
        Me.Label1.TabIndex = 241
        Me.Label1.Text = "Plan krótkookresowy . . . . . . . . ."
        '
        'Label22
        '
        Me.Label22.ForeColor = System.Drawing.Color.Navy
        Me.Label22.Location = New System.Drawing.Point(352, 145)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(134, 16)
        Me.Label22.TabIndex = 240
        Me.Label22.Text = "Plan d³ugookresowy . . . . . . . . ."
        '
        'txtSprzedazUwagi
        '
        Me.txtSprzedazUwagi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSprzedazUwagi.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtSprzedazUwagi.ForeColor = System.Drawing.Color.Purple
        Me.txtSprzedazUwagi.Location = New System.Drawing.Point(570, 106)
        Me.txtSprzedazUwagi.Multiline = True
        Me.txtSprzedazUwagi.Name = "txtSprzedazUwagi"
        Me.txtSprzedazUwagi.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSprzedazUwagi.Size = New System.Drawing.Size(470, 96)
        Me.txtSprzedazUwagi.TabIndex = 163
        Me.txtSprzedazUwagi.Text = "162"
        '
        'lblTelefonKom_Sprzedaz
        '
        Me.lblTelefonKom_Sprzedaz.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefonKom_Sprzedaz.Location = New System.Drawing.Point(940, 56)
        Me.lblTelefonKom_Sprzedaz.Name = "lblTelefonKom_Sprzedaz"
        Me.lblTelefonKom_Sprzedaz.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefonKom_Sprzedaz.TabIndex = 236
        '
        'Label44
        '
        Me.Label44.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label44.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label44.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label44.ForeColor = System.Drawing.Color.White
        Me.Label44.Location = New System.Drawing.Point(570, 87)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(470, 16)
        Me.Label44.TabIndex = 81
        Me.Label44.Text = "Uwagi"
        Me.Label44.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label62
        '
        Me.Label62.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label62.ForeColor = System.Drawing.Color.Navy
        Me.Label62.Location = New System.Drawing.Point(891, 55)
        Me.Label62.Name = "Label62"
        Me.Label62.Size = New System.Drawing.Size(112, 16)
        Me.Label62.TabIndex = 237
        Me.Label62.Text = "Tel. kom. . . . . . . . . . . . . . . "
        '
        'lblTelefon_Sprzedaz
        '
        Me.lblTelefon_Sprzedaz.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTelefon_Sprzedaz.Location = New System.Drawing.Point(673, 56)
        Me.lblTelefon_Sprzedaz.Name = "lblTelefon_Sprzedaz"
        Me.lblTelefon_Sprzedaz.Size = New System.Drawing.Size(200, 16)
        Me.lblTelefon_Sprzedaz.TabIndex = 234
        '
        'Label41
        '
        Me.Label41.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label41.ForeColor = System.Drawing.Color.Navy
        Me.Label41.Location = New System.Drawing.Point(624, 56)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(112, 16)
        Me.Label41.TabIndex = 235
        Me.Label41.Text = "Telefon . . . . . . . . . . . . . . "
        '
        'lblStanowisko_Sprzedaz
        '
        Me.lblStanowisko_Sprzedaz.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStanowisko_Sprzedaz.Location = New System.Drawing.Point(404, 56)
        Me.lblStanowisko_Sprzedaz.Name = "lblStanowisko_Sprzedaz"
        Me.lblStanowisko_Sprzedaz.Size = New System.Drawing.Size(200, 16)
        Me.lblStanowisko_Sprzedaz.TabIndex = 232
        '
        'Label39
        '
        Me.Label39.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label39.ForeColor = System.Drawing.Color.Navy
        Me.Label39.Location = New System.Drawing.Point(333, 56)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(112, 16)
        Me.Label39.TabIndex = 233
        Me.Label39.Text = "Stanowisko . . . . . . . . . . . . . . "
        '
        'cmdSprzedazOKonT
        '
        Me.cmdSprzedazOKonT.Image = CType(resources.GetObject("cmdSprzedazOKonT.Image"), System.Drawing.Image)
        Me.cmdSprzedazOKonT.Location = New System.Drawing.Point(183, 142)
        Me.cmdSprzedazOKonT.Name = "cmdSprzedazOKonT"
        Me.cmdSprzedazOKonT.Size = New System.Drawing.Size(24, 22)
        Me.cmdSprzedazOKonT.TabIndex = 56
        '
        'cmdSprzedazOpiekun
        '
        Me.cmdSprzedazOpiekun.Image = CType(resources.GetObject("cmdSprzedazOpiekun.Image"), System.Drawing.Image)
        Me.cmdSprzedazOpiekun.Location = New System.Drawing.Point(208, 88)
        Me.cmdSprzedazOpiekun.Name = "cmdSprzedazOpiekun"
        Me.cmdSprzedazOpiekun.Size = New System.Drawing.Size(32, 23)
        Me.cmdSprzedazOpiekun.TabIndex = 52
        Me.cmdSprzedazOpiekun.Text = "2"
        '
        'cmdSprzedazNKonT
        '
        Me.cmdSprzedazNKonT.Image = CType(resources.GetObject("cmdSprzedazNKonT.Image"), System.Drawing.Image)
        Me.cmdSprzedazNKonT.Location = New System.Drawing.Point(183, 166)
        Me.cmdSprzedazNKonT.Name = "cmdSprzedazNKonT"
        Me.cmdSprzedazNKonT.Size = New System.Drawing.Size(24, 22)
        Me.cmdSprzedazNKonT.TabIndex = 60
        '
        'txtSprzedazNKonT
        '
        Me.txtSprzedazNKonT.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtSprzedazNKonT.ForeColor = System.Drawing.Color.Purple
        Me.txtSprzedazNKonT.Location = New System.Drawing.Point(163, 167)
        Me.txtSprzedazNKonT.Name = "txtSprzedazNKonT"
        Me.txtSprzedazNKonT.Size = New System.Drawing.Size(44, 20)
        Me.txtSprzedazNKonT.TabIndex = 59
        Me.txtSprzedazNKonT.Text = "8"
        '
        'txtSprzedazOKonT
        '
        Me.txtSprzedazOKonT.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtSprzedazOKonT.ForeColor = System.Drawing.Color.Purple
        Me.txtSprzedazOKonT.Location = New System.Drawing.Point(163, 143)
        Me.txtSprzedazOKonT.Name = "txtSprzedazOKonT"
        Me.txtSprzedazOKonT.Size = New System.Drawing.Size(44, 20)
        Me.txtSprzedazOKonT.TabIndex = 55
        '
        'txtSprzedazOpiekun
        '
        Me.txtSprzedazOpiekun.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtSprzedazOpiekun.ForeColor = System.Drawing.Color.Purple
        Me.txtSprzedazOpiekun.Location = New System.Drawing.Point(112, 88)
        Me.txtSprzedazOpiekun.Name = "txtSprzedazOpiekun"
        Me.txtSprzedazOpiekun.Size = New System.Drawing.Size(112, 20)
        Me.txtSprzedazOpiekun.TabIndex = 51
        Me.txtSprzedazOpiekun.Text = "1"
        '
        'SplitContainer6
        '
        Me.SplitContainer6.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer6.Location = New System.Drawing.Point(0, 230)
        Me.SplitContainer6.Name = "SplitContainer6"
        Me.SplitContainer6.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer6.Panel1
        '
        Me.SplitContainer6.Panel1.Controls.Add(Me.Panel9)
        '
        'SplitContainer6.Panel2
        '
        Me.SplitContainer6.Panel2.Controls.Add(Me.dagKontakty)
        Me.SplitContainer6.Panel2.Controls.Add(Me.Label56)
        Me.SplitContainer6.Size = New System.Drawing.Size(1052, 414)
        Me.SplitContainer6.SplitterDistance = 114
        Me.SplitContainer6.TabIndex = 0
        '
        'Panel9
        '
        Me.Panel9.Controls.Add(Me.dagCoDalej)
        Me.Panel9.Controls.Add(Me.Label46)
        Me.Panel9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel9.Location = New System.Drawing.Point(0, 0)
        Me.Panel9.Name = "Panel9"
        Me.Panel9.Size = New System.Drawing.Size(1052, 114)
        Me.Panel9.TabIndex = 173
        '
        'dagCoDalej
        '
        Me.dagCoDalej.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagCoDalej.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dagCoDalej.ContextMenuStrip = Me.MenuCoDalej
        Me.dagCoDalej.Location = New System.Drawing.Point(3, 21)
        Me.dagCoDalej.Name = "dagCoDalej"
        Me.dagCoDalej.Size = New System.Drawing.Size(1049, 93)
        Me.dagCoDalej.TabIndex = 174
        '
        'Label46
        '
        Me.Label46.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label46.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label46.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label46.ForeColor = System.Drawing.Color.White
        Me.Label46.Location = New System.Drawing.Point(0, 2)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(1049, 16)
        Me.Label46.TabIndex = 82
        Me.Label46.Text = "Plan - co dalej"
        Me.Label46.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dagKontakty
        '
        Me.dagKontakty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dagKontakty.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dagKontakty.ContextMenuStrip = Me.MenuKontakty
        Me.dagKontakty.Location = New System.Drawing.Point(0, 19)
        Me.dagKontakty.Name = "dagKontakty"
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dagKontakty.RowHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.dagKontakty.Size = New System.Drawing.Size(1052, 277)
        Me.dagKontakty.TabIndex = 171
        '
        'Label56
        '
        Me.Label56.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label56.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label56.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label56.ForeColor = System.Drawing.Color.White
        Me.Label56.Location = New System.Drawing.Point(0, 0)
        Me.Label56.Name = "Label56"
        Me.Label56.Size = New System.Drawing.Size(1049, 16)
        Me.Label56.TabIndex = 165
        Me.Label56.Text = "Kontakty handlowe"
        Me.Label56.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Navy
        Me.Label11.Location = New System.Drawing.Point(9, 120)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(86, 16)
        Me.Label11.TabIndex = 208
        Me.Label11.Text = "SPRZEDA¯ :"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdHarmonogramWizyt
        '
        Me.cmdHarmonogramWizyt.ForeColor = System.Drawing.Color.Black
        Me.cmdHarmonogramWizyt.Image = CType(resources.GetObject("cmdHarmonogramWizyt.Image"), System.Drawing.Image)
        Me.cmdHarmonogramWizyt.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdHarmonogramWizyt.Location = New System.Drawing.Point(299, 165)
        Me.cmdHarmonogramWizyt.Name = "cmdHarmonogramWizyt"
        Me.cmdHarmonogramWizyt.Size = New System.Drawing.Size(24, 23)
        Me.cmdHarmonogramWizyt.TabIndex = 205
        Me.cmdHarmonogramWizyt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label48
        '
        Me.Label48.ForeColor = System.Drawing.Color.Navy
        Me.Label48.Location = New System.Drawing.Point(213, 121)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(80, 16)
        Me.Label48.TabIndex = 122
        Me.Label48.Text = "Dzieñ"
        Me.Label48.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label77
        '
        Me.Label77.ForeColor = System.Drawing.Color.Navy
        Me.Label77.Location = New System.Drawing.Point(161, 120)
        Me.Label77.Name = "Label77"
        Me.Label77.Size = New System.Drawing.Size(50, 16)
        Me.Label77.TabIndex = 121
        Me.Label77.Text = "Nr.tyg."
        Me.Label77.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label87
        '
        Me.Label87.ForeColor = System.Drawing.Color.Navy
        Me.Label87.Location = New System.Drawing.Point(112, 120)
        Me.Label87.Name = "Label87"
        Me.Label87.Size = New System.Drawing.Size(53, 16)
        Me.Label87.TabIndex = 120
        Me.Label87.Text = "Rok"
        Me.Label87.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel11
        '
        Me.Panel11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel11.Location = New System.Drawing.Point(97, 121)
        Me.Panel11.Name = "Panel11"
        Me.Panel11.Size = New System.Drawing.Size(1, 68)
        Me.Panel11.TabIndex = 119
        '
        'cbxSprzedazNKontR
        '
        Me.cbxSprzedazNKontR.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxSprzedazNKontR.ForeColor = System.Drawing.Color.Purple
        Me.cbxSprzedazNKontR.Location = New System.Drawing.Point(104, 167)
        Me.cbxSprzedazNKontR.Name = "cbxSprzedazNKontR"
        Me.cbxSprzedazNKontR.Size = New System.Drawing.Size(53, 22)
        Me.cbxSprzedazNKontR.TabIndex = 58
        Me.cbxSprzedazNKontR.Text = "2007"
        '
        'cbxSprzedazOKontR
        '
        Me.cbxSprzedazOKontR.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxSprzedazOKontR.ForeColor = System.Drawing.Color.Purple
        Me.cbxSprzedazOKontR.Location = New System.Drawing.Point(104, 142)
        Me.cbxSprzedazOKontR.Name = "cbxSprzedazOKontR"
        Me.cbxSprzedazOKontR.Size = New System.Drawing.Size(53, 22)
        Me.cbxSprzedazOKontR.TabIndex = 54
        Me.cbxSprzedazOKontR.Text = "2007"
        '
        'Panel12
        '
        Me.Panel12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel12.Location = New System.Drawing.Point(19, 139)
        Me.Panel12.Name = "Panel12"
        Me.Panel12.Size = New System.Drawing.Size(280, 1)
        Me.Panel12.TabIndex = 116
        '
        'lblOstTransak
        '
        Me.lblOstTransak.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblOstTransak.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOstTransak.Location = New System.Drawing.Point(112, 206)
        Me.lblOstTransak.Name = "lblOstTransak"
        Me.lblOstTransak.Size = New System.Drawing.Size(928, 16)
        Me.lblOstTransak.TabIndex = 115
        '
        'Label83
        '
        Me.Label83.ForeColor = System.Drawing.Color.Navy
        Me.Label83.Location = New System.Drawing.Point(9, 207)
        Me.Label83.Name = "Label83"
        Me.Label83.Size = New System.Drawing.Size(112, 16)
        Me.Label83.TabIndex = 114
        Me.Label83.Text = "Ostatnie transakcje . . . . . . . . ."
        '
        'cbxSprzedazNKonD
        '
        Me.cbxSprzedazNKonD.ForeColor = System.Drawing.Color.Purple
        Me.cbxSprzedazNKonD.Location = New System.Drawing.Point(213, 166)
        Me.cbxSprzedazNKonD.Name = "cbxSprzedazNKonD"
        Me.cbxSprzedazNKonD.Size = New System.Drawing.Size(80, 22)
        Me.cbxSprzedazNKonD.TabIndex = 61
        '
        'lblosoba_Sprzedaz
        '
        Me.lblosoba_Sprzedaz.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblosoba_Sprzedaz.Location = New System.Drawing.Point(112, 56)
        Me.lblosoba_Sprzedaz.Name = "lblosoba_Sprzedaz"
        Me.lblosoba_Sprzedaz.Size = New System.Drawing.Size(200, 16)
        Me.lblosoba_Sprzedaz.TabIndex = 60
        '
        'Label65
        '
        Me.Label65.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label65.ForeColor = System.Drawing.Color.Navy
        Me.Label65.Location = New System.Drawing.Point(8, 56)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(112, 16)
        Me.Label65.TabIndex = 106
        Me.Label65.Text = "Osoba kontaktowa . . . . . . . . . . . . . . "
        '
        'lbladres_Sprzedaz
        '
        Me.lbladres_Sprzedaz.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbladres_Sprzedaz.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbladres_Sprzedaz.Location = New System.Drawing.Point(112, 40)
        Me.lbladres_Sprzedaz.Name = "lbladres_Sprzedaz"
        Me.lbladres_Sprzedaz.Size = New System.Drawing.Size(928, 16)
        Me.lbladres_Sprzedaz.TabIndex = 59
        '
        'cbxSprzedazOKonD
        '
        Me.cbxSprzedazOKonD.ForeColor = System.Drawing.Color.Purple
        Me.cbxSprzedazOKonD.Location = New System.Drawing.Point(213, 142)
        Me.cbxSprzedazOKonD.Name = "cbxSprzedazOKonD"
        Me.cbxSprzedazOKonD.Size = New System.Drawing.Size(80, 22)
        Me.cbxSprzedazOKonD.TabIndex = 57
        '
        'Label49
        '
        Me.Label49.ForeColor = System.Drawing.Color.Navy
        Me.Label49.Location = New System.Drawing.Point(9, 170)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(89, 16)
        Me.Label49.TabIndex = 90
        Me.Label49.Text = "Kolejny kontakt . . . . . . . . ."
        '
        'Label50
        '
        Me.Label50.ForeColor = System.Drawing.Color.Navy
        Me.Label50.Location = New System.Drawing.Point(9, 146)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(93, 16)
        Me.Label50.TabIndex = 88
        Me.Label50.Text = "Ostatni kontakt . . . . . . . . ."
        '
        'Label51
        '
        Me.Label51.ForeColor = System.Drawing.Color.Navy
        Me.Label51.Location = New System.Drawing.Point(8, 88)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(112, 16)
        Me.Label51.TabIndex = 86
        Me.Label51.Text = "Sprzeda¿ - opiekun . . . . . . . . ."
        '
        'lblNazwa_Sprzedaz
        '
        Me.lblNazwa_Sprzedaz.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwa_Sprzedaz.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNazwa_Sprzedaz.Location = New System.Drawing.Point(112, 24)
        Me.lblNazwa_Sprzedaz.Name = "lblNazwa_Sprzedaz"
        Me.lblNazwa_Sprzedaz.Size = New System.Drawing.Size(928, 16)
        Me.lblNazwa_Sprzedaz.TabIndex = 62
        '
        'Label13
        '
        Me.Label13.ForeColor = System.Drawing.Color.Navy
        Me.Label13.Location = New System.Drawing.Point(8, 40)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(112, 16)
        Me.Label13.TabIndex = 63
        Me.Label13.Text = "Adres . . . . . . . . . . . . . . ."
        '
        'lblTofi_Sprzedaz
        '
        Me.lblTofi_Sprzedaz.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTofi_Sprzedaz.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTofi_Sprzedaz.Location = New System.Drawing.Point(305, 8)
        Me.lblTofi_Sprzedaz.Name = "lblTofi_Sprzedaz"
        Me.lblTofi_Sprzedaz.Size = New System.Drawing.Size(735, 16)
        Me.lblTofi_Sprzedaz.TabIndex = 61
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Location = New System.Drawing.Point(0, 80)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1054, 1)
        Me.Panel3.TabIndex = 57
        '
        'lblIdent_Sprzedaz
        '
        Me.lblIdent_Sprzedaz.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIdent_Sprzedaz.Location = New System.Drawing.Point(112, 8)
        Me.lblIdent_Sprzedaz.Name = "lblIdent_Sprzedaz"
        Me.lblIdent_Sprzedaz.Size = New System.Drawing.Size(112, 16)
        Me.lblIdent_Sprzedaz.TabIndex = 58
        '
        'Label30
        '
        Me.Label30.ForeColor = System.Drawing.Color.Navy
        Me.Label30.Location = New System.Drawing.Point(256, 8)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(112, 16)
        Me.Label30.TabIndex = 56
        Me.Label30.Text = "Nr TOFI . . . . . "
        '
        'Label31
        '
        Me.Label31.ForeColor = System.Drawing.Color.Navy
        Me.Label31.Location = New System.Drawing.Point(8, 24)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(112, 16)
        Me.Label31.TabIndex = 55
        Me.Label31.Text = "Nazwa klienta . . . . . . . . ."
        '
        'Label32
        '
        Me.Label32.ForeColor = System.Drawing.Color.Navy
        Me.Label32.Location = New System.Drawing.Point(8, 8)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(112, 16)
        Me.Label32.TabIndex = 54
        Me.Label32.Text = "Identyfikator . . . . . . . . ."
        '
        'txtNryKartPayBack
        '
        Me.txtNryKartPayBack.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtNryKartPayBack.ForeColor = System.Drawing.Color.Purple
        Me.txtNryKartPayBack.Location = New System.Drawing.Point(217, 56)
        Me.txtNryKartPayBack.Name = "txtNryKartPayBack"
        Me.txtNryKartPayBack.Size = New System.Drawing.Size(82, 20)
        Me.txtNryKartPayBack.TabIndex = 169
        Me.txtNryKartPayBack.Text = "1"
        '
        'cmdEdytujNumeryKartPayBack
        '
        Me.cmdEdytujNumeryKartPayBack.Image = CType(resources.GetObject("cmdEdytujNumeryKartPayBack.Image"), System.Drawing.Image)
        Me.cmdEdytujNumeryKartPayBack.Location = New System.Drawing.Point(305, 55)
        Me.cmdEdytujNumeryKartPayBack.Name = "cmdEdytujNumeryKartPayBack"
        Me.cmdEdytujNumeryKartPayBack.Size = New System.Drawing.Size(32, 23)
        Me.cmdEdytujNumeryKartPayBack.TabIndex = 170
        Me.cmdEdytujNumeryKartPayBack.Text = "2"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 55)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(68, 22)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 168
        Me.PictureBox1.TabStop = False
        '
        'chkKartyPayback
        '
        Me.chkKartyPayback.Location = New System.Drawing.Point(112, 58)
        Me.chkKartyPayback.Name = "chkKartyPayback"
        Me.chkKartyPayback.Size = New System.Drawing.Size(16, 16)
        Me.chkKartyPayback.TabIndex = 124
        Me.chkKartyPayback.TabStop = False
        '
        'Label88
        '
        Me.Label88.ForeColor = System.Drawing.Color.Navy
        Me.Label88.Location = New System.Drawing.Point(80, 58)
        Me.Label88.Name = "Label88"
        Me.Label88.Size = New System.Drawing.Size(35, 16)
        Me.Label88.TabIndex = 123
        Me.Label88.Text = "Karty   . . . . . . ."
        '
        'Identyfikacja
        '
        Me.Identyfikacja.BackColor = System.Drawing.SystemColors.Control
        Me.Identyfikacja.Controls.Add(Me.Label10)
        Me.Identyfikacja.Controls.Add(Me.txtNryKartPayBack)
        Me.Identyfikacja.Controls.Add(Me.lblNryKartPayback)
        Me.Identyfikacja.Controls.Add(Me.cmdPKontakt)
        Me.Identyfikacja.Controls.Add(Me.txtPKontakt)
        Me.Identyfikacja.Controls.Add(Me.cmdEdytujNumeryKartPayBack)
        Me.Identyfikacja.Controls.Add(Me.chkKartyPayback)
        Me.Identyfikacja.Controls.Add(Me.cmdNowyNrKarty)
        Me.Identyfikacja.Controls.Add(Me.txtIDDomu)
        Me.Identyfikacja.Controls.Add(Me.txtRejon)
        Me.Identyfikacja.Controls.Add(Me.txtNrDomu)
        Me.Identyfikacja.Controls.Add(Me.txtMiejscowosc)
        Me.Identyfikacja.Controls.Add(Me.txtNazwa1)
        Me.Identyfikacja.Controls.Add(Me.txtKtoWpisal)
        Me.Identyfikacja.Controls.Add(Me.txtUlica)
        Me.Identyfikacja.Controls.Add(Me.txtTofi)
        Me.Identyfikacja.Controls.Add(Me.txteMail)
        Me.Identyfikacja.Controls.Add(Me.txtFax)
        Me.Identyfikacja.Controls.Add(Me.txtTelefon)
        Me.Identyfikacja.Controls.Add(Me.txtNIP)
        Me.Identyfikacja.Controls.Add(Me.txtKodP)
        Me.Identyfikacja.Controls.Add(Me.txtIdent)
        Me.Identyfikacja.Controls.Add(Me.Button1)
        Me.Identyfikacja.Controls.Add(Me.cmdNrKarty)
        Me.Identyfikacja.Controls.Add(Me.Label26)
        Me.Identyfikacja.Controls.Add(Me.cmdKtoWpisal)
        Me.Identyfikacja.Controls.Add(Me.Label33)
        Me.Identyfikacja.Controls.Add(Me.Label12)
        Me.Identyfikacja.Controls.Add(Me.Label14)
        Me.Identyfikacja.Controls.Add(Me.Label2)
        Me.Identyfikacja.Controls.Add(Me.Label16)
        Me.Identyfikacja.Controls.Add(Me.Label9)
        Me.Identyfikacja.Controls.Add(Me.Label8)
        Me.Identyfikacja.Controls.Add(Me.Label7)
        Me.Identyfikacja.Controls.Add(Me.Label6)
        Me.Identyfikacja.Controls.Add(Me.Label5)
        Me.Identyfikacja.Controls.Add(Me.Label3)
        Me.Identyfikacja.Controls.Add(Me.Label4)
        Me.Identyfikacja.Controls.Add(Me.PictureBox1)
        Me.Identyfikacja.Controls.Add(Me.Label88)
        Me.Identyfikacja.Controls.Add(Me.Label21)
        Me.Identyfikacja.Controls.Add(Me.cmdWszystkoOsoby)
        Me.Identyfikacja.Controls.Add(Me.stbFiltryOsoby)
        Me.Identyfikacja.Controls.Add(Me.dagOsoby)
        Me.Identyfikacja.Controls.Add(Me.stbOsoby)
        Me.Identyfikacja.Controls.Add(Me.cmdEdytujOsoby)
        Me.Identyfikacja.Controls.Add(Me.cmdUsuñOsoby)
        Me.Identyfikacja.Controls.Add(Me.cmdDodajOsoby)
        Me.Identyfikacja.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Identyfikacja.ForeColor = System.Drawing.Color.Navy
        Me.Identyfikacja.Location = New System.Drawing.Point(4, 22)
        Me.Identyfikacja.Name = "Identyfikacja"
        Me.Identyfikacja.Size = New System.Drawing.Size(1054, 645)
        Me.Identyfikacja.TabIndex = 0
        Me.Identyfikacja.Text = "Identyfikacja"
        '
        'Label10
        '
        Me.Label10.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Location = New System.Drawing.Point(11, 230)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(1031, 16)
        Me.Label10.TabIndex = 208
        Me.Label10.Text = "Osoby Kontaktowe"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblNryKartPayback
        '
        Me.lblNryKartPayback.ForeColor = System.Drawing.Color.Navy
        Me.lblNryKartPayback.Location = New System.Drawing.Point(154, 58)
        Me.lblNryKartPayback.Name = "lblNryKartPayback"
        Me.lblNryKartPayback.Size = New System.Drawing.Size(136, 16)
        Me.lblNryKartPayback.TabIndex = 171
        Me.lblNryKartPayback.Text = "Nry Kart . . . . . . . . "
        '
        'cmdNowyNrKarty
        '
        Me.cmdNowyNrKarty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdNowyNrKarty.ForeColor = System.Drawing.Color.Black
        Me.cmdNowyNrKarty.Location = New System.Drawing.Point(230, 7)
        Me.cmdNowyNrKarty.Name = "cmdNowyNrKarty"
        Me.cmdNowyNrKarty.Size = New System.Drawing.Size(50, 23)
        Me.cmdNowyNrKarty.TabIndex = 2
        Me.cmdNowyNrKarty.Tag = "Nadaj kolejny Identyfikator Klienta"
        Me.cmdNowyNrKarty.Text = "Nowy"
        '
        'txtIDDomu
        '
        Me.txtIDDomu.ForeColor = System.Drawing.Color.Purple
        Me.txtIDDomu.Location = New System.Drawing.Point(547, 132)
        Me.txtIDDomu.Name = "txtIDDomu"
        Me.txtIDDomu.Size = New System.Drawing.Size(53, 20)
        Me.txtIDDomu.TabIndex = 11
        Me.txtIDDomu.Text = "A"
        '
        'txtRejon
        '
        Me.txtRejon.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtRejon.ForeColor = System.Drawing.Color.Purple
        Me.txtRejon.Location = New System.Drawing.Point(876, 132)
        Me.txtRejon.Name = "txtRejon"
        Me.txtRejon.Size = New System.Drawing.Size(112, 20)
        Me.txtRejon.TabIndex = 12
        '
        'txtNrDomu
        '
        Me.txtNrDomu.ForeColor = System.Drawing.Color.Purple
        Me.txtNrDomu.Location = New System.Drawing.Point(491, 132)
        Me.txtNrDomu.Name = "txtNrDomu"
        Me.txtNrDomu.Size = New System.Drawing.Size(53, 20)
        Me.txtNrDomu.TabIndex = 10
        Me.txtNrDomu.Text = "0"
        '
        'txtMiejscowosc
        '
        Me.txtMiejscowosc.AcceptsReturn = True
        Me.txtMiejscowosc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMiejscowosc.ForeColor = System.Drawing.Color.Purple
        Me.txtMiejscowosc.Location = New System.Drawing.Point(304, 108)
        Me.txtMiejscowosc.Name = "txtMiejscowosc"
        Me.txtMiejscowosc.Size = New System.Drawing.Size(741, 20)
        Me.txtMiejscowosc.TabIndex = 8
        '
        'txtNazwa1
        '
        Me.txtNazwa1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNazwa1.ForeColor = System.Drawing.Color.Purple
        Me.txtNazwa1.Location = New System.Drawing.Point(112, 84)
        Me.txtNazwa1.Name = "txtNazwa1"
        Me.txtNazwa1.Size = New System.Drawing.Size(933, 20)
        Me.txtNazwa1.TabIndex = 6
        '
        'txtKtoWpisal
        '
        Me.txtKtoWpisal.ForeColor = System.Drawing.Color.Purple
        Me.txtKtoWpisal.Location = New System.Drawing.Point(491, 56)
        Me.txtKtoWpisal.Name = "txtKtoWpisal"
        Me.txtKtoWpisal.Size = New System.Drawing.Size(109, 20)
        Me.txtKtoWpisal.TabIndex = 22
        '
        'txtUlica
        '
        Me.txtUlica.ForeColor = System.Drawing.Color.Purple
        Me.txtUlica.Location = New System.Drawing.Point(112, 132)
        Me.txtUlica.Name = "txtUlica"
        Me.txtUlica.Size = New System.Drawing.Size(328, 20)
        Me.txtUlica.TabIndex = 9
        '
        'txtTofi
        '
        Me.txtTofi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTofi.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtTofi.ForeColor = System.Drawing.Color.Purple
        Me.txtTofi.Location = New System.Drawing.Point(437, 8)
        Me.txtTofi.Name = "txtTofi"
        Me.txtTofi.Size = New System.Drawing.Size(608, 20)
        Me.txtTofi.TabIndex = 4
        '
        'txteMail
        '
        Me.txteMail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txteMail.ForeColor = System.Drawing.Color.Purple
        Me.txteMail.Location = New System.Drawing.Point(112, 204)
        Me.txteMail.Name = "txteMail"
        Me.txteMail.Size = New System.Drawing.Size(888, 20)
        Me.txteMail.TabIndex = 15
        '
        'txtFax
        '
        Me.txtFax.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFax.ForeColor = System.Drawing.Color.Purple
        Me.txtFax.Location = New System.Drawing.Point(112, 180)
        Me.txtFax.Name = "txtFax"
        Me.txtFax.Size = New System.Drawing.Size(933, 20)
        Me.txtFax.TabIndex = 14
        '
        'txtTelefon
        '
        Me.txtTelefon.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTelefon.ForeColor = System.Drawing.Color.Purple
        Me.txtTelefon.Location = New System.Drawing.Point(112, 156)
        Me.txtTelefon.Name = "txtTelefon"
        Me.txtTelefon.Size = New System.Drawing.Size(933, 20)
        Me.txtTelefon.TabIndex = 13
        '
        'txtNIP
        '
        Me.txtNIP.ForeColor = System.Drawing.Color.Purple
        Me.txtNIP.Location = New System.Drawing.Point(112, 32)
        Me.txtNIP.Name = "txtNIP"
        Me.txtNIP.Size = New System.Drawing.Size(112, 20)
        Me.txtNIP.TabIndex = 5
        '
        'txtKodP
        '
        Me.txtKodP.ForeColor = System.Drawing.Color.Purple
        Me.txtKodP.Location = New System.Drawing.Point(112, 108)
        Me.txtKodP.Name = "txtKodP"
        Me.txtKodP.Size = New System.Drawing.Size(112, 20)
        Me.txtKodP.TabIndex = 7
        '
        'txtIdent
        '
        Me.txtIdent.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtIdent.ForeColor = System.Drawing.Color.Purple
        Me.txtIdent.Location = New System.Drawing.Point(112, 8)
        Me.txtIdent.Name = "txtIdent"
        Me.txtIdent.Size = New System.Drawing.Size(112, 20)
        Me.txtIdent.TabIndex = 1
        Me.txtIdent.Text = "000000"
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(1007, 203)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(38, 23)
        Me.Button1.TabIndex = 18
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdNrKarty
        '
        Me.cmdNrKarty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdNrKarty.ForeColor = System.Drawing.Color.Black
        Me.cmdNrKarty.Location = New System.Drawing.Point(286, 7)
        Me.cmdNrKarty.Name = "cmdNrKarty"
        Me.cmdNrKarty.Size = New System.Drawing.Size(50, 23)
        Me.cmdNrKarty.TabIndex = 3
        Me.cmdNrKarty.Tag = "Wybier Identyfikator klienta z listy niewykorzystanych"
        Me.cmdNrKarty.Text = "Wolny"
        '
        'Label26
        '
        Me.Label26.ForeColor = System.Drawing.Color.Navy
        Me.Label26.Location = New System.Drawing.Point(772, 132)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(112, 16)
        Me.Label26.TabIndex = 70
        Me.Label26.Text = "Rejon klienta . . . . . . . . ."
        '
        'cmdKtoWpisal
        '
        Me.cmdKtoWpisal.Image = CType(resources.GetObject("cmdKtoWpisal.Image"), System.Drawing.Image)
        Me.cmdKtoWpisal.Location = New System.Drawing.Point(605, 55)
        Me.cmdKtoWpisal.Name = "cmdKtoWpisal"
        Me.cmdKtoWpisal.Size = New System.Drawing.Size(32, 23)
        Me.cmdKtoWpisal.TabIndex = 23
        '
        'Label33
        '
        Me.Label33.Location = New System.Drawing.Point(446, 132)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(72, 16)
        Me.Label33.TabIndex = 46
        Me.Label33.Text = "Nr domu . . . . . . . . . . . . "
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(232, 108)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(112, 16)
        Me.Label12.TabIndex = 45
        Me.Label12.Text = "Miejscowoœæ . . . . . . . . . . . . "
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 84)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(112, 16)
        Me.Label14.TabIndex = 43
        Me.Label14.Text = "Nazwa klienta . . . . . . . . ."
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(377, 58)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(168, 16)
        Me.Label2.TabIndex = 42
        Me.Label2.Text = "Kto wpisa³ do Bazy . . . . . "
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(377, 11)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(112, 16)
        Me.Label16.TabIndex = 34
        Me.Label16.Text = "Nr TOFI . . . . . "
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(8, 204)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(112, 16)
        Me.Label9.TabIndex = 21
        Me.Label9.Text = "Adres eMail . . . . . . ."
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 180)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 16)
        Me.Label8.TabIndex = 20
        Me.Label8.Text = "Fax . . . . . . . . . . .. . "
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 156)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 16)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "Telefon . . . . . . . . . . . ."
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 32)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 16)
        Me.Label6.TabIndex = 18
        Me.Label6.Text = "NIP klienta . . . . . ."
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 132)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 16)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Ulica . . . . . . . . . . . . . . "
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 108)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Kod pocztowy . . . . . . . . . . . . "
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Ident. / Nr karty . . . . . . . . ."
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.Identyfikacja)
        Me.TabControl1.Controls.Add(Me.Opiekun)
        Me.TabControl1.Controls.Add(Me.Charakterystyka)
        Me.TabControl1.Controls.Add(Me.UslugiDod)
        Me.TabControl1.Controls.Add(Me.AnalizyZakupu)
        Me.TabControl1.Controls.Add(Me.ObrotyMies)
        Me.TabControl1.Controls.Add(Me.Obroty)
        Me.TabControl1.Controls.Add(Me.AkcjeSpec)
        Me.TabControl1.Location = New System.Drawing.Point(8, 8)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1062, 671)
        Me.TabControl1.TabIndex = 0
        '
        'lblKlientNieaktywny
        '
        Me.lblKlientNieaktywny.BackColor = System.Drawing.Color.Transparent
        Me.lblKlientNieaktywny.Font = New System.Drawing.Font("Tahoma", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblKlientNieaktywny.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblKlientNieaktywny.Location = New System.Drawing.Point(770, -3)
        Me.lblKlientNieaktywny.Name = "lblKlientNieaktywny"
        Me.lblKlientNieaktywny.Size = New System.Drawing.Size(300, 30)
        Me.lblKlientNieaktywny.TabIndex = 532
        Me.lblKlientNieaktywny.Text = "Klient NIEAKTYWNY"
        Me.lblKlientNieaktywny.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'EdytujKatalogKlientow
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1075, 712)
        Me.Controls.Add(Me.lblKlientNieaktywny)
        Me.Controls.Add(Me.cmdCoDalej)
        Me.Controls.Add(Me.btnNast)
        Me.Controls.Add(Me.btnPop)
        Me.Controls.Add(Me.cmdWysylaj)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.cmdWycofajSie)
        Me.Controls.Add(Me.cmdPrzywroc)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdKontaktPromotor)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EdytujKatalogKlientow"
        Me.ShowInTaskbar = False
        Me.Text = "Edytuj informacje o kliencie"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuCoDalej.ResumeLayout(False)
        Me.MenuKontakty.ResumeLayout(False)
        Me.MenuRZUKontakty.ResumeLayout(False)
        Me.ContextMenuStrip1.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuOsoby.ResumeLayout(False)
        Me.AkcjeSpec.ResumeLayout(False)
        Me.PanelAkcjeSpec.ResumeLayout(False)
        CType(Me.dagAkcjeSpec, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuAkcjeSpec.ResumeLayout(False)
        Me.Obroty.ResumeLayout(False)
        CType(Me.dagObroty, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuObroty.ResumeLayout(False)
        Me.ObrotyMies.ResumeLayout(False)
        CType(Me.dagObrotyMies, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuObrotyMies.ResumeLayout(False)
        Me.AnalizyZakupu.ResumeLayout(False)
        CType(Me.dagAnalizyZakupu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menuAnalizyZakupu.ResumeLayout(False)
        CType(Me.dagOsoby, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UslugiDod.ResumeLayout(False)
        Me.PanelUslugiDod.ResumeLayout(False)
        Me.PanelUslugiDod.PerformLayout()
        CType(Me.dagRZUKontakty, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Charakterystyka.ResumeLayout(False)
        Me.Charakterystyka.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Opiekun.ResumeLayout(False)
        Me.PanelSprzedaz.ResumeLayout(False)
        Me.PanelSprzedaz.PerformLayout()
        Me.SplitContainer6.Panel1.ResumeLayout(False)
        Me.SplitContainer6.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer6.ResumeLayout(False)
        Me.Panel9.ResumeLayout(False)
        CType(Me.dagCoDalej, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dagKontakty, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Identyfikacja.ResumeLayout(False)
        Me.Identyfikacja.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim RowHeightSetterKontakty As Softart.myDataGridRowHeightSetter
    Dim RowHeightSetterCoDalej As Softart.myDataGridRowHeightSetter
    Dim pointInCell00 As System.Drawing.Point
    Dim ScrollXOffset As Integer = 0
    Dim _OCoMamRobic As String = oCoMamRobic
    Dim _oOperacja As String = oOperacja
    Dim HorizScrollLastTime As String = ""



    Private StartFormy As Boolean = True
    Private StartFormyOpiekun As Boolean = True
    Private StartFormyCharakterystyka As Boolean = True
    Private StartFormyRZUKontakty As Boolean = True
    Private StartFormyOsoby As Boolean = True
    Private StartFormyAnalizyZakupu As Boolean = True
    Private StartFormyObrotyMies As Boolean = True
    Private StartFormyObroty As Boolean = True
    Private StartFormyAkcjeSpec As Boolean = True

    '*********************************************************
    '** Kontakty
    '*********************************************************

    'Dim dbSelectSQLKontakty As String = sSelectSQLKontakty &
    '                                    " WHERE (IDENTKLIENTA='" & oIdentKli & "') and (IDENTKLIENTA<>'') ORDER BY DATAKONTAKTU DESC"
    Dim dbSelectSQLKontakty As String = ""

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

    Dim DataSetKontakty As New DataSet
    Dim DataViewKontakty As New DataView


    '*********************************************************
    '** RZUKontakty
    '*********************************************************

    'Dim dbSelectSQLRZUKontakty As String = sSelectSQLKontakty &
    '                                    " WHERE (IDENTKLIENTA='" & oIdentKli & "') and (IDENTKLIENTA<>'') ORDER BY DATAKONTAKTU DESC"
    Dim dbSelectSQLRZUKontakty As String = ""

    Dim dbConnectionRZUKontakty As OleDb.OleDbConnection
    Dim dbCommandSelectRZUKontakty As OleDb.OleDbCommand
    Dim dbCommandDeleteRZUKontakty As OleDb.OleDbCommand
    Dim dbCommandUpdateRZUKontakty As OleDb.OleDbCommand
    Dim dbCommandInsertRZUKontakty As OleDb.OleDbCommand
    Dim dbDataAdapterRZUKontakty As OleDb.OleDbDataAdapter

    Dim sqlConnectionRZUKontakty As SqlClient.SqlConnection
    Dim sqlCommandSelectRZUKontakty As SqlClient.SqlCommand
    Dim sqlCommandDeleteRZUKontakty As SqlClient.SqlCommand
    Dim sqlCommandUpdateRZUKontakty As SqlClient.SqlCommand
    Dim sqlCommandInsertRZUKontakty As SqlClient.SqlCommand
    Dim sqlDataAdapterRZUKontakty As SqlClient.SqlDataAdapter

    Dim DataSetRZUKontakty As New DataSet
    Dim DataViewRZUKontakty As New DataView


    '*********************************************************
    '** CoDalej
    '*********************************************************

    'Dim dbSelectSQLCoDalej As String = sSelectSQLCoDalej &
    '                                    " WHERE (IDENTKLIENTA='" & oIdentKli & "') and (IDENTKLIENTA<>'') ORDER BY IDENT DESC"
    Dim dbSelectSQLCoDalej As String = ""

    Dim dbConnectionCoDalej As OleDb.OleDbConnection
    Dim dbCommandSelectCoDalej As OleDb.OleDbCommand
    Dim dbCommandDeleteCoDalej As OleDb.OleDbCommand
    Dim dbCommandUpdateCoDalej As OleDb.OleDbCommand
    Dim dbCommandInsertCoDalej As OleDb.OleDbCommand
    Dim dbDataAdapterCoDalej As OleDb.OleDbDataAdapter

    Dim sqlConnectionCoDalej As SqlClient.SqlConnection
    Dim sqlCommandSelectCoDalej As SqlClient.SqlCommand
    Dim sqlCommandDeleteCoDalej As SqlClient.SqlCommand
    Dim sqlCommandUpdateCoDalej As SqlClient.SqlCommand
    Dim sqlCommandInsertCoDalej As SqlClient.SqlCommand
    Dim sqlDataAdapterCoDalej As SqlClient.SqlDataAdapter

    Dim DataSetCoDalej As New DataSet
    Dim DataViewCoDalej As New DataView


    '*********************************************************
    '** Osoby kontaktowe
    '*********************************************************

    'Dim dbSelectSQLOsoby As String = "SELECT * FROM Osoby WHERE IDENTKLIENTA='" & oIdentKli & "'"
    Dim dbSelectSQLOsoby As String = ""

    Dim dbConnectionOsoby As OleDb.OleDbConnection
    Dim dbCommandSelectOsoby As OleDb.OleDbCommand
    Dim dbCommandDeleteOsoby As OleDb.OleDbCommand
    Dim dbCommandUpdateOsoby As OleDb.OleDbCommand
    Dim dbCommandInsertOsoby As OleDb.OleDbCommand
    Dim dbDataAdapterOsoby As OleDb.OleDbDataAdapter

    Dim sqlConnectionOsoby As SqlClient.SqlConnection
    Dim sqlCommandSelectOsoby As SqlClient.SqlCommand
    Dim sqlCommandDeleteOsoby As SqlClient.SqlCommand
    Dim sqlCommandUpdateOsoby As SqlClient.SqlCommand
    Dim sqlCommandInsertOsoby As SqlClient.SqlCommand
    Dim sqlDataAdapterOsoby As SqlClient.SqlDataAdapter

    Dim DataSetOsoby As New DataSet
    Dim DataViewOsoby As New DataView

    '*********************************************************
    '** Obroty w miesiacu
    '*********************************************************
    'UWAGA : do sumy iloœci/wartoœci/mar¿y nie doliczamy SKUPU 
    Dim baseSelectSQLObrotyMies As String = "SELECT " &
                                              "ObrotyMies.IDENTKLIENTA," &
                                              "Klienci.NRTOFITXT," &
                                              "ObrotyMies.MIESIAC," &
                                                 "ISNULL(ObrotyMies.AILOSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.AWARTSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.AMARSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.LILOSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.LWARTSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.LMARSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.AOILOSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.AOWARTSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.AOMARSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.LOILOSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.LOWARTSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.LOMARSPRZED,0)," &
                                                   "ISNULL(ObrotyMies.NAJEMILOSPRZED,0)," &
                                                   "ISNULL(ObrotyMies.NAJEMWARTSPRZED,0)," &
                                                   "ISNULL(ObrotyMies.NAJEMMARSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.STRONYILOSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.STRONYWARTSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.STRONYMARSPRZED,0)," &
                                                   "ISNULL(ObrotyMies.DRUKARKIILOSPRZED,0)," &
                                                   "ISNULL(ObrotyMies.DRUKARKIWARTSPRZED,0)," &
                                                   "ISNULL(ObrotyMies.DRUKARKIMARSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.SKUPILOSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.SKUPWARTSPRZED,0)," &
                                                 "ISNULL(ObrotyMies.SKUPMARSPRZED,0)," &
                                                   "ISNULL(ObrotyMies.AILOSPRZED,0) + " &
                                                   "ISNULL(ObrotyMies.LILOSPRZED,0) + " &
                                                   "ISNULL(ObrotyMies.AOILOSPRZED,0) + " &
                                                   "ISNULL(ObrotyMies.LOILOSPRZED,0) + " &
                                                   "ISNULL(ObrotyMies.NAJEMILOSPRZED,0) + " &
                                                   "ISNULL(ObrotyMies.STRONYILOSPRZED,0) + " &
                                                   "ISNULL(ObrotyMies.DRUKARKIILOSPRZED,0) AS SUMAILO, " &
                                                         "ISNULL(ObrotyMies.AWARTSPRZED,0) + " &
                                                         "ISNULL(ObrotyMies.LWARTSPRZED,0) + " &
                                                         "ISNULL(ObrotyMies.AOWARTSPRZED,0) + " &
                                                         "ISNULL(ObrotyMies.LOWARTSPRZED,0) + " &
                                                         "ISNULL(ObrotyMies.NAJEMWARTSPRZED,0) + " &
                                                         "ISNULL(ObrotyMies.STRONYWARTSPRZED,0) + " &
                                                         "ISNULL(ObrotyMies.DRUKARKIWARTSPRZED,0) AS SUMAWART, " &
                                                              "ISNULL(ObrotyMies.AMARSPRZED,0) + " &
                                                              "ISNULL(ObrotyMies.LMARSPRZED,0) + " &
                                                              "ISNULL(ObrotyMies.AOMARSPRZED,0) + " &
                                                              "ISNULL(ObrotyMies.LOMARSPRZED,0) + " &
                                                              "ISNULL(ObrotyMies.NAJEMMARSPRZED,0) + " &
                                                              "ISNULL(ObrotyMies.STRONYMARSPRZED,0) + " &
                                                              "ISNULL(ObrotyMies.DRUKARKIMARSPRZED,0) AS SUMAMAR, " &
                                                 "ObrotyMies.WERSJA " &
                                            "FROM ObrotyMies INNER JOIN Klienci ON ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA "
    Dim dbSelectSQLObrotyMies As String = ""

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

    Dim DataSetObrotyMies As New DataSet
    Dim DataViewObrotyMies As New DataView
    Dim ConnectionState As System.Data.ConnectionState


    '*********************************************************
    '** Analizy zakupu
    '*********************************************************

    Dim dbSelectSQLAnalizyZakupu As String = sSelectSQLAnalizyZakupu & " WHERE IDENTKLIENTA='" & oIdentKli & "'"

    Dim dbConnectionAnalizyZakupu As OleDb.OleDbConnection
    Dim dbCommandSelectAnalizyZakupu As OleDb.OleDbCommand
    Dim dbCommandDeleteAnalizyZakupu As OleDb.OleDbCommand
    Dim dbCommandUpdateAnalizyZakupu As OleDb.OleDbCommand
    Dim dbCommandInsertAnalizyZakupu As OleDb.OleDbCommand
    Dim dbDataAdapterAnalizyZakupu As OleDb.OleDbDataAdapter

    Dim sqlConnectionAnalizyZakupu As SqlClient.SqlConnection
    Dim sqlCommandSelectAnalizyZakupu As SqlClient.SqlCommand
    Dim sqlCommandDeleteAnalizyZakupu As SqlClient.SqlCommand
    Dim sqlCommandUpdateAnalizyZakupu As SqlClient.SqlCommand
    Dim sqlCommandInsertAnalizyZakupu As SqlClient.SqlCommand
    Dim sqlDataAdapterAnalizyZakupu As SqlClient.SqlDataAdapter

    Dim DataSetAnalizyZakupu As New DataSet
    Dim DataViewAnalizyZakupu As New DataView


    '*********************************************************
    '** Obroty
    '*********************************************************

    Dim baseSelectSQLObroty As String = "SELECT " &
                                           "Obroty.IDENTKLIENTA," &
                                           "Klienci.NRTOFITXT," &
                                           "Obroty.DATA," &
                                                 "ISNULL(Obroty.AILOSPRZED,0)," &
                                                 "ISNULL(Obroty.AWARTSPRZED,0)," &
                                                 "ISNULL(Obroty.AMARSPRZED,0)," &
                                                 "ISNULL(Obroty.LILOSPRZED,0)," &
                                                 "ISNULL(Obroty.LWARTSPRZED,0)," &
                                                 "ISNULL(Obroty.LMARSPRZED,0)," &
                                                 "ISNULL(Obroty.AOILOSPRZED,0)," &
                                                 "ISNULL(Obroty.AOWARTSPRZED,0)," &
                                                 "ISNULL(Obroty.AOMARSPRZED,0)," &
                                                 "ISNULL(Obroty.LOILOSPRZED,0)," &
                                                 "ISNULL(Obroty.LOWARTSPRZED,0)," &
                                                 "ISNULL(Obroty.LOMARSPRZED,0)," &
                                                   "ISNULL(Obroty.NAJEMILOSPRZED,0)," &
                                                   "ISNULL(Obroty.NAJEMWARTSPRZED,0)," &
                                                   "ISNULL(Obroty.NAJEMMARSPRZED,0)," &
                                                 "ISNULL(Obroty.STRONYILOSPRZED,0)," &
                                                 "ISNULL(Obroty.STRONYWARTSPRZED,0)," &
                                                 "ISNULL(Obroty.STRONYMARSPRZED,0)," &
                                                   "ISNULL(Obroty.DRUKARKIILOSPRZED,0)," &
                                                   "ISNULL(Obroty.DRUKARKIWARTSPRZED,0)," &
                                                   "ISNULL(Obroty.DRUKARKIMARSPRZED,0)," &
                                                 "ISNULL(Obroty.SKUPILOSPRZED,0)," &
                                                 "ISNULL(Obroty.SKUPWARTSPRZED,0)," &
                                                 "ISNULL(Obroty.SKUPMARSPRZED,0)," &
                                                   "ISNULL(Obroty.AILOSPRZED,0) + " &
                                                   "ISNULL(Obroty.LILOSPRZED,0) + " &
                                                   "ISNULL(Obroty.AOILOSPRZED,0) + " &
                                                   "ISNULL(Obroty.LOILOSPRZED,0) + " &
                                                   "ISNULL(Obroty.NAJEMILOSPRZED,0) + " &
                                                   "ISNULL(Obroty.STRONYILOSPRZED,0) + " &
                                                   "ISNULL(Obroty.DRUKARKIILOSPRZED,0) + " &
                                                   "ISNULL(Obroty.SKUPILOSPRZED,0) AS SUMAILO, " &
                                                         "ISNULL(Obroty.AWARTSPRZED,0) + " &
                                                         "ISNULL(Obroty.LWARTSPRZED,0) + " &
                                                         "ISNULL(Obroty.AOWARTSPRZED,0) + " &
                                                         "ISNULL(Obroty.LOWARTSPRZED,0) + " &
                                                         "ISNULL(Obroty.NAJEMWARTSPRZED,0) + " &
                                                         "ISNULL(Obroty.STRONYWARTSPRZED,0) + " &
                                                         "ISNULL(Obroty.DRUKARKIWARTSPRZED,0) + " &
                                                         "ISNULL(Obroty.SKUPWARTSPRZED,0) AS SUMAWART, " &
                                                              "ISNULL(Obroty.AMARSPRZED,0) + " &
                                                              "ISNULL(Obroty.LMARSPRZED,0) + " &
                                                              "ISNULL(Obroty.AOMARSPRZED,0) + " &
                                                              "ISNULL(Obroty.LOMARSPRZED,0) + " &
                                                              "ISNULL(Obroty.NAJEMMARSPRZED,0) + " &
                                                              "ISNULL(Obroty.STRONYMARSPRZED,0) + " &
                                                              "ISNULL(Obroty.DRUKARKIMARSPRZED,0) + " &
                                                              "ISNULL(Obroty.SKUPMARSPRZED,0) AS SUMAMAR, " &
                                                 "Obroty.WERSJA " &
                                            "FROM Obroty INNER JOIN Klienci ON Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA "

    Dim dbSelectSQLObroty As String = ""

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

    Dim DataSetObroty As New DataSet
    Dim DataViewObroty As New DataView

    '*********************************************************
    '** Akcje
    '*********************************************************

    'Dim dbSelectSQLAkcjeSpec As String = "SELECT AkcjeSpec.IDENTKLI, " &
    '                                           "AkcjeSpec.IDENTAKCJI, " &
    '                                           "AkcjeOpis.DATA, " &
    '                                           "AkcjeOpis.OPIS " &
    '                                           "FROM AkcjeSpec INNER JOIN AkcjeOpis " &
    '                                           "ON AkcjeSpec.IDENTAKCJI=AkcjeOpis.IDENT " &
    '                                           "WHERE AkcjeSpec.IDENTKLI='" & oIdentKli & "'"
    Dim dbSelectSQLAkcjeSpec As String = ""

    Dim dbConnectionAkcjeSpec As OleDb.OleDbConnection
    Dim dbCommandSelectAkcjeSpec As OleDb.OleDbCommand
    'Dim dbCommandDeleteAkcjeSpec As OleDb.OleDbCommand
    'Dim dbCommandUpdateAkcjeSpec As OleDb.OleDbCommand
    'Dim dbCommandInsertAkcjeSpec As OleDb.OleDbCommand
    Dim dbDataAdapterAkcjeSpec As OleDb.OleDbDataAdapter

    Dim sqlConnectionAkcjeSpec As SqlClient.SqlConnection
    Dim sqlCommandSelectAkcjeSpec As SqlClient.SqlCommand
    'Dim sqlCommandDeleteAkcjeSpec As SqlClient.SqlCommand
    'Dim sqlCommandUpdateAkcjeSpec As SqlClient.SqlCommand
    'Dim sqlCommandInsertAkcjeSpec As SqlClient.SqlCommand
    Dim sqlDataAdapterAkcjeSpec As SqlClient.SqlDataAdapter

    Dim DataSetAkcjeSpec As New DataSet
    Dim DataViewAkcjeSpec As New DataView


    Dim Okres1 As String = Mid(Par_OstOkresAnalizy, 1, 7)
    Dim Okres2 As String = Mid(Par_OstOkresAnalizy, 8, 7)







    '====================================================
    'UWAGA - trzeba nadpisaæ metode WinProc tej formy
    'zdarzenie kontroluje sposob zamkniecia formy
    '====================================================
    Private Sub EdytujKatalogKlientow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me._closeClick Then
            If Len(Trim(txtIdent.Text)) = 0 Then
                MessageBox.Show("Nie mogê zapisaæ informacji o kliencie bez definicji identyfikatora klienta." & vbCrLf &
                "Proszê zdefiniowaæ identyfikator klienta.",
                "Brak identyfikatora klienta.",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation)
                e.Cancel = True     'nie pozwalaj zamkn¹c forme
            Else
                ZapiszDane()
                oCoDalej = "STÓJ"
                oAktualizuj = True
                'Me.Close()
            End If







            ''e.Cancel = True     'nie pozwalaj zamkn¹c forme
            'cmdPowrot_Click(sender, e)

            'If MessageBox.Show("Zdecydowa³eœ opuœciæ formê bez zapisania wprowadzonych zmian." & vbCrLf & _
            '    "Czy mam zapisaæ zmiany w Bazie Danych ?", _
            '    "Uwaga :", _
            '    System.Windows.Forms.MessageBoxButtons.YesNo, _
            '    MessageBoxIcon.Question, _
            '    MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then

            '    ZapiszDane()
            '    oAktualizuj = True
            'Else
            '    oAktualizuj = False
            'End If
        Else
            'MsgBox("ZAPISZ lub WYCOFAJ SIÊ")
        End If
    End Sub
    '====================================================








    '*********************************************************
    '** Inicjowanie
    '*********************************************************

    Private Sub EdytujKatalogKlientow_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        If oAktywnyKli Then
            RysujGradient(Me)
        Else
            RysujGradientC(Me)
        End If
    End Sub


    Private Sub EdytujKatalogKlientow_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True   'ustaw siê na szczycie
        'Me.TopMost = False  'pozwol innym tez byc na szczycie...
        'TabControl1.TabPages.RemoveAt(9)

        ''80%wysokosci i 80%szerokosci
        'Me.Size = New Size(Int(0.8 * Screen.PrimaryScreen.Bounds.Width), Int(0.8 * Screen.PrimaryScreen.Bounds.Height))
        ''wyœrodkuj
        'Me.Location = New System.Drawing.Point(Int((Screen.PrimaryScreen.Bounds.Width - Me.Size.Width) / 2),
        '                                       Int((Screen.PrimaryScreen.Bounds.Height - Me.Size.Height) / 2))
        'Me.Visible = False



        cmdKontaktPromotor.Visible = False
        cmdCoDalej.Visible = False

        'inicjowanie...
        PobierzDane()
        Select Case _oOperacja
            Case "EDYTUJ", "Edytuj", "Poka¿"
                'nie pozwalaj na edycja PrimaryKey
                txtIdent.Enabled = False
                btnPop.Visible = True
                btnNast.Visible = True
            Case Else
                txtIdent.Enabled = True
                btnPop.Visible = False
                btnNast.Visible = False
        End Select

        InicjujListeLat(cbxSprzedazOKontR)
        InicjujListeLat(cbxSprzedazNKontR)

        InicjujListeDniTygodnia(cbxSprzedazOKonD)
        InicjujListeDniTygodnia(cbxSprzedazNKonD)

        InicjujSposobWyboruDostawcy(cbxSposobWyboruDostawcy)

        Identyfikacja.Visible = False
        If Not oAktywnyKli Then
            lblKlientNieaktywny.Visible = True
            lblKlientNieaktywny.Location = New System.Drawing.Point(8, 8)

            TabControl1.Size = New System.Drawing.Size(TabControl1.Size.Width, TabControl1.Size.Height - 30)
            TabControl1.Location = New System.Drawing.Point(8, 8 + 30)
        Else
            lblKlientNieaktywny.Visible = False
        End If


        If _MaUprawnieniaAdministratora() Or _MaUprawnieniaSzefa() Then
            cmdPKontakt.Visible = True
            txtPKontakt.Enabled = True
        Else
            cmdPKontakt.Visible = False
            txtPKontakt.Enabled = False
        End If



        'TabControl1_SelectedIndexChanged(sender, e)
        TabControl1.SelectedIndex = 1
        StartFormy = True
        'TabControl1_SelectedIndexChanged(sender, e)
        Identyfikacja.Visible = True






        'Me.Visible = True
    End Sub

    Private Sub EdytujKatalogKlientow_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If StartFormy Then
            'uniemo¿liwiæ edycje pól we wszystkich kontenerach
            Select Case _oOperacja
                Case "Poka¿"
                    If TabControl1.TabCount >= 1 Then NieEdytowac(Me.TabControl1.TabPages(0))
                    If TabControl1.TabCount >= 2 Then NieEdytowac(Me.TabControl1.TabPages(1))
                    If TabControl1.TabCount >= 3 Then NieEdytowac(Me.TabControl1.TabPages(2))
                    If TabControl1.TabCount >= 4 Then NieEdytowac(Me.TabControl1.TabPages(3))
                    If TabControl1.TabCount >= 5 Then NieEdytowac(Me.TabControl1.TabPages(4))
                    If TabControl1.TabCount >= 6 Then NieEdytowac(Me.TabControl1.TabPages(5))
                    If TabControl1.TabCount >= 7 Then NieEdytowac(Me.TabControl1.TabPages(6))
                    If TabControl1.TabCount >= 8 Then NieEdytowac(Me.TabControl1.TabPages(7))
                    If TabControl1.TabCount >= 9 Then NieEdytowac(Me.TabControl1.TabPages(8))
                    If TabControl1.TabCount >= 10 Then NieEdytowac(Me.TabControl1.TabPages(9))
                    If TabControl1.TabCount >= 11 Then NieEdytowac(Me.TabControl1.TabPages(10))
                    If TabControl1.TabCount >= 12 Then NieEdytowac(Me.TabControl1.TabPages(11))
                    If TabControl1.TabCount >= 13 Then NieEdytowac(Me.TabControl1.TabPages(12))
                    If TabControl1.TabCount >= 14 Then NieEdytowac(Me.TabControl1.TabPages(13))
                    If TabControl1.TabCount >= 15 Then NieEdytowac(Me.TabControl1.TabPages(14))
                    If TabControl1.TabCount >= 16 Then NieEdytowac(Me.TabControl1.TabPages(15))
                    If TabControl1.TabCount >= 17 Then NieEdytowac(Me.TabControl1.TabPages(16))
                    If TabControl1.TabCount >= 18 Then NieEdytowac(Me.TabControl1.TabPages(17))
                    If TabControl1.TabCount >= 19 Then NieEdytowac(Me.TabControl1.TabPages(18))
                    If TabControl1.TabCount >= 20 Then NieEdytowac(Me.TabControl1.TabPages(19))
                Case Else
            End Select

            ''ustawienie kursora
            'Select Case _CoRobimy
            '    Case "ZMIANA VAT"
            '        'Umo¿liwienie edycji tylko stawki VAT
            '        cbxVAT.Visible = True
            '        cbxVAT.Enabled = True
            '        cbxVAT.Focus()
            '    Case "ZMIANA DOSTAWCY"
            '        'Umo¿liwienie edycji tylko stawki VAT
            '        cmdDostawca.Visible = True
            '        cmdDostawca.Enabled = True
            '        txtDostawca.Visible = True
            '        txtDostawca.Enabled = True
            '        txtDostawca.ReadOnly = False
            '        lblDostawca.Visible = True
            '        lblDostawca.Enabled = True
            '        txtDostawca.Focus()
            '    Case "ZMIANA STANUMIN"
            '        'Umo¿liwienie edycji tylko stawki VAT
            '        txtStanMin.Visible = True
            '        txtStanMin.Enabled = True
            '        txtStanMin.ReadOnly = False
            '        txtStanMin.Focus()

            '    Case "POKA¯"
            '        cmdPowrot.Focus()

            '    Case Else
            '        'edycja pozycji cennika
            '        'ustaw na pocz¹tek na górze....
            '        cmdPowrot.Focus()
            'End Select
            StartFormy = False
        Else
        End If

    End Sub


















    '*********************************************************
    '** Komendy ekranowe
    '*********************************************************
    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        If Len(Trim(txtIdent.Text)) = 0 Then
            MessageBox.Show("Nie mogê zapisaæ informacji o kliencie bez definicji identyfikatora klienta." & vbCrLf &
                "Proszê zdefiniowaæ identyfikator klienta.",
                "Brak identyfikatora klienta.",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation)
        Else
            ZapiszDane()
            oCoDalej = "STÓJ"
            oAktualizuj = True
            Me.Close()
        End If
    End Sub

    Private Sub cmdWycofajSie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWycofajSie.Click
        oCoDalej = "STÓJ"
        oAktualizuj = False
        Me.Close()
    End Sub

    Private Sub cmdPrzywroc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrzywroc.Click
        PobierzDane()
    End Sub

    Private Sub cmdPKontakt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPKontakt.Click
        oData = txtPKontakt.Text

        Dim Proc As New WybierzDate
        Proc.ShowDialog()

        txtPKontakt.Text = oData
    End Sub


    Private Sub cmdSprzedazOKonT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSprzedazOKonT.Click
        oData = ""
        oWeek = Val(txtSprzedazOKonT.Text)
        Dim Proc As New WybierzDate
        Proc.ShowDialog()
        cbxSprzedazOKontR.Text = Mid(oData, 1, 4)
        txtSprzedazOKonT.Text = Mid(Trim(Str(100 + oWeek)), 2, 2)
        cbxSprzedazOKonD.Text = DzienTygodnia(oData)
    End Sub

    Private Sub cmdSprzedazNKonT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSprzedazNKonT.Click
        oData = ""
        oWeek = Val(txtSprzedazNKonT.Text)
        Dim Proc As New WybierzDate
        Proc.ShowDialog()
        cbxSprzedazNKontR.Text = Mid(oData, 1, 4)
        txtSprzedazNKonT.Text = Mid(Trim(Str(100 + oWeek)), 2, 2)
        cbxSprzedazNKonD.Text = DzienTygodnia(oData)
    End Sub

    Private Sub cmdKtoWpisal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKtoWpisal.Click
        oCoMamRobic = "WYBIERAJ"
        oPracownik = Trim(txtKtoWpisal.Text)
        Dim form As New KatalogUzytkownikow
        form.ShowDialog()
        If Len(Trim(oPracownik)) > 0 Then
            txtKtoWpisal.Text = Trim(oPracownik)
        End If
    End Sub

    Private Sub cmdSprzedazOpiekun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSprzedazOpiekun.Click
        oCoMamRobic = "WYBIERAJ"
        oPracownik = Trim(txtSprzedazOpiekun.Text)
        Dim form As New KatalogUzytkownikow
        form.ShowDialog()
        If Len(Trim(oPracownik)) > 0 Then
            txtSprzedazOpiekun.Text = Trim(oPracownik)
        End If
    End Sub

    Private Sub cmdEdytujNumeryKartPayBack_Click(sender As Object, e As EventArgs) Handles cmdEdytujNumeryKartPayBack.Click
        'oCoMamRobic = "WYBIERAJ"
        oNumer = Trim(txtNryKartPayBack.Text)
        Dim form As New KatalogKartPayBack
        form.ShowDialog()
        If Len(Trim(oNumer)) > 0 Then
            txtNryKartPayBack.Text = oNumer
        End If
    End Sub





    Private Sub cmdWybierzPlik_Click(sender As Object, e As EventArgs) Handles cmdWybierzPlik.Click
        PobierzParametry()
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik z ostatnio z³o¿on¹ ofert¹..."
            .InitialDirectory = Par_KatalogOferty
            .DefaultExt = "*"
            .Filter = "Wszystkie pliki|*.*"
            .FileName = Trim(lblOfertaPlik.Text)

            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                lblOfertaPlik.Text = IO.Path.GetFileName(.FileName)
            End If
        End With
    End Sub

    Private Sub cmdWybierzR1_Click(sender As Object, e As EventArgs) Handles cmdWybierzR1.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik z raportem realizacji us³ugi..."
            .DefaultExt = "*"
            .Filter = "Wszystkie pliki|*.*"
            If Len(txtRaport01.Text) > 0 Then
                .InitialDirectory = IO.Path.GetFullPath(txtRaport01.Text)
                .FileName = IO.Path.GetFileName(txtRaport01.Text)
            Else
                .InitialDirectory = "c:\"
                .FileName = ""
            End If

            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtRaport01.Text = .FileName
            End If
        End With
    End Sub
    Private Sub cmdWybierzR2_Click(sender As Object, e As EventArgs) Handles cmdWybierzR2.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik z raportem realizacji us³ugi..."
            .DefaultExt = "*"
            .Filter = "Wszystkie pliki|*.*"
            If Len(txtRaport02.Text) > 0 Then
                .InitialDirectory = IO.Path.GetFullPath(txtRaport02.Text)
                .FileName = IO.Path.GetFileName(txtRaport02.Text)
            Else
                .InitialDirectory = "c:\"
                .FileName = ""
            End If

            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtRaport02.Text = .FileName
            End If
        End With
    End Sub
    Private Sub cmdWybierzR3_Click(sender As Object, e As EventArgs) Handles cmdWybierzR3.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik z raportem realizacji us³ugi..."
            .DefaultExt = "*"
            .Filter = "Wszystkie pliki|*.*"
            If Len(txtRaport03.Text) > 0 Then
                .InitialDirectory = IO.Path.GetFullPath(txtRaport03.Text)
                .FileName = IO.Path.GetFileName(txtRaport03.Text)
            Else
                .InitialDirectory = "c:\"
                .FileName = ""
            End If

            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtRaport03.Text = .FileName
            End If
        End With
    End Sub
    Private Sub cmdWybierzR4_Click(sender As Object, e As EventArgs) Handles cmdWybierzR4.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik z raportem realizacji us³ugi..."
            .DefaultExt = "*"
            .Filter = "Wszystkie pliki|*.*"
            If Len(txtRaport04.Text) > 0 Then
                .InitialDirectory = IO.Path.GetFullPath(txtRaport04.Text)
                .FileName = IO.Path.GetFileName(txtRaport04.Text)
            Else
                .InitialDirectory = "c:\"
                .FileName = ""
            End If

            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtRaport04.Text = .FileName
            End If
        End With
    End Sub
    Private Sub cmdWybierzR5_Click(sender As Object, e As EventArgs) Handles cmdWybierzR5.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik z raportem realizacji us³ugi..."
            .DefaultExt = "*"
            .Filter = "Wszystkie pliki|*.*"
            If Len(txtRaport05.Text) > 0 Then
                .InitialDirectory = IO.Path.GetFullPath(txtRaport05.Text)
                .FileName = IO.Path.GetFileName(txtRaport05.Text)
            Else
                .InitialDirectory = "c:\"
                .FileName = ""
            End If

            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtRaport05.Text = .FileName
            End If
        End With
    End Sub
    Private Sub cmdWybierzR6_Click(sender As Object, e As EventArgs) Handles cmdWybierzR6.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik z raportem realizacji us³ugi..."
            .DefaultExt = "*"
            .Filter = "Wszystkie pliki|*.*"
            If Len(txtRaport06.Text) > 0 Then
                .InitialDirectory = IO.Path.GetFullPath(txtRaport06.Text)
                .FileName = IO.Path.GetFileName(txtRaport06.Text)
            Else
                .InitialDirectory = "c:\"
                .FileName = ""
            End If

            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtRaport06.Text = .FileName
            End If
        End With
    End Sub
    Private Sub cmdWybierzR7_Click(sender As Object, e As EventArgs) Handles cmdWybierzR7.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik z raportem realizacji us³ugi..."
            .DefaultExt = "*"
            .Filter = "Wszystkie pliki|*.*"
            If Len(txtRaport07.Text) > 0 Then
                .InitialDirectory = IO.Path.GetFullPath(txtRaport07.Text)
                .FileName = IO.Path.GetFileName(txtRaport07.Text)
            Else
                .InitialDirectory = "c:\"
                .FileName = ""
            End If

            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtRaport07.Text = .FileName
            End If
        End With
    End Sub
    Private Sub cmdWybierzR8_Click(sender As Object, e As EventArgs) Handles cmdWybierzR8.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik z raportem realizacji us³ugi..."
            .DefaultExt = "*"
            .Filter = "Wszystkie pliki|*.*"
            If Len(txtRaport08.Text) > 0 Then
                .InitialDirectory = IO.Path.GetFullPath(txtRaport08.Text)
                .FileName = IO.Path.GetFileName(txtRaport08.Text)
            Else
                .InitialDirectory = "c:\"
                .FileName = ""
            End If

            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtRaport08.Text = .FileName
            End If
        End With
    End Sub
    Private Sub cmdWybierzR9_Click(sender As Object, e As EventArgs) Handles cmdWybierzR9.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik z raportem realizacji us³ugi..."
            .DefaultExt = "*"
            .Filter = "Wszystkie pliki|*.*"
            If Len(txtRaport09.Text) > 0 Then
                .InitialDirectory = IO.Path.GetFullPath(txtRaport09.Text)
                .FileName = IO.Path.GetFileName(txtRaport09.Text)
            Else
                .InitialDirectory = "c:\"
                .FileName = ""
            End If

            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtRaport09.Text = .FileName
            End If
        End With
    End Sub






    Private Sub cmdPoka¿Plik_Click(sender As Object, e As EventArgs) Handles cmdPoka¿Plik.Click
        Dim PlikDok As String = ""
        PobierzParametry()
        If Len(lblOfertaPlik.Text) > 0 Then
            PlikDok = Par_KatalogOferty & "\" & Trim(lblOfertaPlik.Text)

            ''przepisz dokument na dysk
            'Dim fs As New IO.FileStream(PlikDok, IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
            'Dim DlugoscPliku As Int64 = DZ_ZawartoscPliku.Length
            'fs.Write(DZ_ZawartoscPliku, 0, DlugoscPliku)
            'fs.Close()

            If My.Computer.FileSystem.FileExists(PlikDok) Then
                Shell("cmd /C""" & PlikDok & """", AppWinStyle.Hide, True)
            Else
                MessageBox.Show("Nie znalaz³em pliku z dokumentem" & vbCrLf & PlikDok,
                    "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End If

            'If IO.File.Exists(PlikDok) Then IO.File.Delete(PlikDok)
        End If
    End Sub


    Private Sub cmdPokazR1_Click(sender As Object, e As EventArgs) Handles cmdPokazR1.Click
        Dim PlikDok As String = ""
        PobierzParametry()
        If Len(txtRaport01.Text) > 0 Then
            PlikDok = Trim(txtRaport01.Text)
            If My.Computer.FileSystem.FileExists(PlikDok) Then
                Shell("cmd /C""" & PlikDok & """", AppWinStyle.Hide, True)
            Else
                MessageBox.Show("Nie znalaz³em pliku z dokumentem" & vbCrLf & PlikDok,
                    "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End If
        End If
    End Sub
    Private Sub cmdPokazR2_Click(sender As Object, e As EventArgs) Handles cmdPokazR2.Click
        Dim PlikDok As String = ""
        PobierzParametry()
        If Len(txtRaport02.Text) > 0 Then
            PlikDok = Trim(txtRaport02.Text)
            If My.Computer.FileSystem.FileExists(PlikDok) Then
                Shell("cmd /C""" & PlikDok & """", AppWinStyle.Hide, True)
            Else
                MessageBox.Show("Nie znalaz³em pliku z dokumentem" & vbCrLf & PlikDok,
                    "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End If
        End If
    End Sub
    Private Sub cmdPokazR3_Click(sender As Object, e As EventArgs) Handles cmdPokazR3.Click
        Dim PlikDok As String = ""
        PobierzParametry()
        If Len(txtRaport03.Text) > 0 Then
            PlikDok = Trim(txtRaport03.Text)
            If My.Computer.FileSystem.FileExists(PlikDok) Then
                Shell("cmd /C""" & PlikDok & """", AppWinStyle.Hide, True)
            Else
                MessageBox.Show("Nie znalaz³em pliku z dokumentem" & vbCrLf & PlikDok,
                    "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End If
        End If
    End Sub
    Private Sub cmdPokazR4_Click(sender As Object, e As EventArgs) Handles cmdPokazR4.Click
        Dim PlikDok As String = ""
        PobierzParametry()
        If Len(txtRaport04.Text) > 0 Then
            PlikDok = Trim(txtRaport04.Text)
            If My.Computer.FileSystem.FileExists(PlikDok) Then
                Shell("cmd /C""" & PlikDok & """", AppWinStyle.Hide, True)
            Else
                MessageBox.Show("Nie znalaz³em pliku z dokumentem" & vbCrLf & PlikDok,
                    "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End If
        End If
    End Sub
    Private Sub cmdPokazR5_Click(sender As Object, e As EventArgs) Handles cmdPokazR5.Click
        Dim PlikDok As String = ""
        PobierzParametry()
        If Len(txtRaport05.Text) > 0 Then
            PlikDok = Trim(txtRaport05.Text)
            If My.Computer.FileSystem.FileExists(PlikDok) Then
                Shell("cmd /C""" & PlikDok & """", AppWinStyle.Hide, True)
            Else
                MessageBox.Show("Nie znalaz³em pliku z dokumentem" & vbCrLf & PlikDok,
                    "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End If
        End If
    End Sub
    Private Sub cmdPokazR6_Click(sender As Object, e As EventArgs) Handles cmdPokazR6.Click
        Dim PlikDok As String = ""
        PobierzParametry()
        If Len(txtRaport06.Text) > 0 Then
            PlikDok = Trim(txtRaport06.Text)
            If My.Computer.FileSystem.FileExists(PlikDok) Then
                Shell("cmd /C""" & PlikDok & """", AppWinStyle.Hide, True)
            Else
                MessageBox.Show("Nie znalaz³em pliku z dokumentem" & vbCrLf & PlikDok,
                    "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End If
        End If
    End Sub
    Private Sub cmdPokazR7_Click(sender As Object, e As EventArgs) Handles cmdPokazR7.Click
        Dim PlikDok As String = ""
        PobierzParametry()
        If Len(txtRaport07.Text) > 0 Then
            PlikDok = Trim(txtRaport07.Text)
            If My.Computer.FileSystem.FileExists(PlikDok) Then
                Shell("cmd /C""" & PlikDok & """", AppWinStyle.Hide, True)
            Else
                MessageBox.Show("Nie znalaz³em pliku z dokumentem" & vbCrLf & PlikDok,
                    "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End If
        End If
    End Sub
    Private Sub cmdPokazR8_Click(sender As Object, e As EventArgs) Handles cmdPokazR8.Click
        Dim PlikDok As String = ""
        PobierzParametry()
        If Len(txtRaport08.Text) > 0 Then
            PlikDok = Trim(txtRaport08.Text)
            If My.Computer.FileSystem.FileExists(PlikDok) Then
                Shell("cmd /C""" & PlikDok & """", AppWinStyle.Hide, True)
            Else
                MessageBox.Show("Nie znalaz³em pliku z dokumentem" & vbCrLf & PlikDok,
                    "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End If
        End If
    End Sub
    Private Sub cmdPokazR9_Click(sender As Object, e As EventArgs) Handles cmdPokazR9.Click
        Dim PlikDok As String = ""
        PobierzParametry()
        If Len(txtRaport09.Text) > 0 Then
            PlikDok = Trim(txtRaport09.Text)
            If My.Computer.FileSystem.FileExists(PlikDok) Then
                Shell("cmd /C""" & PlikDok & """", AppWinStyle.Hide, True)
            Else
                MessageBox.Show("Nie znalaz³em pliku z dokumentem" & vbCrLf & PlikDok,
                    "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            End If
        End If
    End Sub







    '***************************************************
    '* procedury obslugi ekranu
    '***************************************************

    'Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
    '    WypelnijOpisy()
    'End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        Select Case TabControl1.SelectedIndex
            Case 0  'identyfikacja
                cmdKontaktPromotor.Visible = False
                cmdCoDalej.Visible = False
                '-----------------
                'txtTofi.Focus()
                cmdPowrot.Focus()
                '-----------------
                StartFormyOsoby = True
                dbSelectSQLOsoby = sSelectSQLOsoby & " WHERE IDENTKLIENTA='" & oIdentKli & "'"
                DataSetOsoby = New DataSet
                DataViewOsoby = New DataView

                DataSetOsoby.Locale = New System.Globalization.CultureInfo("pl-PL")
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbConnectionOsoby = New OleDb.OleDbConnection(sConnectionOsoby)
                    'dbCommandSelectOsoby = New OleDb.OleDbCommand(dbSelectSQLOsoby, dbConnectionOsoby)
                    'dbCommandDeleteOsoby = New OleDb.OleDbCommand(sDeleteSQLOsoby, dbConnectionOsoby)
                    'dbCommandUpdateOsoby = New OleDb.OleDbCommand(sUpdateSQLOsoby, dbConnectionOsoby)
                    'dbCommandInsertOsoby = New OleDb.OleDbCommand(sInsertSQLOsoby, dbConnectionOsoby)
                    'dbDataAdapterOsoby = New OleDb.OleDbDataAdapter(dbCommandSelectOsoby)

                    ''mapujemy domyslna nazwe tabeli
                    'Dim MapowanieTabeliOsoby As System.Data.Common.DataTableMapping
                    'MapowanieTabeliOsoby = dbDataAdapterOsoby.TableMappings.Add("Table", "TABELA_Osoby")

                    'DBDeleteOsoby(dbCommandDeleteOsoby, dbDataAdapterOsoby)
                    'DBUpdateOsoby(dbCommandUpdateOsoby, dbDataAdapterOsoby)
                    'DBInsertOsoby(dbCommandInsertOsoby, dbDataAdapterOsoby)

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
                    'Finally
                    '    ConnectionState = dbConnectionOsoby.State
                    'End Try

                Else
                    sqlConnectionOsoby = New SqlClient.SqlConnection(sConnectionOsoby)
                    sqlCommandSelectOsoby = New SqlClient.SqlCommand(dbSelectSQLOsoby, sqlConnectionOsoby)
                    sqlCommandDeleteOsoby = New SqlClient.SqlCommand(sDeleteSQLOsoby, sqlConnectionOsoby)
                    sqlCommandUpdateOsoby = New SqlClient.SqlCommand(sUpdateSQLOsoby, sqlConnectionOsoby)
                    sqlCommandInsertOsoby = New SqlClient.SqlCommand(sInsertSQLOsoby, sqlConnectionOsoby)
                    sqlDataAdapterOsoby = New SqlClient.SqlDataAdapter(sqlCommandSelectOsoby)

                    'mapujemy domyslna nazwe tabeli
                    Dim MapowanieTabeliOsoby As System.Data.Common.DataTableMapping
                    MapowanieTabeliOsoby = sqlDataAdapterOsoby.TableMappings.Add("Table", "TABELA_Osoby")

                    SQLDeleteOsoby(sqlCommandDeleteOsoby, sqlDataAdapterOsoby)
                    SQLUpdateOsoby(sqlCommandUpdateOsoby, sqlDataAdapterOsoby)
                    SQLInsertOsoby(sqlCommandInsertOsoby, sqlDataAdapterOsoby)

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
                    Finally
                        ConnectionState = sqlConnectionOsoby.State
                    End Try
                End If
                If ConnectionState = ConnectionState.Open Then
                    Try
                        If ParL_DataBaseType = "MS ACCESS" Then
                            'dbDataAdapterOsoby.FillSchema(DataSetOsoby, SchemaType.Mapped, "TABELA_Osoby")
                            'dbDataAdapterOsoby.Fill(DataSetOsoby)
                            'dbConnectionOsoby.Close()
                        Else
                            sqlDataAdapterOsoby.FillSchema(DataSetOsoby, SchemaType.Mapped, "TABELA_Osoby")
                            sqlDataAdapterOsoby.Fill(DataSetOsoby)
                            sqlConnectionOsoby.Close()
                        End If

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        DataSetOsoby.Tables(0).PrimaryKey = New DataColumn() {DataSetOsoby.Tables(0).Columns("IDENTKLIENTA"), DataSetOsoby.Tables(0).Columns("OSOBA")}

                        'definiuj DataView
                        DataViewOsoby = New DataView(DataSetOsoby.Tables(0))
                        DataViewOsoby.AllowDelete = False
                        DataViewOsoby.AllowNew = False
                        DataViewOsoby.AllowEdit = False

                        '-------------------------------------------------
                        'parametry DataGridView
                        '-------------------------------------------------
                        ' Initialize basic DataGridView properties.
                        dagOsoby.BackgroundColor = System.Drawing.SystemColors.Control
                        dagOsoby.GridColor = System.Drawing.SystemColors.ControlDark
                        dagOsoby.ForeColor = System.Drawing.Color.Purple
                        dagOsoby.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagOsoby.BorderStyle = BorderStyle.Fixed3D
                        'dagOsoby.Dock = DockStyle.Fill
                        dagOsoby.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                        ' Set property values appropriate for read-only display and limited interactivity. 
                        dagOsoby.AllowUserToAddRows = False
                        dagOsoby.AllowUserToDeleteRows = False
                        dagOsoby.AllowUserToOrderColumns = True
                        dagOsoby.AllowUserToResizeColumns = True
                        dagOsoby.AllowUserToResizeRows = True
                        dagOsoby.ReadOnly = True
                        dagOsoby.SelectionMode = DataGridViewSelectionMode.CellSelect
                        dagOsoby.MultiSelect = False

                        'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                        'dagOsoby.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                        dagOsoby.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                        dagOsoby.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                        dagOsoby.ColumnHeadersVisible = True
                        dagOsoby.ColumnHeadersHeight = 40
                        dagOsoby.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                        dagOsoby.RowHeadersVisible = True
                        dagOsoby.RowHeadersWidth = 50
                        dagOsoby.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                        ' Set the selection background color for all the cells.
                        dagOsoby.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                        dagOsoby.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                        dagOsoby.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagOsoby.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                        dagOsoby.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagOsoby.DefaultCellStyle.NullValue = ""
                        dagOsoby.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        dagOsoby.DefaultCellStyle.WrapMode = DataGridViewTriState.True


                        ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                        ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                        dagOsoby.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                        dagOsoby.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                        ' Set the background color for all rows and for alternating rows.  
                        ' The value for alternating rows overrides the value for all rows. 
                        dagOsoby.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagOsoby.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                        ' Set the row and column header styles.
                        dagOsoby.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagOsoby.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        dagOsoby.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagOsoby.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        '-----------------------------------
                        'wypelnienie DataGridView
                        dagOsoby.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                        'dagOsoby.DataMember = DataSetOsoby.Tables(0).TableName
                        'dagOsoby.DataSource = DataSetOsobty
                        '-----------------------------------

                        ' Add a GridColumnStyle and set the MappingName 
                        ' to the name of a DataColumn in the DataTable. 
                        ' Set the HeaderText and Width properties. 

                        dagOsoby.Columns.Clear()

                        TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(0).ColumnName, "Klient", 50, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(1).ColumnName, "Osoba", 250, DataGridViewContentAlignment.MiddleLeft, True)
                        LogColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(2).ColumnName, "V.I.P.", 40, True)
                        TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(3).ColumnName, "eMail", 250, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(4).ColumnName, "Telefon s³u¿bowy", 250, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(5).ColumnName, "Telefon komórkowy", 250, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(6).ColumnName, "Stanowisko", 250, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(7).ColumnName, "Kompetencje", 250, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagOsoby, DataSetOsoby.Tables(0).Columns(8).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

                        dagOsoby.Columns(dagOsoby.Columns.Count - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                        Me.Cursor = Cursors.WaitCursor
                        dagOsoby.DataSource = DataViewOsoby
                        Me.Cursor = Cursors.Default

                        ''ustaw sie na pierwszym zapisie
                        If DataSetOsoby.Tables(0).Rows.Count > 0 Then
                            dagOsoby.CurrentCell = dagOsoby(2, 0)
                            dagOsoby.CurrentCell.Selected = True
                        End If

                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    End Try
                    '---------------------------------
                    dagOsoby.Focus()
                    '---------------------------------
                    'dodaj StatusBar na koncu
                    stbOsoby.Panels.Clear()

                    stbOsoby.BackColor = System.Drawing.SystemColors.Control
                    stbOsoby.ForeColor = System.Drawing.Color.DarkGreen
                    stbOsoby.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                    stbOsoby.Panels.Add("stbPanelIleRec")
                    stbOsoby.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbOsoby.Panels(0).Width = 80
                    stbOsoby.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbOsoby.Panels(0).Text = "Iloœæ rec.: " & DataSetOsoby.Tables(0).Rows.Count.ToString

                    stbOsoby.Panels.Add("stbPanelWiersz")
                    stbOsoby.Panels(1).AutoSize = StatusBarPanelAutoSize.None
                    stbOsoby.Panels(1).Width = 80
                    stbOsoby.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagOsoby.CurrentCell Is Nothing Then
                        stbOsoby.Panels(1).Text = "Wi: "
                    Else
                        stbOsoby.Panels(1).Text = "Wi: " & dagOsoby.CurrentCell.RowIndex.ToString
                    End If

                    stbOsoby.Panels.Add("stbPanelKolumna")
                    stbOsoby.Panels(2).AutoSize = StatusBarPanelAutoSize.None
                    stbOsoby.Panels(2).Width = 80
                    stbOsoby.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagOsoby.CurrentCell Is Nothing Then
                        stbOsoby.Panels(2).Text = "Ko: "
                    Else
                        stbOsoby.Panels(2).Text = "Ko: " & dagOsoby.CurrentCell.ColumnIndex.ToString
                    End If

                    stbOsoby.Panels.Add("stbPanelSort")
                    stbOsoby.Panels(3).AutoSize = StatusBarPanelAutoSize.None
                    stbOsoby.Panels(3).Width = 150
                    stbOsoby.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbOsoby.Panels(3).Text = "Sort : "

                    stbOsoby.Panels.Add("stbPanelFiltr")
                    stbOsoby.Panels(4).AutoSize = StatusBarPanelAutoSize.None
                    stbOsoby.Panels(4).Width = 150
                    stbOsoby.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbOsoby.Panels(4).Text = "Filtr : "

                    stbOsoby.Panels.Add("stbPanelReszta")
                    stbOsoby.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
                    'stbOsoby.Panels(5).Width = 150
                    stbOsoby.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbOsoby.Panels(5).Text = ""

                    stbOsoby.ShowPanels = True
                    '---------------------------------
                    'dodaj StatusBar na koncu
                    stbFiltryOsoby.Panels.Clear()

                    stbFiltryOsoby.BackColor = System.Drawing.SystemColors.Control
                    stbFiltryOsoby.ForeColor = System.Drawing.Color.DarkGreen
                    stbFiltryOsoby.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                    stbFiltryOsoby.Panels.Add("stbFiltryRowHead")
                    stbFiltryOsoby.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbFiltryOsoby.Panels(0).Width = 50  'dagOsoby.TableStyles(0).RowHeaderWidth
                    stbFiltryOsoby.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbFiltryOsoby.Panels(0).Text = "Filtry"

                    stbFiltryOsoby.Panels.Add("stbFiltryReszta")
                    stbFiltryOsoby.Panels(1).AutoSize = StatusBarPanelAutoSize.Spring
                    'stbFiltryOsoby.Panels(1).Width = 150
                    stbFiltryOsoby.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbFiltryOsoby.Panels(1).Text = ""

                    stbFiltryOsoby.ShowPanels = True

                    'dodaj ctrl do pobierania szablonu
                    cmdWszystkoOsoby.Size = New Size(20, 22)
                    cmdWszystkoOsoby.Location = New System.Drawing.Point(stbFiltryOsoby.Location.X + 50 - 20,
                                             stbFiltryOsoby.Location.Y)
                    cmdWszystkoOsoby.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)

                End If
                InitOsoby() 'inicjuj zmienne

                '--------------------------------------------------------------------
                'set the header width to something apporpriate
                dagOsoby.RowHeadersWidth = 50       'use if tablestyle
                'Me.dagOsoby.RowHeaderWidth = 80 'use if no tablestyle
                '--------------------------------------------------------------------
                ''set topleft corner point so we can locate the toprow
                'If DataSetOsoby.Tables(0).Rows.Count > 0 Then
                '    StartPointInCell00 = New System.Drawing.Point((dagOsoby.GetCellDisplayRectangle(0, 0, True).X + 4),
                '                      (dagOsoby.GetCellDisplayRectangle(0, 0, True).Y + 4))
                '    ScrollXOffset = 10000 * dagOsoby.FirstDisplayedScrollingColumnIndex +
                '        dagOsoby.GetCellDisplayRectangle(dagOsoby.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                '        StartPointInCell00.X
                'Else
                '    'brak zapisow - poczatek DataGrid
                '    StartPointInCell00 = New System.Drawing.Point((dagOsoby.Left + 4), (dagOsoby.Top + 4))
                '    ScrollXOffset = 0
                'End If
                ''--------------------------------------------------
                ''inicjowanie zegara dla scrollingu poziomego
                'HorizScrollLastTime = ""
                'HorizScrollTimer.Interval = 2 * 1000
                'HorizScrollTimer.Enabled = True
                '--------------------------------------------------
                'Stworz zmienne do filtrowania
                Dim ii As Integer = 0
                For ii = 0 To DataViewOsoby.Table.Columns.Count - 1
                    'stworz zmienn¹ TXTBOX
                    Dim CtrlT As New System.Windows.Forms.TextBox
                    CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                    CtrlT.Visible = False
                    Me.Identyfikacja.Controls.Add(CtrlT)
                    AddHandler CtrlT.TextChanged, New EventHandler(AddressOf txtFiXXOsoby_TextChanged)
                    AddHandler CtrlT.GotFocus, New EventHandler(AddressOf txtFiXXOsoby_GotFocus)
                    AddHandler CtrlT.LostFocus, New EventHandler(AddressOf txtFiXXOsoby_LostFocus)
                    AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXXOsoby_KeyDown)

                    'stworz zmienn¹ BUTTON
                    Dim CtrlB As New System.Windows.Forms.Button
                    CtrlB.Name = "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                    CtrlB.Image = cmdFi.Image
                    CtrlB.ImageAlign = ContentAlignment.MiddleCenter
                    CtrlB.Visible = False
                    Me.Identyfikacja.Controls.Add(CtrlB)
                    AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXXOsoby_Click)
                    AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXXOsoby_GotFocus)
                    AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXXOsoby_LostFocus)
                Next
                Me.Refresh()
                '--------------------------------------------------
                OdswiezBazeOsoby()
                StartFormy = False
                StartFormyOsoby = False
                dagOsoby_CurrentCellChanged(sender, e)
                InicjujFiltryColumnDGV(Me.Identyfikacja, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltryOsoby, StartFormyOsoby)
                '--------------------------------------------------------------------
                cmdPowrot.Focus()










            Case 1  'opiekun
                lblIdent_Sprzedaz.Text = txtIdent.Text
                lblTofi_Sprzedaz.Text = txtTofi.Text
                lblNazwa_Sprzedaz.Text = txtNazwa1.Text
                lbladres_Sprzedaz.Text = Trim(txtKodP.Text) & " " & Trim(txtMiejscowosc.Text) & ", " & Trim(txtUlica.Text) & " " & Trim(txtNrDomu.Text) & " " & Trim(txtIDDomu.Text)
                lblosoba_Sprzedaz.Text = OsobaKontaktowa(txtIdent.Text)
                lblStanowisko_Sprzedaz.Text = oStanowiskoOso
                lblTelefon_Sprzedaz.Text = oTelefonOso
                lblTelefonKom_Sprzedaz.Text = oTelefonKomOso

                cmdKontaktPromotor.Visible = True
                cmdCoDalej.Visible = True
                '-----------------
                StartFormyOpiekun = True
                dbSelectSQLKontakty = sSelectSQLKontakty &
                                        " WHERE (IDENTKLIENTA='" & oIdentKli & "') and (IDENTKLIENTA<>'') ORDER BY DATAKONTAKTU DESC"
                DataSetKontakty = New DataSet
                DataViewKontakty = New DataView

                DataSetKontakty.Locale = New System.Globalization.CultureInfo("pl-PL")
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbConnectionKontakty = New OleDb.OleDbConnection(sConnectionKontakty)
                    'dbCommandSelectKontakty = New OleDb.OleDbCommand(dbSelectSQLKontakty, dbConnectionKontakty)
                    'dbCommandDeleteKontakty = New OleDb.OleDbCommand(sDeleteSQLKontakty, dbConnectionKontakty)
                    'dbCommandUpdateKontakty = New OleDb.OleDbCommand(sUpdateSQLKontakty, dbConnectionKontakty)
                    'dbCommandInsertKontakty = New OleDb.OleDbCommand(sInsertSQLKontakty, dbConnectionKontakty)
                    'dbDataAdapterKontakty = New OleDb.OleDbDataAdapter(dbCommandSelectKontakty)

                    ''mapujemy domyslna nazwe tabeli
                    'Dim MapowanieTabeliKontakty As System.Data.Common.DataTableMapping
                    'MapowanieTabeliKontakty = dbDataAdapterKontakty.TableMappings.Add("Table", "TABELA_KontaktyHandlowe")

                    'DBDeleteKontaktyHandlowe(dbCommandDeleteKontakty, dbDataAdapterKontakty)
                    'DBUpdateKontaktyHandlowe(dbCommandUpdateKontakty, dbDataAdapterKontakty)
                    'DBInsertKontaktyHandlowe(dbCommandInsertKontakty, dbDataAdapterKontakty)


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
                    'Finally
                    '    ConnectionState = dbConnectionKontakty.State
                    'End Try

                Else
                    sqlConnectionKontakty = New SqlClient.SqlConnection(sConnectionKontakty)
                    sqlCommandSelectKontakty = New SqlClient.SqlCommand(dbSelectSQLKontakty, sqlConnectionKontakty)
                    sqlCommandDeleteKontakty = New SqlClient.SqlCommand(sDeleteSQLKontakty, sqlConnectionKontakty)
                    sqlCommandUpdateKontakty = New SqlClient.SqlCommand(sUpdateSQLKontakty, sqlConnectionKontakty)
                    sqlCommandInsertKontakty = New SqlClient.SqlCommand(sInsertSQLKontakty, sqlConnectionKontakty)
                    sqlDataAdapterKontakty = New SqlClient.SqlDataAdapter(sqlCommandSelectKontakty)

                    'mapujemy domyslna nazwe tabeli
                    Dim MapowanieTabeliKontakty As System.Data.Common.DataTableMapping
                    MapowanieTabeliKontakty = sqlDataAdapterKontakty.TableMappings.Add("Table", "TABELA_KontaktyHandlowe")

                    SQLDeleteKontaktyHandlowe(sqlCommandDeleteKontakty, sqlDataAdapterKontakty)
                    SQLUpdateKontaktyHandlowe(sqlCommandUpdateKontakty, sqlDataAdapterKontakty)
                    SQLInsertKontaktyHandlowe(sqlCommandInsertKontakty, sqlDataAdapterKontakty)

                    ' Add the RowUpdated event handler.
                    AddHandler sqlDataAdapterKontakty.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

                    '------------------------------------------
                    'wypelnij DATASET
                    Try
                        sqlConnectionKontakty.Open()
                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    Finally
                        ConnectionState = sqlConnectionKontakty.State
                    End Try
                End If
                If ConnectionState = ConnectionState.Open Then
                    Try
                        If ParL_DataBaseType = "MS ACCESS" Then
                            'dbDataAdapterKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_KontaktyHandlowe")
                            'dbDataAdapterKontakty.Fill(DataSetKontakty)
                            'dbConnectionKontakty.Close()
                        Else
                            sqlDataAdapterKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_KontaktyHandlowe")
                            sqlDataAdapterKontakty.Fill(DataSetKontakty)
                            sqlConnectionKontakty.Close()
                        End If

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        DataSetKontakty.Tables(0).PrimaryKey = New DataColumn() {DataSetKontakty.Tables(0).Columns("UNIQID")}

                        'definiuj DataView
                        DataViewKontakty = New DataView(DataSetKontakty.Tables(0))
                        DataViewKontakty.AllowDelete = False
                        DataViewKontakty.AllowNew = False
                        DataViewKontakty.AllowEdit = False

                        '-------------------------------------------------
                        'parametry DataGridView
                        '-------------------------------------------------
                        ' Initialize basic DataGridView properties.
                        dagKontakty.BackgroundColor = System.Drawing.SystemColors.Control
                        dagKontakty.GridColor = System.Drawing.SystemColors.ControlDark
                        dagKontakty.ForeColor = System.Drawing.Color.Purple
                        dagKontakty.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagKontakty.BorderStyle = BorderStyle.Fixed3D
                        'dagKontakty.Dock = DockStyle.Fill
                        dagKontakty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                        ' Set property values appropriate for read-only display and limited interactivity. 
                        dagKontakty.AllowUserToAddRows = False
                        dagKontakty.AllowUserToDeleteRows = False
                        dagKontakty.AllowUserToOrderColumns = True
                        dagKontakty.AllowUserToResizeColumns = True
                        dagKontakty.AllowUserToResizeRows = True
                        dagKontakty.ReadOnly = True
                        dagKontakty.SelectionMode = DataGridViewSelectionMode.CellSelect
                        dagKontakty.MultiSelect = False

                        'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                        'dagKontakty.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                        dagKontakty.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                        dagKontakty.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                        dagKontakty.ColumnHeadersVisible = True
                        dagKontakty.ColumnHeadersHeight = 40
                        dagKontakty.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                        dagKontakty.RowHeadersVisible = False
                        dagKontakty.RowHeadersWidth = 50
                        dagKontakty.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                        ' Set the selection background color for all the cells.
                        dagKontakty.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                        dagKontakty.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                        dagKontakty.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagKontakty.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                        dagKontakty.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagKontakty.DefaultCellStyle.NullValue = ""
                        dagKontakty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        dagKontakty.DefaultCellStyle.WrapMode = DataGridViewTriState.True


                        ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                        ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                        dagKontakty.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                        dagKontakty.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                        ' Set the background color for all rows and for alternating rows.  
                        ' The value for alternating rows overrides the value for all rows. 
                        dagKontakty.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagKontakty.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                        ' Set the row and column header styles.
                        dagKontakty.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagKontakty.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        dagKontakty.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagKontakty.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        '-----------------------------------
                        'wypelnienie DataGridView
                        dagKontakty.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                        'dagKontakty.DataMember = DataSetKontakty.Tables(0).TableName
                        'dagKontakty.DataSource = DataSetKontakty
                        '-----------------------------------

                        ' Add a GridColumnStyle and set the MappingName 
                        ' to the name of a DataColumn in the DataTable. 
                        ' Set the HeaderText and Width properties. 

                        dagKontakty.Columns.Clear()

                        TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                        TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(1).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                        TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(2).ColumnName, "Data", 70, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(3).ColumnName, "Pracownik", 80, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(4).ColumnName, "Rodzaj kontaktu", 100, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(5).ColumnName, "Opis", 750, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagKontakty, DataSetKontakty.Tables(0).Columns(6).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

                        'dagKontakty.Columns(dagKontakty.Columns.Count - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                        Me.Cursor = Cursors.WaitCursor
                        dagKontakty.DataSource = DataViewKontakty
                        Me.Cursor = Cursors.Default

                        ''ustaw sie na pierwszym zapisie
                        If DataSetKontakty.Tables(0).Rows.Count > 0 Then
                            dagKontakty.CurrentCell = dagKontakty(2, 0)
                            dagKontakty.CurrentCell.Selected = True
                        End If


                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    End Try
                    '---------------------------------
                    dagKontakty.Focus()
                End If
                InitKontakty() 'inicjuj zmienne
                '--------------------------------------------------------------------






                '-----------------
                dbSelectSQLCoDalej = sSelectSQLCoDalej &
                                        " WHERE (IDENTKLIENTA='" & oIdentKli & "') and (IDENTKLIENTA<>'') ORDER BY IDENT DESC"
                DataSetCoDalej = New DataSet
                DataViewCoDalej = New DataView

                DataSetCoDalej.Locale = New System.Globalization.CultureInfo("pl-PL")
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbConnectionCoDalej = New OleDb.OleDbConnection(sConnectionCoDalej)
                    'dbCommandSelectCoDalej = New OleDb.OleDbCommand(dbSelectSQLCoDalej, dbConnectionCoDalej)
                    'dbCommandDeleteCoDalej = New OleDb.OleDbCommand(sDeleteSQLCoDalej, dbConnectionCoDalej)
                    'dbCommandUpdateCoDalej = New OleDb.OleDbCommand(sUpdateSQLCoDalej, dbConnectionCoDalej)
                    'dbCommandInsertCoDalej = New OleDb.OleDbCommand(sInsertSQLCoDalej, dbConnectionCoDalej)
                    'dbDataAdapterCoDalej = New OleDb.OleDbDataAdapter(dbCommandSelectCoDalej)

                    ''mapujemy domyslna nazwe tabeli
                    'Dim MapowanieTabeliCoDalej As System.Data.Common.DataTableMapping
                    'MapowanieTabeliCoDalej = dbDataAdapterCoDalej.TableMappings.Add("Table", "TABELA_CoDalej")

                    'DBDeleteCoDalej(dbCommandDeleteCoDalej, dbDataAdapterCoDalej)
                    'DBUpdateCoDalej(dbCommandUpdateCoDalej, dbDataAdapterCoDalej)
                    'DBInsertCoDalej(dbCommandInsertCoDalej, dbDataAdapterCoDalej)

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
                    'Finally
                    '    ConnectionState = dbConnectionCoDalej.State
                    'End Try

                Else
                    sqlConnectionCoDalej = New SqlClient.SqlConnection(sConnectionCoDalej)
                    sqlCommandSelectCoDalej = New SqlClient.SqlCommand(dbSelectSQLCoDalej, sqlConnectionCoDalej)
                    sqlCommandDeleteCoDalej = New SqlClient.SqlCommand(sDeleteSQLCoDalej, sqlConnectionCoDalej)
                    sqlCommandUpdateCoDalej = New SqlClient.SqlCommand(sUpdateSQLCoDalej, sqlConnectionCoDalej)
                    sqlCommandInsertCoDalej = New SqlClient.SqlCommand(sInsertSQLCoDalej, sqlConnectionCoDalej)
                    sqlDataAdapterCoDalej = New SqlClient.SqlDataAdapter(sqlCommandSelectCoDalej)

                    'mapujemy domyslna nazwe tabeli
                    Dim MapowanieTabeliCoDalej As System.Data.Common.DataTableMapping
                    MapowanieTabeliCoDalej = sqlDataAdapterCoDalej.TableMappings.Add("Table", "TABELA_CoDalej")

                    SQLDeleteCoDalej(sqlCommandDeleteCoDalej, sqlDataAdapterCoDalej)
                    SQLUpdateCoDalej(sqlCommandUpdateCoDalej, sqlDataAdapterCoDalej)
                    SQLInsertCoDalej(sqlCommandInsertCoDalej, sqlDataAdapterCoDalej)

                    ' Add the RowUpdated event handler.
                    AddHandler sqlDataAdapterCoDalej.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)
                    '------------------------------------------
                    'wypelnij DATASET
                    Try
                        sqlConnectionCoDalej.Open()
                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    Finally
                        ConnectionState = sqlConnectionCoDalej.State
                    End Try
                End If
                If ConnectionState = ConnectionState.Open Then
                    Try
                        If ParL_DataBaseType = "MS ACCESS" Then
                            'dbDataAdapterCoDalej.FillSchema(DataSetCoDalej, SchemaType.Mapped, "TABELA_CoDalej")
                            'dbDataAdapterCoDalej.Fill(DataSetCoDalej)
                            'dbConnectionCoDalej.Close()
                        Else
                            sqlDataAdapterCoDalej.FillSchema(DataSetCoDalej, SchemaType.Mapped, "TABELA_CoDalej")
                            sqlDataAdapterCoDalej.Fill(DataSetCoDalej)
                            sqlConnectionCoDalej.Close()
                        End If

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        DataSetCoDalej.Tables(0).PrimaryKey = New DataColumn() {DataSetCoDalej.Tables(0).Columns("UNIQID"),
                                                                         DataSetCoDalej.Tables(0).Columns("IDENTKLIENTA")}

                        'definiuj DataView
                        DataViewCoDalej = New DataView(DataSetCoDalej.Tables(0))
                        DataViewCoDalej.AllowDelete = False
                        DataViewCoDalej.AllowNew = False
                        DataViewCoDalej.AllowEdit = False

                        '-------------------------------------------------
                        'parametry DataGridView
                        '-------------------------------------------------
                        ' Initialize basic DataGridView properties.
                        dagCoDalej.BackgroundColor = System.Drawing.SystemColors.Control
                        dagCoDalej.GridColor = System.Drawing.SystemColors.ControlDark
                        dagCoDalej.ForeColor = System.Drawing.Color.Purple
                        dagCoDalej.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagCoDalej.BorderStyle = BorderStyle.Fixed3D
                        'dagCoDalej.Dock = DockStyle.Fill
                        dagCoDalej.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                        ' Set property values appropriate for read-only display and limited interactivity. 
                        dagCoDalej.AllowUserToAddRows = False
                        dagCoDalej.AllowUserToDeleteRows = False
                        dagCoDalej.AllowUserToOrderColumns = True
                        dagCoDalej.AllowUserToResizeColumns = True
                        dagCoDalej.AllowUserToResizeRows = True
                        dagCoDalej.ReadOnly = True
                        dagCoDalej.SelectionMode = DataGridViewSelectionMode.CellSelect
                        dagCoDalej.MultiSelect = False

                        'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                        'dagCoDalej.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                        dagCoDalej.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                        dagCoDalej.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                        dagCoDalej.ColumnHeadersVisible = True
                        dagCoDalej.ColumnHeadersHeight = 22
                        dagCoDalej.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                        dagCoDalej.RowHeadersVisible = False
                        dagCoDalej.RowHeadersWidth = 50
                        dagCoDalej.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                        ' Set the selection background color for all the cells.
                        dagCoDalej.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                        dagCoDalej.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                        dagCoDalej.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagCoDalej.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                        dagCoDalej.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagCoDalej.DefaultCellStyle.NullValue = ""
                        dagCoDalej.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        dagCoDalej.DefaultCellStyle.WrapMode = DataGridViewTriState.True


                        ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                        ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                        dagCoDalej.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                        dagCoDalej.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                        ' Set the background color for all rows and for alternating rows.  
                        ' The value for alternating rows overrides the value for all rows. 
                        dagCoDalej.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagCoDalej.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                        ' Set the row and column header styles.
                        dagCoDalej.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagCoDalej.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        dagCoDalej.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagCoDalej.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        '-----------------------------------
                        'wypelnienie DataGridView
                        dagCoDalej.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                        'dagCoDalej.DataMember = DataSetCoDalej.Tables(0).TableName
                        'dagCoDalej.DataSource = DataSetCoDalej
                        '-----------------------------------

                        ' Add a GridColumnStyle and set the MappingName 
                        ' to the name of a DataColumn in the DataTable. 
                        ' Set the HeaderText and Width properties. 

                        dagCoDalej.Columns.Clear()

                        TxtColumnView(dagCoDalej, DataSetCoDalej.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                        TxtColumnView(dagCoDalej, DataSetCoDalej.Tables(0).Columns(1).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                        TxtColumnView(dagCoDalej, DataSetCoDalej.Tables(0).Columns(2).ColumnName, "Identyfik.", 100, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagCoDalej, DataSetCoDalej.Tables(0).Columns(3).ColumnName, "Opis", 900, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagCoDalej, DataSetCoDalej.Tables(0).Columns(4).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

                        'dagCoDalej.Columns(dagCoDalej.Columns.Count - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                        Me.Cursor = Cursors.WaitCursor
                        dagCoDalej.DataSource = DataViewCoDalej
                        Me.Cursor = Cursors.Default

                        ''ustaw sie na pierwszym zapisie
                        If DataSetCoDalej.Tables(0).Rows.Count > 0 Then
                            dagCoDalej.CurrentCell = dagCoDalej(2, 0)
                            dagCoDalej.CurrentCell.Selected = True
                        End If


                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    End Try
                    '---------------------------------
                    dagCoDalej.Focus()
                End If
                InitCoDalej() 'inicjuj zmienne
                StartFormyOpiekun = False
                '--------------------------------------------------------------------
                'txtPKontakt.Focus()
                cmdPowrot.Focus()




            Case 2  'charakterystyka
                lblIdent_Charakterystyka.Text = txtIdent.Text
                lblTofi_Charakterystyka.Text = txtTofi.Text
                lblNazwaHandlowa_Charakterystyka.Text = txtNazwa1.Text
                lblNazwa1_Charakterystyka.Text = Trim(txtKodP.Text) & " " & Trim(txtMiejscowosc.Text) & ", " & Trim(txtUlica.Text) & " " & Trim(txtNrDomu.Text) & " " & Trim(txtIDDomu.Text)
                lblNazwa2_Charakterystyka.Text = OsobaKontaktowa(txtIdent.Text)
                lblStanowisko_Charakterystyka.Text = oStanowiskoOso
                lblTelefon_Charakterystyka.Text = oTelefonOso
                lblTelefonKom_Charakterystyka.Text = oTelefonKomOso

                cmdKontaktPromotor.Visible = False
                cmdCoDalej.Visible = False
                '-----------------
                'txtSkupOpiekun.Focus()
                cmdPowrot.Focus()




            Case 3  'RZU
                lblIdent_UDod.Text = txtIdent.Text
                lblTofi_UDod.Text = txtTofi.Text
                lblNazwa1_UDod.Text = txtNazwa1.Text
                lblNazwa2_UDod.Text = Trim(txtKodP.Text) & " " & Trim(txtMiejscowosc.Text) & ", " & Trim(txtUlica.Text) & " " & Trim(txtNrDomu.Text) & " " & Trim(txtIDDomu.Text)
                lblNazwaHandlowa_UDod.Text = OsobaKontaktowa(txtIdent.Text)
                lblStanowisko_UDod.Text = oStanowiskoOso
                lblTelefon_UDod.Text = oTelefonOso
                lblTelefonKom_UDod.Text = oTelefonKomOso

                cmdKontaktPromotor.Visible = False
                cmdCoDalej.Visible = False
                '-----------------
                StartFormyRZUKontakty = True
                dbSelectSQLRZUKontakty = sSelectSQLKontakty &
                                        " WHERE (IDENTKLIENTA='" & oIdentKli & "') and (IDENTKLIENTA<>'') ORDER BY DATAKONTAKTU DESC"
                DataSetRZUKontakty = New DataSet
                DataViewRZUKontakty = New DataView

                DataSetRZUKontakty.Locale = New System.Globalization.CultureInfo("pl-PL")
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbConnectionRZUKontakty = New OleDb.OleDbConnection(sConnectionKontakty)
                    'dbCommandSelectRZUKontakty = New OleDb.OleDbCommand(dbSelectSQLRZUKontakty, dbConnectionRZUKontakty)
                    'dbCommandDeleteRZUKontakty = New OleDb.OleDbCommand(sDeleteSQLKontakty, dbConnectionRZUKontakty)
                    'dbCommandUpdateRZUKontakty = New OleDb.OleDbCommand(sUpdateSQLKontakty, dbConnectionRZUKontakty)
                    'dbCommandInsertRZUKontakty = New OleDb.OleDbCommand(sInsertSQLKontakty, dbConnectionRZUKontakty)
                    'dbDataAdapterRZUKontakty = New OleDb.OleDbDataAdapter(dbCommandSelectRZUKontakty)

                    ''mapujemy domyslna nazwe tabeli
                    'Dim MapowanieTabeliRZUKontakty As System.Data.Common.DataTableMapping
                    'MapowanieTabeliRZUKontakty = dbDataAdapterRZUKontakty.TableMappings.Add("Table", "TABELA_RZUKontaktyHandlowe")

                    'DBDeleteKontaktyHandlowe(dbCommandDeleteRZUKontakty, dbDataAdapterRZUKontakty)
                    'DBUpdateKontaktyHandlowe(dbCommandUpdateRZUKontakty, dbDataAdapterRZUKontakty)
                    'DBInsertKontaktyHandlowe(dbCommandInsertRZUKontakty, dbDataAdapterRZUKontakty)


                    '' Add the RowUpdated event handler.
                    'AddHandler dbDataAdapterRZUKontakty.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

                    ''------------------------------------------
                    ''wypelnij DATASET
                    'Try
                    '    dbConnectionRZUKontakty.Open()
                    'Catch Ex As System.Exception
                    '    ViewInLog(Ex, Me.Name(), Nothing)
                    '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    '    MessageBoxIcon.Information, _
                    '    '    MessageBoxDefaultButton.Button1)
                    'Finally
                    '    ConnectionState = dbConnectionRZUKontakty.State
                    'End Try

                Else
                    sqlConnectionRZUKontakty = New SqlClient.SqlConnection(sConnectionKontakty)
                    sqlCommandSelectRZUKontakty = New SqlClient.SqlCommand(dbSelectSQLRZUKontakty, sqlConnectionRZUKontakty)
                    sqlCommandDeleteRZUKontakty = New SqlClient.SqlCommand(sDeleteSQLKontakty, sqlConnectionRZUKontakty)
                    sqlCommandUpdateRZUKontakty = New SqlClient.SqlCommand(sUpdateSQLKontakty, sqlConnectionRZUKontakty)
                    sqlCommandInsertRZUKontakty = New SqlClient.SqlCommand(sInsertSQLKontakty, sqlConnectionRZUKontakty)
                    sqlDataAdapterRZUKontakty = New SqlClient.SqlDataAdapter(sqlCommandSelectRZUKontakty)

                    'mapujemy domyslna nazwe tabeli
                    Dim MapowanieTabeliRZUKontakty As System.Data.Common.DataTableMapping
                    MapowanieTabeliRZUKontakty = sqlDataAdapterRZUKontakty.TableMappings.Add("Table", "TABELA_RZUKontaktyHandlowe")

                    SQLDeleteKontaktyHandlowe(sqlCommandDeleteRZUKontakty, sqlDataAdapterRZUKontakty)
                    SQLUpdateKontaktyHandlowe(sqlCommandUpdateRZUKontakty, sqlDataAdapterRZUKontakty)
                    SQLInsertKontaktyHandlowe(sqlCommandInsertRZUKontakty, sqlDataAdapterRZUKontakty)

                    ' Add the RowUpdated event handler.
                    AddHandler sqlDataAdapterRZUKontakty.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

                    '------------------------------------------
                    'wypelnij DATASET
                    Try
                        sqlConnectionRZUKontakty.Open()
                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    Finally
                        ConnectionState = sqlConnectionRZUKontakty.State
                    End Try
                End If
                If ConnectionState = ConnectionState.Open Then
                    Try
                        If ParL_DataBaseType = "MS ACCESS" Then
                            'dbDataAdapterRZUKontakty.FillSchema(DataSetRZUKontakty, SchemaType.Mapped, "TABELA_RZUKontaktyHandlowe")
                            'dbDataAdapterRZUKontakty.Fill(DataSetRZUKontakty)
                            'dbConnectionRZUKontakty.Close()
                        Else
                            sqlDataAdapterRZUKontakty.FillSchema(DataSetRZUKontakty, SchemaType.Mapped, "TABELA_RZUKontaktyHandlowe")
                            sqlDataAdapterRZUKontakty.Fill(DataSetRZUKontakty)
                            sqlConnectionRZUKontakty.Close()
                        End If

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        DataSetRZUKontakty.Tables(0).PrimaryKey = New DataColumn() {DataSetRZUKontakty.Tables(0).Columns("UNIQID")}

                        'definiuj DataView
                        DataViewRZUKontakty = New DataView(DataSetRZUKontakty.Tables(0))
                        DataViewRZUKontakty.AllowDelete = False
                        DataViewRZUKontakty.AllowNew = False
                        DataViewRZUKontakty.AllowEdit = False

                        '-------------------------------------------------
                        'parametry DataGridView
                        '-------------------------------------------------
                        ' Initialize basic DataGridView properties.
                        dagRZUKontakty.BackgroundColor = System.Drawing.SystemColors.Control
                        dagRZUKontakty.GridColor = System.Drawing.SystemColors.ControlDark
                        dagRZUKontakty.ForeColor = System.Drawing.Color.Purple
                        dagRZUKontakty.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagRZUKontakty.BorderStyle = BorderStyle.Fixed3D
                        'dagRZUKontakty.Dock = DockStyle.Fill
                        dagRZUKontakty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                        ' Set property values appropriate for read-only display and limited interactivity. 
                        dagRZUKontakty.AllowUserToAddRows = False
                        dagRZUKontakty.AllowUserToDeleteRows = False
                        dagRZUKontakty.AllowUserToOrderColumns = True
                        dagRZUKontakty.AllowUserToResizeColumns = True
                        dagRZUKontakty.AllowUserToResizeRows = True
                        dagRZUKontakty.ReadOnly = True
                        dagRZUKontakty.SelectionMode = DataGridViewSelectionMode.CellSelect
                        dagRZUKontakty.MultiSelect = False

                        'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                        'dagRZUKontakty.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                        dagRZUKontakty.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                        dagRZUKontakty.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                        dagRZUKontakty.ColumnHeadersVisible = True
                        dagRZUKontakty.ColumnHeadersHeight = 40
                        dagRZUKontakty.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                        dagRZUKontakty.RowHeadersVisible = True
                        dagRZUKontakty.RowHeadersWidth = 50
                        dagRZUKontakty.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                        ' Set the selection background color for all the cells.
                        dagRZUKontakty.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                        dagRZUKontakty.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                        dagRZUKontakty.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagRZUKontakty.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                        dagRZUKontakty.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagRZUKontakty.DefaultCellStyle.NullValue = ""
                        dagRZUKontakty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        dagRZUKontakty.DefaultCellStyle.WrapMode = DataGridViewTriState.True


                        ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                        ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                        dagRZUKontakty.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                        dagRZUKontakty.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                        ' Set the background color for all rows and for alternating rows.  
                        ' The value for alternating rows overrides the value for all rows. 
                        dagRZUKontakty.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagRZUKontakty.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                        ' Set the row and column header styles.
                        dagRZUKontakty.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagRZUKontakty.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        dagRZUKontakty.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagRZUKontakty.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        '-----------------------------------
                        'wypelnienie DataGridView
                        dagRZUKontakty.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                        'dagRZUKontakty.DataMember = DataSetKontakty.Tables(0).TableName
                        'dagRZUKontakty.DataSource = DataSetKontakty
                        '-----------------------------------

                        ' Add a GridColumnStyle and set the MappingName 
                        ' to the name of a DataColumn in the DataTable. 
                        ' Set the HeaderText and Width properties. 

                        dagRZUKontakty.Columns.Clear()

                        TxtColumnView(dagRZUKontakty, DataSetKontakty.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                        TxtColumnView(dagRZUKontakty, DataSetKontakty.Tables(0).Columns(1).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                        TxtColumnView(dagRZUKontakty, DataSetKontakty.Tables(0).Columns(2).ColumnName, "Data", 80, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagRZUKontakty, DataSetKontakty.Tables(0).Columns(3).ColumnName, "Pracownik", 80, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagRZUKontakty, DataSetKontakty.Tables(0).Columns(4).ColumnName, "Rodzaj kontaktu", 100, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagRZUKontakty, DataSetKontakty.Tables(0).Columns(5).ColumnName, "Opis", 750, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagRZUKontakty, DataSetKontakty.Tables(0).Columns(6).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, False)

                        'dagRZUKontakty.Columns(dagRZUKontakty.Columns.Count - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                        Me.Cursor = Cursors.WaitCursor
                        dagRZUKontakty.DataSource = DataViewRZUKontakty
                        Me.Cursor = Cursors.Default

                        ''ustaw sie na pierwszym zapisie
                        If DataSetKontakty.Tables(0).Rows.Count > 0 Then
                            dagRZUKontakty.CurrentCell = dagRZUKontakty(2, 0)
                            dagRZUKontakty.CurrentCell.Selected = True
                        End If


                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    End Try
                    '---------------------------------
                    'dodaj StatusBar na koncu
                    stbRZUKontakty.Panels.Clear()

                    stbRZUKontakty.BackColor = System.Drawing.SystemColors.Control
                    stbRZUKontakty.ForeColor = System.Drawing.Color.DarkGreen
                    stbRZUKontakty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                    stbRZUKontakty.Panels.Add("stbPanelIleRec")
                    stbRZUKontakty.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbRZUKontakty.Panels(0).Width = 80
                    stbRZUKontakty.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbRZUKontakty.Panels(0).Text = "Iloœæ rec.: " & DataSetRZUKontakty.Tables(0).Rows.Count.ToString

                    stbRZUKontakty.Panels.Add("stbPanelWiersz")
                    stbRZUKontakty.Panels(1).AutoSize = StatusBarPanelAutoSize.None
                    stbRZUKontakty.Panels(1).Width = 80
                    stbRZUKontakty.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagRZUKontakty.CurrentCell Is Nothing Then
                        stbRZUKontakty.Panels(1).Text = "Wi: "
                    Else
                        stbRZUKontakty.Panels(1).Text = "Wi: " & dagRZUKontakty.CurrentCell.RowIndex.ToString
                    End If

                    stbRZUKontakty.Panels.Add("stbPanelKolumna")
                    stbRZUKontakty.Panels(2).AutoSize = StatusBarPanelAutoSize.None
                    stbRZUKontakty.Panels(2).Width = 80
                    stbRZUKontakty.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagRZUKontakty.CurrentCell Is Nothing Then
                        stbRZUKontakty.Panels(2).Text = "Ko: "
                    Else
                        stbRZUKontakty.Panels(2).Text = "Ko: " & dagRZUKontakty.CurrentCell.ColumnIndex.ToString
                    End If

                    stbRZUKontakty.Panels.Add("stbPanelSort")
                    stbRZUKontakty.Panels(3).AutoSize = StatusBarPanelAutoSize.None
                    stbRZUKontakty.Panels(3).Width = 150
                    stbRZUKontakty.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbRZUKontakty.Panels(3).Text = "Sort : "

                    stbRZUKontakty.Panels.Add("stbPanelFiltr")
                    stbRZUKontakty.Panels(4).AutoSize = StatusBarPanelAutoSize.None
                    stbRZUKontakty.Panels(4).Width = 150
                    stbRZUKontakty.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbRZUKontakty.Panels(4).Text = "Filtr : "

                    stbRZUKontakty.Panels.Add("stbPanelReszta")
                    stbRZUKontakty.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
                    'stbRZUKontakty.Panels(5).Width = 150
                    stbRZUKontakty.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbRZUKontakty.Panels(5).Text = ""

                    stbRZUKontakty.ShowPanels = True
                    '---------------------------------
                    'dodaj StatusBar na koncu
                    stbFiltryRZUKontakty.Panels.Clear()

                    stbFiltryRZUKontakty.BackColor = System.Drawing.SystemColors.Control
                    stbFiltryRZUKontakty.ForeColor = System.Drawing.Color.DarkGreen
                    stbFiltryRZUKontakty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                    stbFiltryRZUKontakty.Panels.Add("stbFiltryRowHead")
                    stbFiltryRZUKontakty.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbFiltryRZUKontakty.Panels(0).Width = 50  'dagRZUKontakty.TableStyles(0).RowHeaderWidth
                    stbFiltryRZUKontakty.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbFiltryRZUKontakty.Panels(0).Text = "Filtry"

                    stbFiltryRZUKontakty.Panels.Add("stbFiltryReszta")
                    stbFiltryRZUKontakty.Panels(1).AutoSize = StatusBarPanelAutoSize.Spring
                    'stbFiltryRZUKontakty.Panels(1).Width = 150
                    stbFiltryRZUKontakty.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbFiltryRZUKontakty.Panels(1).Text = ""

                    stbFiltryRZUKontakty.ShowPanels = True

                    'dodaj ctrl do pobierania szablonu
                    cmdWszystkoRZUKontakty.Size = New Size(20, 22)
                    cmdWszystkoRZUKontakty.Location = New System.Drawing.Point(stbFiltryRZUKontakty.Location.X + 50 - 20,
                                             stbFiltryRZUKontakty.Location.Y)
                    cmdWszystkoRZUKontakty.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)

                End If
                InitRZUKontakty() 'inicjuj zmienne

                '--------------------------------------------------------------------
                'set the header width to something apporpriate
                dagRZUKontakty.RowHeadersWidth = 50       'use if tablestyle
                'Me.dagRZUKontakty.RowHeaderWidth = 80 'use if no tablestyle
                '--------------------------------------------------------------------
                ''set topleft corner point so we can locate the toprow
                'If DataSetRZUKontakty.Tables(0).Rows.Count > 0 Then
                '    StartPointInCell00 = New System.Drawing.Point((dagRZUKontakty.GetCellDisplayRectangle(0, 0, True).X + 4),
                '                      (dagRZUKontakty.GetCellDisplayRectangle(0, 0, True).Y + 4))
                '    ScrollXOffset = 10000 * dagRZUKontakty.FirstDisplayedScrollingColumnIndex +
                '        dagRZUKontakty.GetCellDisplayRectangle(dagRZUKontakty.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                '        StartPointInCell00.X
                'Else
                '    'brak zapisow - poczatek DataGrid
                '    StartPointInCell00 = New System.Drawing.Point((dagRZUKontakty.Left + 4), (dagRZUKontakty.Top + 4))
                '    ScrollXOffset = 0
                'End If
                ''--------------------------------------------------
                ''inicjowanie zegara dla scrollingu poziomego
                'HorizScrollLastTime = ""
                'HorizScrollTimer.Interval = 2 * 1000
                'HorizScrollTimer.Enabled = True
                '--------------------------------------------------
                'Stworz zmienne do filtrowania
                Dim ii As Integer = 0
                For ii = 0 To DataViewRZUKontakty.Table.Columns.Count - 1
                    'stworz zmienn¹ TXTBOX
                    Dim CtrlT As New System.Windows.Forms.TextBox
                    CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                    CtrlT.Visible = False
                    Me.UslugiDod.Controls.Add(CtrlT)
                    AddHandler CtrlT.TextChanged, New EventHandler(AddressOf txtFiXXRZUKontakty_TextChanged)
                    AddHandler CtrlT.GotFocus, New EventHandler(AddressOf txtFiXXRZUKontakty_GotFocus)
                    AddHandler CtrlT.LostFocus, New EventHandler(AddressOf txtFiXXRZUKontakty_LostFocus)
                    AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXXRZUKontakty_KeyDown)

                    'stworz zmienn¹ BUTTON
                    Dim CtrlB As New System.Windows.Forms.Button
                    CtrlB.Name = "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                    CtrlB.Image = cmdFi.Image
                    CtrlB.ImageAlign = ContentAlignment.MiddleCenter
                    CtrlB.Visible = False
                    Me.UslugiDod.Controls.Add(CtrlB)
                    AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXXRZUKontakty_Click)
                    AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXXRZUKontakty_GotFocus)
                    AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXXRZUKontakty_LostFocus)
                Next
                Me.Refresh()
                '--------------------------------------------------
                StartFormy = False
                StartFormyRZUKontakty = False
                dagRZUKontakty_CurrentCellChanged(sender, e)
                InicjujFiltryColumnDGV(Me.UslugiDod, dagRZUKontakty, DataViewRZUKontakty.Table.Columns.Count, stbFiltryRZUKontakty, StartFormyRZUKontakty)
                '--------------------------------------------------------------------
                cmdPowrot.Focus()








            Case 4  'analizy zakupu
                lblIdent_Analizy.Text = txtIdent.Text
                lblTOFI_Analizy.Text = txtTofi.Text
                lblNazwaHandlowa_Analizy.Text = txtNazwa1.Text
                lblNazwa1_Analizy.Text = Trim(txtKodP.Text) & " " & Trim(txtMiejscowosc.Text) & ", " & Trim(txtUlica.Text) & " " & Trim(txtNrDomu.Text) & " " & Trim(txtIDDomu.Text)
                lblNazwa2_Analizy.Text = OsobaKontaktowa(txtIdent.Text)
                lblStanowisko_Analizy.Text = oStanowiskoOso
                lblTelefon_Analizy.Text = oTelefonOso
                lblTelefonKom_Analizy.Text = oTelefonKomOso

                cmdKontaktPromotor.Visible = False
                cmdCoDalej.Visible = False
                '---------------------------------------------
                'Analizy zakupu
                StartFormyAnalizyZakupu = True
                dbSelectSQLAnalizyZakupu = sSelectSQLAnalizyZakupu & " WHERE IDENTKLIENTA='" & oIdentKli & "'"
                DataSetAnalizyZakupu.Locale = New System.Globalization.CultureInfo("pl-PL")
                DataSetAnalizyZakupu = New DataSet
                DataViewAnalizyZakupu = New DataView

                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbConnectionAnalizyZakupu = New OleDb.OleDbConnection(sConnectionAnalizyZakupu)
                    'dbCommandSelectAnalizyZakupu = New OleDb.OleDbCommand(dbSelectSQLAnalizyZakupu, dbConnectionAnalizyZakupu)
                    'dbCommandDeleteAnalizyZakupu = New OleDb.OleDbCommand(sDeleteSQLAnalizyZakupu, dbConnectionAnalizyZakupu)
                    'dbCommandUpdateAnalizyZakupu = New OleDb.OleDbCommand(sUpdateSQLAnalizyZakupu, dbConnectionAnalizyZakupu)
                    'dbCommandInsertAnalizyZakupu = New OleDb.OleDbCommand(sInsertSQLAnalizyZakupu, dbConnectionAnalizyZakupu)
                    'dbDataAdapterAnalizyZakupu = New OleDb.OleDbDataAdapter(dbCommandSelectAnalizyZakupu)

                    ''mapujemy domyslna nazwe tabeli
                    'Dim MapowanieTabeliAnalizyZakupu As System.Data.Common.DataTableMapping
                    'MapowanieTabeliAnalizyZakupu = dbDataAdapterAnalizyZakupu.TableMappings.Add("Table", "TABELA_AnalizyZakupu")

                    'DBDeleteAnalizyZakupu(dbCommandDeleteAnalizyZakupu, dbDataAdapterAnalizyZakupu)
                    'DBUpdateAnalizyZakupu(dbCommandUpdateAnalizyZakupu, dbDataAdapterAnalizyZakupu)
                    'DBInsertAnalizyZakupu(dbCommandInsertAnalizyZakupu, dbDataAdapterAnalizyZakupu)

                    '' Add the RowUpdated event handler.
                    'AddHandler dbDataAdapterAnalizyZakupu.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

                    ''------------------------------------------
                    ''wypelnij DATASET
                    'Try
                    '    dbConnectionAnalizyZakupu.Open()
                    'Catch Ex As System.Exception
                    '    ViewInLog(Ex, Me.Name(), Nothing)
                    '    'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                    '    '    System.Windows.Forms.MessageBoxButtons.OK, _
                    '    '    MessageBoxIcon.Information, _
                    '    '    MessageBoxDefaultButton.Button1)
                    'Finally
                    '    ConnectionState = dbConnectionAnalizyZakupu.State
                    'End Try
                Else
                    sqlConnectionAnalizyZakupu = New SqlClient.SqlConnection(sConnectionAnalizyZakupu)
                    sqlCommandSelectAnalizyZakupu = New SqlClient.SqlCommand(dbSelectSQLAnalizyZakupu, sqlConnectionAnalizyZakupu)
                    sqlCommandDeleteAnalizyZakupu = New SqlClient.SqlCommand(sDeleteSQLAnalizyZakupu, sqlConnectionAnalizyZakupu)
                    sqlCommandUpdateAnalizyZakupu = New SqlClient.SqlCommand(sUpdateSQLAnalizyZakupu, sqlConnectionAnalizyZakupu)
                    sqlCommandInsertAnalizyZakupu = New SqlClient.SqlCommand(sInsertSQLAnalizyZakupu, sqlConnectionAnalizyZakupu)
                    sqlDataAdapterAnalizyZakupu = New SqlClient.SqlDataAdapter(sqlCommandSelectAnalizyZakupu)

                    'mapujemy domyslna nazwe tabeli
                    Dim MapowanieTabeliAnalizyZakupu As System.Data.Common.DataTableMapping
                    MapowanieTabeliAnalizyZakupu = sqlDataAdapterAnalizyZakupu.TableMappings.Add("Table", "TABELA_AnalizyZakupu")

                    SQLDeleteAnalizyZakupu(sqlCommandDeleteAnalizyZakupu, sqlDataAdapterAnalizyZakupu)
                    SQLUpdateAnalizyZakupu(sqlCommandUpdateAnalizyZakupu, sqlDataAdapterAnalizyZakupu)
                    SQLInsertAnalizyZakupu(sqlCommandInsertAnalizyZakupu, sqlDataAdapterAnalizyZakupu)

                    ' Add the RowUpdated event handler.
                    AddHandler sqlDataAdapterAnalizyZakupu.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

                    '------------------------------------------
                    'wypelnij DATASET
                    Try
                        sqlConnectionAnalizyZakupu.Open()
                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    Finally
                        ConnectionState = sqlConnectionAnalizyZakupu.State
                    End Try
                End If

                If ConnectionState = ConnectionState.Open Then
                    Try
                        If ParL_DataBaseType = "MS ACCESS" Then
                            'dbDataAdapterAnalizyZakupu.FillSchema(DataSetAnalizyZakupu, SchemaType.Mapped, "TABELA_AnalizyZakupu")
                            'dbDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
                            'dbConnectionAnalizyZakupu.Close()
                        Else
                            sqlDataAdapterAnalizyZakupu.FillSchema(DataSetAnalizyZakupu, SchemaType.Mapped, "TABELA_AnalizyZakupu")
                            sqlDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
                            sqlConnectionAnalizyZakupu.Close()
                        End If

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        DataSetAnalizyZakupu.Tables(0).PrimaryKey = New DataColumn() {DataSetAnalizyZakupu.Tables(0).Columns("IDENTKLIENTA")}

                        'definiuj DataView
                        DataViewAnalizyZakupu = New DataView(DataSetAnalizyZakupu.Tables(0))
                        DataViewAnalizyZakupu.AllowDelete = False
                        DataViewAnalizyZakupu.AllowNew = False
                        DataViewAnalizyZakupu.AllowEdit = False

                        '-------------------------------------------------
                        'parametry DataGridView
                        '-------------------------------------------------
                        ' Initialize basic DataGridView properties.
                        dagAnalizyZakupu.BackgroundColor = System.Drawing.SystemColors.Control
                        dagAnalizyZakupu.GridColor = System.Drawing.SystemColors.ControlDark
                        dagAnalizyZakupu.ForeColor = System.Drawing.Color.Purple
                        dagAnalizyZakupu.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagAnalizyZakupu.BorderStyle = BorderStyle.Fixed3D
                        'dagAnalizyZakupu.Dock = DockStyle.Fill
                        dagAnalizyZakupu.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                        ' Set property values appropriate for read-only display and limited interactivity. 
                        dagAnalizyZakupu.AllowUserToAddRows = False
                        dagAnalizyZakupu.AllowUserToDeleteRows = False
                        dagAnalizyZakupu.AllowUserToOrderColumns = True
                        dagAnalizyZakupu.AllowUserToResizeColumns = True
                        dagAnalizyZakupu.AllowUserToResizeRows = True
                        dagAnalizyZakupu.ReadOnly = True
                        dagAnalizyZakupu.SelectionMode = DataGridViewSelectionMode.CellSelect
                        dagAnalizyZakupu.MultiSelect = False

                        'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                        'dagAnalizyZakupu.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                        dagAnalizyZakupu.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                        dagAnalizyZakupu.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                        dagAnalizyZakupu.ColumnHeadersVisible = True
                        dagAnalizyZakupu.ColumnHeadersHeight = 40
                        dagAnalizyZakupu.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                        dagAnalizyZakupu.RowHeadersVisible = True
                        dagAnalizyZakupu.RowHeadersWidth = 50
                        dagAnalizyZakupu.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                        ' Set the selection background color for all the cells.
                        dagAnalizyZakupu.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                        dagAnalizyZakupu.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                        dagAnalizyZakupu.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagAnalizyZakupu.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                        dagAnalizyZakupu.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagAnalizyZakupu.DefaultCellStyle.NullValue = ""
                        dagAnalizyZakupu.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        dagAnalizyZakupu.DefaultCellStyle.WrapMode = DataGridViewTriState.True


                        ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                        ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                        dagAnalizyZakupu.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                        dagAnalizyZakupu.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                        ' Set the background color for all rows and for alternating rows.  
                        ' The value for alternating rows overrides the value for all rows. 
                        dagAnalizyZakupu.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagAnalizyZakupu.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                        ' Set the row and column header styles.
                        dagAnalizyZakupu.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagAnalizyZakupu.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        dagAnalizyZakupu.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagAnalizyZakupu.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        '-----------------------------------
                        'wypelnienie DataGridView
                        dagAnalizyZakupu.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                        'dagAnalizyZakupu.DataMember = DataSetKontakty.Tables(0).TableName
                        'dagAnalizyZakupu.DataSource = DataSetKontakty
                        '-----------------------------------

                        ' Add a GridColumnStyle and set the MappingName 
                        ' to the name of a DataColumn in the DataTable. 
                        ' Set the HeaderText and Width properties. 

                        dagAnalizyZakupu.Columns.Clear()

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


                        TxtColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(0).ColumnName, "Klient", 50, DataGridViewContentAlignment.MiddleLeft, True)

                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(1).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(3).ColumnName, "Okres I ", 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(4).ColumnName, "Okres II", 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)

                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(3 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(4 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(5 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(6 + 2).ColumnName, InnyMiesiac(Okres1, -0), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(7 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(8 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(9 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(10 + 2).ColumnName, InnyMiesiac(Okres1, -1), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(11 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(12 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(13 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(14 + 2).ColumnName, InnyMiesiac(Okres1, -2), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(15 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(16 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(17 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(18 + 2).ColumnName, InnyMiesiac(Okres1, -3), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(19 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(20 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(21 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(22 + 2).ColumnName, InnyMiesiac(Okres1, -4), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(23 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(24 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(25 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(26 + 2).ColumnName, InnyMiesiac(Okres1, -5), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(27 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(28 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(29 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(30 + 2).ColumnName, InnyMiesiac(Okres1, -6), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(31 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(32 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(33 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(34 + 2).ColumnName, InnyMiesiac(Okres1, -7), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(35 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(36 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(37 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(38 + 2).ColumnName, InnyMiesiac(Okres1, -8), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(39 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(40 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(41 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(42 + 2).ColumnName, InnyMiesiac(Okres1, -9), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(43 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(44 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(45 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(46 + 2).ColumnName, InnyMiesiac(Okres1, -10), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(47 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(48 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(49 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(50 + 2).ColumnName, InnyMiesiac(Okres1, -11), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)

                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(51 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(52 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(53 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(54 + 2).ColumnName, InnyMiesiac(Okres2, -0), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(55 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(56 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(57 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(58 + 2).ColumnName, InnyMiesiac(Okres2, -1), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(59 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(60 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(61 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(62 + 2).ColumnName, InnyMiesiac(Okres2, -2), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(63 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(64 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(65 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(66 + 2).ColumnName, InnyMiesiac(Okres2, -3), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(67 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(68 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(69 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(70 + 2).ColumnName, InnyMiesiac(Okres2, -4), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(71 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(72 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(73 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(74 + 2).ColumnName, InnyMiesiac(Okres2, -5), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(75 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(76 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(77 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(78 + 2).ColumnName, InnyMiesiac(Okres2, -6), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(79 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(80 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(81 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(82 + 2).ColumnName, InnyMiesiac(Okres2, -7), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(83 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(84 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(85 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(86 + 2).ColumnName, InnyMiesiac(Okres2, -8), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(87 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(88 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(89 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(90 + 2).ColumnName, InnyMiesiac(Okres2, -9), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(91 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(92 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(93 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(94 + 2).ColumnName, InnyMiesiac(Okres2, -10), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(95 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(96 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(97 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(98 + 2).ColumnName, InnyMiesiac(Okres2, -11), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)

                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(99 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(100 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(101 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleCenter, "N2", False)
                        NumColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(102 + 2).ColumnName, InnyMiesiac(Okres2, -12), 60, DataGridViewContentAlignment.MiddleCenter, "N0", True)

                        TxtColumnView(dagAnalizyZakupu, DataSetAnalizyZakupu.Tables(0).Columns(103 + 2).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)

                        dagAnalizyZakupu.Columns(dagAnalizyZakupu.Columns.Count - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                        Me.Cursor = Cursors.WaitCursor
                        dagAnalizyZakupu.DataSource = DataViewAnalizyZakupu
                        Me.Cursor = Cursors.Default

                        ''ustaw sie na pierwszym zapisie
                        If DataSetKontakty.Tables(0).Rows.Count > 0 Then
                            dagAnalizyZakupu.CurrentCell = dagAnalizyZakupu(0, 0)
                            dagAnalizyZakupu.CurrentCell.Selected = True
                        End If


                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    End Try
                    '---------------------------------
                    dagAnalizyZakupu.Focus()

                    'dodaj StatusBar na koncu
                    stbAnalizyZakupu.Panels.Clear()

                    stbAnalizyZakupu.BackColor = System.Drawing.SystemColors.Control
                    stbAnalizyZakupu.ForeColor = System.Drawing.Color.DarkGreen
                    stbAnalizyZakupu.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                    stbAnalizyZakupu.Panels.Add("stbPanelIleRec")
                    stbAnalizyZakupu.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbAnalizyZakupu.Panels(0).Width = 80
                    stbAnalizyZakupu.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbAnalizyZakupu.Panels(0).Text = "Iloœæ rec.: " & DataSetAnalizyZakupu.Tables(0).Rows.Count.ToString

                    stbAnalizyZakupu.Panels.Add("stbPanelWiersz")
                    stbAnalizyZakupu.Panels(1).AutoSize = StatusBarPanelAutoSize.None
                    stbAnalizyZakupu.Panels(1).Width = 80
                    stbAnalizyZakupu.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagAnalizyZakupu.CurrentCell Is Nothing Then
                        stbAnalizyZakupu.Panels(1).Text = "Wi: "
                    Else
                        stbAnalizyZakupu.Panels(1).Text = "Wi: " & dagAnalizyZakupu.CurrentCell.RowIndex.ToString
                    End If

                    stbAnalizyZakupu.Panels.Add("stbPanelKolumna")
                    stbAnalizyZakupu.Panels(2).AutoSize = StatusBarPanelAutoSize.None
                    stbAnalizyZakupu.Panels(2).Width = 80
                    stbAnalizyZakupu.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagAnalizyZakupu.CurrentCell Is Nothing Then
                        stbAnalizyZakupu.Panels(2).Text = "Ko: "
                    Else
                        stbAnalizyZakupu.Panels(2).Text = "Ko: " & dagAnalizyZakupu.CurrentCell.ColumnIndex.ToString
                    End If

                    stbAnalizyZakupu.Panels.Add("stbPanelSort")
                    stbAnalizyZakupu.Panels(3).AutoSize = StatusBarPanelAutoSize.None
                    stbAnalizyZakupu.Panels(3).Width = 150
                    stbAnalizyZakupu.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbAnalizyZakupu.Panels(3).Text = "Sort : "

                    stbAnalizyZakupu.Panels.Add("stbPanelFiltr")
                    stbAnalizyZakupu.Panels(4).AutoSize = StatusBarPanelAutoSize.None
                    stbAnalizyZakupu.Panels(4).Width = 150
                    stbAnalizyZakupu.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbAnalizyZakupu.Panels(4).Text = "Filtr : "

                    stbAnalizyZakupu.Panels.Add("stbPanelReszta")
                    stbAnalizyZakupu.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
                    'stbAnalizyZakupu.Panels(5).Width = 150
                    stbAnalizyZakupu.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbAnalizyZakupu.Panels(5).Text = ""

                    stbAnalizyZakupu.ShowPanels = True
                    '---------------------------------
                    'dodaj StatusBar na koncu
                    stbFiltryAnalizyZakupu.Panels.Clear()

                    stbFiltryAnalizyZakupu.BackColor = System.Drawing.SystemColors.Control
                    stbFiltryAnalizyZakupu.ForeColor = System.Drawing.Color.DarkGreen
                    stbFiltryAnalizyZakupu.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                    stbFiltryAnalizyZakupu.Panels.Add("stbFiltryRowHead")
                    stbFiltryAnalizyZakupu.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbFiltryAnalizyZakupu.Panels(0).Width = 50  'dagAnalizyZakupu.TableStyles(0).RowHeaderWidth
                    stbFiltryAnalizyZakupu.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbFiltryAnalizyZakupu.Panels(0).Text = "Filtry"

                    stbFiltryAnalizyZakupu.Panels.Add("stbFiltryReszta")
                    stbFiltryAnalizyZakupu.Panels(1).AutoSize = StatusBarPanelAutoSize.Spring
                    'stbFiltryAnalizyZakupu.Panels(1).Width = 150
                    stbFiltryAnalizyZakupu.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbFiltryAnalizyZakupu.Panels(1).Text = ""

                    stbFiltryAnalizyZakupu.ShowPanels = True

                    'dodaj ctrl do pobierania szablonu
                    cmdWszystkoAnalizyZakupu.Size = New Size(20, 22)
                    cmdWszystkoAnalizyZakupu.Location = New System.Drawing.Point(stbFiltryAnalizyZakupu.Location.X + 50 - 20,
                                             stbFiltryAnalizyZakupu.Location.Y)
                    cmdWszystkoAnalizyZakupu.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)

                End If
                InitAnalizyZakupu() 'inicjuj zmienne

                '--------------------------------------------------------------------
                'set the header width to something apporpriate
                dagAnalizyZakupu.RowHeadersWidth = 50       'use if tablestyle
                'Me.dagAnalizyZakupu.RowHeaderWidth = 80 'use if no tablestyle
                '--------------------------------------------------------------------
                ''set topleft corner point so we can locate the toprow
                'If DataSetAnalizyZakupu.Tables(0).Rows.Count > 0 Then
                '    StartPointInCell00 = New System.Drawing.Point((dagAnalizyZakupu.GetCellDisplayRectangle(0, 0, True).X + 4),
                '                      (dagAnalizyZakupu.GetCellDisplayRectangle(0, 0, True).Y + 4))
                '    ScrollXOffset = 10000 * dagAnalizyZakupu.FirstDisplayedScrollingColumnIndex +
                '        dagAnalizyZakupu.GetCellDisplayRectangle(dagAnalizyZakupu.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                '        StartPointInCell00.X
                'Else
                '    'brak zapisow - poczatek DataGrid
                '    StartPointInCell00 = New System.Drawing.Point((dagAnalizyZakupu.Left + 4), (dagAnalizyZakupu.Top + 4))
                '    ScrollXOffset = 0
                'End If
                ''--------------------------------------------------
                ''inicjowanie zegara dla scrollingu poziomego
                'HorizScrollLastTime = ""
                'HorizScrollTimer.Interval = 2 * 1000
                'HorizScrollTimer.Enabled = True
                '--------------------------------------------------
                'Stworz zmienne do filtrowania
                Dim ii As Integer = 0


                For ii = 0 To DataViewAnalizyZakupu.Table.Columns.Count - 1
                    'stworz zmienn¹ TXTBOX
                    Dim CtrlT As New System.Windows.Forms.TextBox
                    CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                    CtrlT.Visible = False
                    Me.AnalizyZakupu.Controls.Add(CtrlT)
                    AddHandler CtrlT.TextChanged, New EventHandler(AddressOf txtFiXXAnalizyZakupu_TextChanged)
                    AddHandler CtrlT.GotFocus, New EventHandler(AddressOf txtFiXXAnalizyZakupu_GotFocus)
                    AddHandler CtrlT.LostFocus, New EventHandler(AddressOf txtFiXXAnalizyZakupu_LostFocus)
                    AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXXAnalizyZakupu_KeyDown)

                    'stworz zmienn¹ BUTTON
                    Dim CtrlB As New System.Windows.Forms.Button
                    CtrlB.Name = "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                    CtrlB.Image = cmdFi.Image
                    CtrlB.ImageAlign = ContentAlignment.MiddleCenter
                    CtrlB.Visible = False
                    Me.AnalizyZakupu.Controls.Add(CtrlB)
                    AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXXAnalizyZakupu_Click)
                    AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXXAnalizyZakupu_GotFocus)
                    AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXXAnalizyZakupu_LostFocus)
                Next
                Me.Refresh()
                '--------------------------------------------------
                StartFormy = False
                StartFormyAnalizyZakupu = False
                dagAnalizyZakupu_CurrentCellChanged(sender, e)
                InicjujFiltryColumnDGV(Me.AnalizyZakupu, dagAnalizyZakupu, DataViewAnalizyZakupu.Table.Columns.Count, stbFiltryAnalizyZakupu, StartFormyAnalizyZakupu)
                '--------------------------------------------------------------------
                cmdPowrot.Focus()







            Case 5  'obroty miesiêczne
                lblIdent_ObrotyMies.Text = txtIdent.Text
                lblTofi_ObrotyMies.Text = txtTofi.Text
                lblNazwaHandlowa_ObrotyMies.Text = txtNazwa1.Text
                lblNazwa1_ObrotyMies.Text = Trim(txtKodP.Text) & " " & Trim(txtMiejscowosc.Text) & ", " & Trim(txtUlica.Text) & " " & Trim(txtNrDomu.Text) & " " & Trim(txtIDDomu.Text)
                lblNazwa2_ObrotyMies.Text = OsobaKontaktowa(txtIdent.Text)
                lblStanowisko_ObrotyMies.Text = oStanowiskoOso
                lblTelefon_ObrotyMies.Text = oTelefonOso
                lblTelefonKom_ObrotyMies.Text = oTelefonKomOso

                cmdKontaktPromotor.Visible = False
                cmdCoDalej.Visible = False
                '---------------------------------------------
                'Obroty w miesi¹cu
                StartFormyObrotyMies = True
                dbSelectSQLObrotyMies = baseSelectSQLObrotyMies & " WHERE ObrotyMies.IDENTKLIENTA='" & oIdentKli & "' ORDER BY ObrotyMies.MIESIAC DESC"

                DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")
                DataSetObrotyMies = New DataSet
                DataViewObrotyMies = New DataView

                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbConnectionObrotyMies = New OleDb.OleDbConnection(sConnectionObrotyMies)
                    'dbCommandSelectObrotyMies = New OleDb.OleDbCommand(dbSelectSQLObrotyMies, dbConnectionObrotyMies)
                    'dbCommandDeleteObrotyMies = New OleDb.OleDbCommand(sDeleteSQLObrotyMies, dbConnectionObrotyMies)
                    'dbCommandUpdateObrotyMies = New OleDb.OleDbCommand(sUpdateSQLObrotyMies, dbConnectionObrotyMies)
                    'dbCommandInsertObrotyMies = New OleDb.OleDbCommand(sInsertSQLObrotyMies, dbConnectionObrotyMies)
                    'dbDataAdapterObrotyMies = New OleDb.OleDbDataAdapter(dbCommandSelectObrotyMies)

                    ''mapujemy domyslna nazwe tabeli
                    'Dim MapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
                    'MapowanieTabeliObrotyMies = dbDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")

                    'DBDeleteObrotyMies(dbCommandDeleteObrotyMies, dbDataAdapterObrotyMies)
                    'DBUpdateObrotyMies(dbCommandUpdateObrotyMies, dbDataAdapterObrotyMies)
                    'DBInsertObrotyMies(dbCommandInsertObrotyMies, dbDataAdapterObrotyMies)

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
                    'Finally
                    '    ConnectionState = dbConnectionObrotyMies.State
                    'End Try
                Else
                    sqlConnectionObrotyMies = New SqlClient.SqlConnection(sConnectionObrotyMies)
                    sqlCommandSelectObrotyMies = New SqlClient.SqlCommand(dbSelectSQLObrotyMies, sqlConnectionObrotyMies)
                    sqlCommandDeleteObrotyMies = New SqlClient.SqlCommand(sDeleteSQLObrotyMies, sqlConnectionObrotyMies)
                    sqlCommandUpdateObrotyMies = New SqlClient.SqlCommand(sUpdateSQLObrotyMies, sqlConnectionObrotyMies)
                    sqlCommandInsertObrotyMies = New SqlClient.SqlCommand(sInsertSQLObrotyMies, sqlConnectionObrotyMies)
                    sqlDataAdapterObrotyMies = New SqlClient.SqlDataAdapter(sqlCommandSelectObrotyMies)

                    'mapujemy domyslna nazwe tabeli
                    Dim MapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
                    MapowanieTabeliObrotyMies = sqlDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")

                    SQLDeleteObrotyMies(sqlCommandDeleteObrotyMies, sqlDataAdapterObrotyMies)
                    SQLUpdateObrotyMies(sqlCommandUpdateObrotyMies, sqlDataAdapterObrotyMies)
                    SQLInsertObrotyMies(sqlCommandInsertObrotyMies, sqlDataAdapterObrotyMies)

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
                    Finally
                        ConnectionState = sqlConnectionObrotyMies.State
                    End Try
                End If

                If ConnectionState = ConnectionState.Open Then
                    Try
                        If ParL_DataBaseType = "MS ACCESS" Then
                            'dbDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                            'dbDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                            'dbConnectionObrotyMies.Close()
                        Else
                            sqlDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                            sqlDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                            sqlConnectionObrotyMies.Close()
                        End If

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        DataSetObrotyMies.Tables(0).PrimaryKey = New DataColumn() {DataSetObrotyMies.Tables(0).Columns("IDENTKLIENTA"), DataSetObrotyMies.Tables(0).Columns("MIESIAC")}

                        'definiuj DataView
                        DataViewObrotyMies = New DataView(DataSetObrotyMies.Tables(0))
                        DataViewObrotyMies.AllowDelete = False
                        DataViewObrotyMies.AllowNew = False
                        DataViewObrotyMies.AllowEdit = False

                        '-------------------------------------------------
                        'parametry DataGridView
                        '-------------------------------------------------
                        ' Initialize basic DataGridView properties.
                        dagObrotyMies.BackgroundColor = System.Drawing.SystemColors.Control
                        dagObrotyMies.GridColor = System.Drawing.SystemColors.ControlDark
                        dagObrotyMies.ForeColor = System.Drawing.Color.Purple
                        dagObrotyMies.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagObrotyMies.BorderStyle = BorderStyle.Fixed3D
                        'dagObrotyMies.Dock = DockStyle.Fill
                        dagObrotyMies.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                        ' Set property values appropriate for read-only display and limited interactivity. 
                        dagObrotyMies.AllowUserToAddRows = False
                        dagObrotyMies.AllowUserToDeleteRows = False
                        dagObrotyMies.AllowUserToOrderColumns = True
                        dagObrotyMies.AllowUserToResizeColumns = True
                        dagObrotyMies.AllowUserToResizeRows = True
                        dagObrotyMies.ReadOnly = True
                        dagObrotyMies.SelectionMode = DataGridViewSelectionMode.CellSelect
                        dagObrotyMies.MultiSelect = False

                        'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                        'dagObrotyMies.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                        dagObrotyMies.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                        dagObrotyMies.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                        dagObrotyMies.ColumnHeadersVisible = True
                        dagObrotyMies.ColumnHeadersHeight = 60
                        dagObrotyMies.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                        dagObrotyMies.RowHeadersVisible = True
                        dagObrotyMies.RowHeadersWidth = 50
                        dagObrotyMies.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                        ' Set the selection background color for all the cells.
                        dagObrotyMies.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                        dagObrotyMies.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                        dagObrotyMies.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagObrotyMies.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                        dagObrotyMies.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagObrotyMies.DefaultCellStyle.NullValue = ""
                        dagObrotyMies.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        dagObrotyMies.DefaultCellStyle.WrapMode = DataGridViewTriState.True


                        ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                        ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                        dagObrotyMies.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                        dagObrotyMies.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                        ' Set the background color for all rows and for alternating rows.  
                        ' The value for alternating rows overrides the value for all rows. 
                        dagObrotyMies.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagObrotyMies.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                        ' Set the row and column header styles.
                        dagObrotyMies.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagObrotyMies.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        dagObrotyMies.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagObrotyMies.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        '-----------------------------------
                        'wypelnienie DataGridView
                        dagObrotyMies.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                        'dagObrotyMies.DataMember = DataSetObrotyMies.Tables(0).TableName
                        'dagObrotyMies.DataSource = DataSetOsobty
                        '-----------------------------------

                        ' Add a GridColumnStyle and set the MappingName 
                        ' to the name of a DataColumn in the DataTable. 
                        ' Set the HeaderText and Width properties. 

                        dagObrotyMies.Columns.Clear()

                        TxtColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(0).ColumnName, "Klient", 50, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(1).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(2).ColumnName, "Miesiac", 80, DataGridViewContentAlignment.MiddleLeft, True)

                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(8).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(7).ColumnName, "PRYZMAT Laser Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(6).ColumnName, "PRYZMAT Laser Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(5).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(4).ColumnName, "PRYZMAT Atrament Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(3).ColumnName, "PRYZMAT Atrament Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(14).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(13).ColumnName, "ORYGINA£Y Laser Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(12).ColumnName, "ORYGINA£Y Laser Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(11).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(10).ColumnName, "ORYGINA£Y Atrament Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(9).ColumnName, "ORYGINA£Y Atrament Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(17).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(16).ColumnName, "Najem Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(15).ColumnName, "Najem Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(20).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(19).ColumnName, "Strony Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(18).ColumnName, "Strony Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(23).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(22).ColumnName, "Drukarki Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(21).ColumnName, "Drukarki Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(26).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(25).ColumnName, "Skup Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(24).ColumnName, "Skup Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(29).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(28).ColumnName, "Suma Wartoœci", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(27).ColumnName, "Suma Iloœci", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        TxtColumnView(dagObrotyMies, DataSetObrotyMies.Tables(0).Columns(30).ColumnName, "", 0, HorizontalAlignment.Left, False)

                        dagObrotyMies.Columns(dagObrotyMies.Columns.Count - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                        Me.Cursor = Cursors.WaitCursor
                        dagObrotyMies.DataSource = DataViewObrotyMies
                        Me.Cursor = Cursors.Default

                        ''ustaw sie na pierwszym zapisie
                        If DataSetObrotyMies.Tables(0).Rows.Count > 0 Then
                            dagObrotyMies.CurrentCell = dagObrotyMies(2, 0)
                            dagObrotyMies.CurrentCell.Selected = True
                        End If


                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    End Try
                    '---------------------------------
                    '---------------------------------
                    dagObrotyMies.Focus()
                    '---------------------------------
                    'dodaj StatusBar na koncu
                    stbObrotyMies.Panels.Clear()

                    stbObrotyMies.BackColor = System.Drawing.SystemColors.Control
                    stbObrotyMies.ForeColor = System.Drawing.Color.DarkGreen
                    stbObrotyMies.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                    stbObrotyMies.Panels.Add("stbPanelIleRec")
                    stbObrotyMies.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbObrotyMies.Panels(0).Width = 80
                    stbObrotyMies.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbObrotyMies.Panels(0).Text = "Iloœæ rec.: " & DataSetObrotyMies.Tables(0).Rows.Count.ToString

                    stbObrotyMies.Panels.Add("stbPanelWiersz")
                    stbObrotyMies.Panels(1).AutoSize = StatusBarPanelAutoSize.None
                    stbObrotyMies.Panels(1).Width = 80
                    stbObrotyMies.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagObrotyMies.CurrentCell Is Nothing Then
                        stbObrotyMies.Panels(1).Text = "Wi: "
                    Else
                        stbObrotyMies.Panels(1).Text = "Wi: " & dagObrotyMies.CurrentCell.RowIndex.ToString
                    End If

                    stbObrotyMies.Panels.Add("stbPanelKolumna")
                    stbObrotyMies.Panels(2).AutoSize = StatusBarPanelAutoSize.None
                    stbObrotyMies.Panels(2).Width = 80
                    stbObrotyMies.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagObrotyMies.CurrentCell Is Nothing Then
                        stbObrotyMies.Panels(2).Text = "Ko: "
                    Else
                        stbObrotyMies.Panels(2).Text = "Ko: " & dagObrotyMies.CurrentCell.ColumnIndex.ToString
                    End If

                    stbObrotyMies.Panels.Add("stbPanelSort")
                    stbObrotyMies.Panels(3).AutoSize = StatusBarPanelAutoSize.None
                    stbObrotyMies.Panels(3).Width = 150
                    stbObrotyMies.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbObrotyMies.Panels(3).Text = "Sort : "

                    stbObrotyMies.Panels.Add("stbPanelFiltr")
                    stbObrotyMies.Panels(4).AutoSize = StatusBarPanelAutoSize.None
                    stbObrotyMies.Panels(4).Width = 150
                    stbObrotyMies.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbObrotyMies.Panels(4).Text = "Filtr : "

                    stbObrotyMies.Panels.Add("stbPanelReszta")
                    stbObrotyMies.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
                    'stbobrotyMIES.Panels(5).Width = 150
                    stbObrotyMies.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbObrotyMies.Panels(5).Text = ""

                    stbObrotyMies.ShowPanels = True
                    '---------------------------------
                    'dodaj StatusBar na koncu
                    stbFiltryObrotyMies.Panels.Clear()

                    stbFiltryObrotyMies.BackColor = System.Drawing.SystemColors.Control
                    stbFiltryObrotyMies.ForeColor = System.Drawing.Color.DarkGreen
                    stbFiltryObrotyMies.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                    stbFiltryObrotyMies.Panels.Add("stbFiltryRowHead")
                    stbFiltryObrotyMies.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbFiltryObrotyMies.Panels(0).Width = 50  'dagobrotyMIES.TableStyles(0).RowHeaderWidth
                    stbFiltryObrotyMies.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbFiltryObrotyMies.Panels(0).Text = "Filtry"

                    stbFiltryObrotyMies.Panels.Add("stbFiltryReszta")
                    stbFiltryObrotyMies.Panels(1).AutoSize = StatusBarPanelAutoSize.Spring
                    'stbFiltryobrotyMIES.Panels(1).Width = 150
                    stbFiltryObrotyMies.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbFiltryObrotyMies.Panels(1).Text = ""

                    stbFiltryObrotyMies.ShowPanels = True

                    'dodaj ctrl do pobierania szablonu
                    cmdWszystkoObrotyMies.Size = New Size(20, 22)
                    cmdWszystkoObrotyMies.Location = New System.Drawing.Point(stbFiltryObrotyMies.Location.X + 50 - 20,
                                             stbFiltryObrotyMies.Location.Y)
                    cmdWszystkoObrotyMies.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)

                End If
                InitObrotyMies() 'inicjuj zmienne

                '--------------------------------------------------------------------
                'set the header width to something apporpriate
                dagObrotyMies.RowHeadersWidth = 50       'use if tablestyle
                'Me.dagobrotyMIES.RowHeaderWidth = 80 'use if no tablestyle
                '--------------------------------------------------------------------
                ''set topleft corner point so we can locate the toprow
                'If DataSetobrotyMIES.Tables(0).Rows.Count > 0 Then
                '    StartPointInCell00 = New System.Drawing.Point((dagobrotyMIES.GetCellDisplayRectangle(0, 0, True).X + 4),
                '                      (dagobrotyMIES.GetCellDisplayRectangle(0, 0, True).Y + 4))
                '    ScrollXOffset = 10000 * dagobrotyMIES.FirstDisplayedScrollingColumnIndex +
                '        dagobrotyMIES.GetCellDisplayRectangle(dagobrotyMIES.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                '        StartPointInCell00.X
                'Else
                '    'brak zapisow - poczatek DataGrid
                '    StartPointInCell00 = New System.Drawing.Point((dagobrotyMIES.Left + 4), (dagobrotyMIES.Top + 4))
                '    ScrollXOffset = 0
                'End If
                ''--------------------------------------------------
                ''inicjowanie zegara dla scrollingu poziomego
                'HorizScrollLastTime = ""
                'HorizScrollTimer.Interval = 2 * 1000
                'HorizScrollTimer.Enabled = True
                '--------------------------------------------------
                'Stworz zmienne do filtrowania
                Dim ii As Integer = 0
                For ii = 0 To DataViewObrotyMies.Table.Columns.Count - 1
                    'stworz zmienn¹ TXTBOX
                    Dim CtrlT As New System.Windows.Forms.TextBox
                    CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                    CtrlT.Visible = False
                    Me.ObrotyMies.Controls.Add(CtrlT)
                    AddHandler CtrlT.TextChanged, New EventHandler(AddressOf txtFiXXObrotyMies_TextChanged)
                    AddHandler CtrlT.GotFocus, New EventHandler(AddressOf txtFiXXObrotyMies_GotFocus)
                    AddHandler CtrlT.LostFocus, New EventHandler(AddressOf txtFiXXObrotyMies_LostFocus)
                    AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXXObrotyMies_KeyDown)

                    'stworz zmienn¹ BUTTON
                    Dim CtrlB As New System.Windows.Forms.Button
                    CtrlB.Name = "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                    CtrlB.Image = cmdFi.Image
                    CtrlB.ImageAlign = ContentAlignment.MiddleCenter
                    CtrlB.Visible = False
                    Me.ObrotyMies.Controls.Add(CtrlB)
                    AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXXObrotyMies_Click)
                    AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXXObrotyMies_GotFocus)
                    AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXXObrotyMies_LostFocus)
                Next
                Me.Refresh()
                '--------------------------------------------------
                StartFormy = False
                StartFormyObrotyMies = False
                dagObrotyMies_CurrentCellChanged(sender, e)
                InicjujFiltryColumnDGV(Me.ObrotyMies, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltryObrotyMies, StartFormyObrotyMies)
                '--------------------------------------------------------------------
                cmdPowrot.Focus()







            Case 6  'obroty
                lblIdent_Obroty.Text = txtIdent.Text
                lblTofi_Obroty.Text = txtTofi.Text
                lblNazwaHandlowa_Obroty.Text = txtNazwa1.Text
                lblNazwa1_Obroty.Text = Trim(txtKodP.Text) & " " & Trim(txtMiejscowosc.Text) & ", " & Trim(txtUlica.Text) & " " & Trim(txtNrDomu.Text) & " " & Trim(txtIDDomu.Text)
                lblNazwa2_Obroty.Text = OsobaKontaktowa(txtIdent.Text)
                lblStanowisko_Obroty.Text = oStanowiskoOso
                lblTelefon_Obroty.Text = oTelefonOso
                lblTelefonKom_Obroty.Text = oTelefonKomOso

                cmdKontaktPromotor.Visible = False
                cmdCoDalej.Visible = False
                '---------------------------------------------
                'Obroty
                StartFormyObroty = True
                dbSelectSQLObroty = baseSelectSQLObroty & " WHERE Obroty.IDENTKLIENTA='" & oIdentKli & "' ORDER BY Obroty.DATA DESC"

                DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")
                DataSetObroty = New DataSet
                DataViewObroty = New DataView

                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbConnectionObroty = New OleDb.OleDbConnection(sConnectionObroty)
                    'dbCommandSelectObroty = New OleDb.OleDbCommand(dbSelectSQLObroty, dbConnectionObroty)
                    'dbCommandDeleteObroty = New OleDb.OleDbCommand(sDeleteSQLObroty, dbConnectionObroty)
                    'dbCommandUpdateObroty = New OleDb.OleDbCommand(sUpdateSQLObroty, dbConnectionObroty)
                    'dbCommandInsertObroty = New OleDb.OleDbCommand(sInsertSQLObroty, dbConnectionObroty)
                    'dbDataAdapterObroty = New OleDb.OleDbDataAdapter(dbCommandSelectObroty)

                    ''mapujemy domyslna nazwe tabeli
                    'Dim MapowanieTabeliObroty As System.Data.Common.DataTableMapping
                    'MapowanieTabeliObroty = dbDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")

                    'DBDeleteObroty(dbCommandDeleteObroty, dbDataAdapterObroty)
                    'DBUpdateObroty(dbCommandUpdateObroty, dbDataAdapterObroty)
                    'DBInsertObroty(dbCommandInsertObroty, dbDataAdapterObroty)

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
                    'Finally
                    '    ConnectionState = dbConnectionObroty.State
                    'End Try
                Else
                    sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionObroty)
                    sqlCommandSelectObroty = New SqlClient.SqlCommand(dbSelectSQLObroty, sqlConnectionObroty)
                    sqlCommandDeleteObroty = New SqlClient.SqlCommand(sDeleteSQLObroty, sqlConnectionObroty)
                    sqlCommandUpdateObroty = New SqlClient.SqlCommand(sUpdateSQLObroty, sqlConnectionObroty)
                    sqlCommandInsertObroty = New SqlClient.SqlCommand(sInsertSQLObroty, sqlConnectionObroty)
                    sqlDataAdapterObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)

                    'mapujemy domyslna nazwe tabeli
                    Dim MapowanieTabeliObroty As System.Data.Common.DataTableMapping
                    MapowanieTabeliObroty = sqlDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")

                    SQLDeleteObroty(sqlCommandDeleteObroty, sqlDataAdapterObroty)
                    SQLUpdateObroty(sqlCommandUpdateObroty, sqlDataAdapterObroty)
                    SQLInsertObroty(sqlCommandInsertObroty, sqlDataAdapterObroty)

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

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        DataSetObroty.Tables(0).PrimaryKey = New DataColumn() {DataSetObroty.Tables(0).Columns("IDENTKLIENTA"), DataSetObroty.Tables(0).Columns("DATA")}

                        'definiuj DataView
                        DataViewObroty = New DataView(DataSetObroty.Tables(0))
                        DataViewObroty.AllowDelete = False
                        DataViewObroty.AllowNew = False
                        DataViewObroty.AllowEdit = False

                        '-------------------------------------------------
                        'parametry DataGridView
                        '-------------------------------------------------
                        ' Initialize basic DataGridView properties.
                        dagObroty.BackgroundColor = System.Drawing.SystemColors.Control
                        dagObroty.GridColor = System.Drawing.SystemColors.ControlDark
                        dagObroty.ForeColor = System.Drawing.Color.Purple
                        dagObroty.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagObroty.BorderStyle = BorderStyle.Fixed3D
                        'dagObroty.Dock = DockStyle.Fill
                        dagObroty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                        ' Set property values appropriate for read-only display and limited interactivity. 
                        dagObroty.AllowUserToAddRows = False
                        dagObroty.AllowUserToDeleteRows = False
                        dagObroty.AllowUserToOrderColumns = True
                        dagObroty.AllowUserToResizeColumns = True
                        dagObroty.AllowUserToResizeRows = True
                        dagObroty.ReadOnly = True
                        dagObroty.SelectionMode = DataGridViewSelectionMode.CellSelect
                        dagObroty.MultiSelect = False

                        'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                        'dagobroty.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                        dagObroty.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                        dagObroty.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                        dagObroty.ColumnHeadersVisible = True
                        dagObroty.ColumnHeadersHeight = 60
                        dagObroty.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                        dagObroty.RowHeadersVisible = True
                        dagObroty.RowHeadersWidth = 50
                        dagObroty.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                        ' Set the selection background color for all the cells.
                        dagObroty.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                        dagObroty.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                        dagObroty.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagObroty.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                        dagObroty.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagObroty.DefaultCellStyle.NullValue = ""
                        dagObroty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        dagObroty.DefaultCellStyle.WrapMode = DataGridViewTriState.True


                        ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                        ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                        dagObroty.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                        dagObroty.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                        ' Set the background color for all rows and for alternating rows.  
                        ' The value for alternating rows overrides the value for all rows. 
                        dagObroty.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagObroty.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                        ' Set the row and column header styles.
                        dagObroty.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagObroty.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        dagObroty.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagObroty.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        '-----------------------------------
                        'wypelnienie DataGridView
                        dagObroty.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                        'dagobroty.DataMember = DataSetobroty.Tables(0).TableName
                        'dagobroty.DataSource = DataSetOsobty
                        '-----------------------------------

                        ' Add a GridColumnStyle and set the MappingName 
                        ' to the name of a DataColumn in the DataTable. 
                        ' Set the HeaderText and Width properties. 

                        dagObroty.Columns.Clear()

                        TxtColumnView(dagObroty, DataSetObroty.Tables(0).Columns(0).ColumnName, "Klient", 50, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagObroty, DataSetObroty.Tables(0).Columns(1).ColumnName, "Nr TOFI", 100, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagObroty, DataSetObroty.Tables(0).Columns(2).ColumnName, "Data", 80, DataGridViewContentAlignment.MiddleLeft, True)

                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(8).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(7).ColumnName, "PRYZMAT Laser Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(6).ColumnName, "PRYZMAT Laser Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(5).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(4).ColumnName, "PRYZMAT Atrament Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(3).ColumnName, "PRYZMAT Atrament Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(14).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(13).ColumnName, "ORYGINA£Y Laser Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(12).ColumnName, "ORYGINA£Y Laser Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(11).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(10).ColumnName, "ORYGINA£Y Atrament Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(9).ColumnName, "ORYGINA£Y Atrament Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(17).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(16).ColumnName, "Najem Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(15).ColumnName, "Najem Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(20).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(19).ColumnName, "Strony Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(18).ColumnName, "Strony Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(23).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(22).ColumnName, "Drukarki Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(21).ColumnName, "Drukarki Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(26).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(25).ColumnName, "Skup Wartoœæ", 100, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(24).ColumnName, "Skup Iloœæ", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(29).ColumnName, "", 0, DataGridViewContentAlignment.MiddleRight, "N2", False)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(28).ColumnName, "Suma Wartoœci", 100, DataGridViewContentAlignment.MiddleRight, "N2", True)
                        NumColumnView(dagObroty, DataSetObroty.Tables(0).Columns(27).ColumnName, "Suma Iloœci", 100, DataGridViewContentAlignment.MiddleRight, "N0", False)

                        TxtColumnView(dagObroty, DataSetObroty.Tables(0).Columns(30).ColumnName, "", 0, HorizontalAlignment.Left, False)

                        dagObroty.Columns(dagObroty.Columns.Count - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

                        Me.Cursor = Cursors.WaitCursor
                        dagObroty.DataSource = DataViewObroty
                        Me.Cursor = Cursors.Default

                        ''ustaw sie na pierwszym zapisie
                        If DataSetObroty.Tables(0).Rows.Count > 0 Then
                            dagObroty.CurrentCell = dagObroty(2, 0)
                            dagObroty.CurrentCell.Selected = True
                        End If

                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    End Try
                    '---------------------------------
                    '---------------------------------
                    dagObroty.Focus()
                    '---------------------------------
                    'dodaj StatusBar na koncu
                    stbObroty.Panels.Clear()

                    stbObroty.BackColor = System.Drawing.SystemColors.Control
                    stbObroty.ForeColor = System.Drawing.Color.DarkGreen
                    stbObroty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                    stbObroty.Panels.Add("stbPanelIleRec")
                    stbObroty.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbObroty.Panels(0).Width = 80
                    stbObroty.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbObroty.Panels(0).Text = "Iloœæ rec.: " & DataSetObroty.Tables(0).Rows.Count.ToString

                    stbObroty.Panels.Add("stbPanelWiersz")
                    stbObroty.Panels(1).AutoSize = StatusBarPanelAutoSize.None
                    stbObroty.Panels(1).Width = 80
                    stbObroty.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagObroty.CurrentCell Is Nothing Then
                        stbObroty.Panels(1).Text = "Wi: "
                    Else
                        stbObroty.Panels(1).Text = "Wi: " & dagObroty.CurrentCell.RowIndex.ToString
                    End If

                    stbObroty.Panels.Add("stbPanelKolumna")
                    stbObroty.Panels(2).AutoSize = StatusBarPanelAutoSize.None
                    stbObroty.Panels(2).Width = 80
                    stbObroty.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagObroty.CurrentCell Is Nothing Then
                        stbObroty.Panels(2).Text = "Ko: "
                    Else
                        stbObroty.Panels(2).Text = "Ko: " & dagObroty.CurrentCell.ColumnIndex.ToString
                    End If

                    stbObroty.Panels.Add("stbPanelSort")
                    stbObroty.Panels(3).AutoSize = StatusBarPanelAutoSize.None
                    stbObroty.Panels(3).Width = 150
                    stbObroty.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbObroty.Panels(3).Text = "Sort : "

                    stbObroty.Panels.Add("stbPanelFiltr")
                    stbObroty.Panels(4).AutoSize = StatusBarPanelAutoSize.None
                    stbObroty.Panels(4).Width = 150
                    stbObroty.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbObroty.Panels(4).Text = "Filtr : "

                    stbObroty.Panels.Add("stbPanelReszta")
                    stbObroty.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
                    'stbobroty.Panels(5).Width = 150
                    stbObroty.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbObroty.Panels(5).Text = ""

                    stbObroty.ShowPanels = True
                    '---------------------------------
                    'dodaj StatusBar na koncu
                    stbFiltryObroty.Panels.Clear()

                    stbFiltryObroty.BackColor = System.Drawing.SystemColors.Control
                    stbFiltryObroty.ForeColor = System.Drawing.Color.DarkGreen
                    stbFiltryObroty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                    stbFiltryObroty.Panels.Add("stbFiltryRowHead")
                    stbFiltryObroty.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbFiltryObroty.Panels(0).Width = 50  'dagobroty.TableStyles(0).RowHeaderWidth
                    stbFiltryObroty.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbFiltryObroty.Panels(0).Text = "Filtry"

                    stbFiltryObroty.Panels.Add("stbFiltryReszta")
                    stbFiltryObroty.Panels(1).AutoSize = StatusBarPanelAutoSize.Spring
                    'stbFiltryobroty.Panels(1).Width = 150
                    stbFiltryObroty.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbFiltryObroty.Panels(1).Text = ""

                    stbFiltryObroty.ShowPanels = True

                    'dodaj ctrl do pobierania szablonu
                    cmdWszystkoObroty.Size = New Size(20, 22)
                    cmdWszystkoObroty.Location = New System.Drawing.Point(stbFiltryObroty.Location.X + 50 - 20,
                                             stbFiltryObroty.Location.Y)
                    cmdWszystkoObroty.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)

                End If
                InitObroty() 'inicjuj zmienne

                '--------------------------------------------------------------------
                'set the header width to something apporpriate
                dagObroty.RowHeadersWidth = 50       'use if tablestyle
                'Me.dagobroty.RowHeaderWidth = 80 'use if no tablestyle
                '--------------------------------------------------------------------
                ''set topleft corner point so we can locate the toprow
                'If DataSetobroty.Tables(0).Rows.Count > 0 Then
                '    StartPointInCell00 = New System.Drawing.Point((dagobroty.GetCellDisplayRectangle(0, 0, True).X + 4),
                '                      (dagobroty.GetCellDisplayRectangle(0, 0, True).Y + 4))
                '    ScrollXOffset = 10000 * dagobroty.FirstDisplayedScrollingColumnIndex +
                '        dagobroty.GetCellDisplayRectangle(dagobroty.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                '        StartPointInCell00.X
                'Else
                '    'brak zapisow - poczatek DataGrid
                '    StartPointInCell00 = New System.Drawing.Point((dagobroty.Left + 4), (dagobroty.Top + 4))
                '    ScrollXOffset = 0
                'End If
                ''--------------------------------------------------
                ''inicjowanie zegara dla scrollingu poziomego
                'HorizScrollLastTime = ""
                'HorizScrollTimer.Interval = 2 * 1000
                'HorizScrollTimer.Enabled = True
                '--------------------------------------------------
                'Stworz zmienne do filtrowania
                Dim ii As Integer = 0
                For ii = 0 To DataViewObroty.Table.Columns.Count - 1
                    'stworz zmienn¹ TXTBOX
                    Dim CtrlT As New System.Windows.Forms.TextBox
                    CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                    CtrlT.Visible = False
                    Me.Obroty.Controls.Add(CtrlT)
                    AddHandler CtrlT.TextChanged, New EventHandler(AddressOf txtFiXXObroty_TextChanged)
                    AddHandler CtrlT.GotFocus, New EventHandler(AddressOf txtFiXXObroty_GotFocus)
                    AddHandler CtrlT.LostFocus, New EventHandler(AddressOf txtFiXXObroty_LostFocus)
                    AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXXObroty_KeyDown)

                    'stworz zmienn¹ BUTTON
                    Dim CtrlB As New System.Windows.Forms.Button
                    CtrlB.Name = "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                    CtrlB.Image = cmdFi.Image
                    CtrlB.ImageAlign = ContentAlignment.MiddleCenter
                    CtrlB.Visible = False
                    Me.Obroty.Controls.Add(CtrlB)
                    AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXXObroty_Click)
                    AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXXObroty_GotFocus)
                    AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXXObroty_LostFocus)
                Next
                Me.Refresh()
                '--------------------------------------------------
                StartFormy = False
                StartFormyObroty = False
                dagObroty_CurrentCellChanged(sender, e)
                InicjujFiltryColumnDGV(Me.Obroty, dagObroty, DataViewObroty.Table.Columns.Count, stbFiltryObroty, StartFormyObroty)
                '--------------------------------------------------------------------
                cmdPowrot.Focus()





            Case 7  'akcje marketingowe
                lblIdent_Akcje.Text = txtIdent.Text
                lblTofi_Akcje.Text = txtTofi.Text
                lblNazwaHandlowa_Akcje.Text = txtNazwa1.Text
                lblNazwa1_Akcje.Text = Trim(txtKodP.Text) & " " & Trim(txtMiejscowosc.Text) & ", " & Trim(txtUlica.Text) & " " & Trim(txtNrDomu.Text) & " " & Trim(txtIDDomu.Text)
                lblNazwa2_Akcje.Text = OsobaKontaktowa(txtIdent.Text)
                lblStanowisko_Akcje.Text = oStanowiskoOso
                lblTelefon_Akcje.Text = oTelefonOso
                lblTelefonKom_Akcje.Text = oTelefonKomOso

                cmdKontaktPromotor.Visible = False
                cmdCoDalej.Visible = False
                '---------------------------------------------
                'AkcjeSpec
                StartFormyAnalizyZakupu = True
                dbSelectSQLAkcjeSpec = "SELECT AkcjeSpec.IDENTKLI, " &
                                                           "AkcjeSpec.IDENTAKCJI, " &
                                                           "AkcjeOpis.DATA, " &
                                                           "AkcjeOpis.OPIS " &
                                                           "FROM AkcjeSpec INNER JOIN AkcjeOpis " &
                                                           "ON AkcjeSpec.IDENTAKCJI=AkcjeOpis.IDENT " &
                                                           "WHERE AkcjeSpec.IDENTKLI='" & oIdentKli & "'"
                DataSetAkcjeSpec.Locale = New System.Globalization.CultureInfo("pl-PL")
                DataSetAkcjeSpec = New DataSet
                DataViewAkcjeSpec = New DataView

                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbConnectionAkcjeSpec = New OleDb.OleDbConnection(sConnectionAkcjeSpec)
                    'dbCommandSelectAkcjeSpec = New OleDb.OleDbCommand(dbSelectSQLAkcjeSpec, dbConnectionAkcjeSpec)
                    ''dbCommandDeleteAkcjeSpec = New OleDb.OleDbCommand(sDeleteSQLAkcjeSpec, dbConnectionAkcjeSpec)
                    ''dbCommandUpdateAkcjeSpec = New OleDb.OleDbCommand(sUpdateSQLAkcjeSpec, dbConnectionAkcjeSpec)
                    ''dbCommandInsertAkcjeSpec = New OleDb.OleDbCommand(sInsertSQLAkcjeSpec, dbConnectionAkcjeSpec)
                    'dbDataAdapterAkcjeSpec = New OleDb.OleDbDataAdapter(dbCommandSelectAkcjeSpec)

                    ''mapujemy domyslna nazwe tabeli
                    'Dim MapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
                    'MapowanieTabeliAkcjeSpec = dbDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_AkcjeSpec")

                    ''komendy DataAdapter
                    ''Dim objParamAkcjeSpec As OleDb.OleDbParameter
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
                    'Finally
                    '    ConnectionState = dbConnectionAkcjeSpec.State
                    'End Try
                Else
                    sqlConnectionAkcjeSpec = New SqlClient.SqlConnection(sConnectionAkcjeSpec)
                    sqlCommandSelectAkcjeSpec = New SqlClient.SqlCommand(dbSelectSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                    'sqlCommandDeleteAkcjeSpec = New SqlClient.SqlCommand(sDeleteSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                    'sqlCommandUpdateAkcjeSpec = New SqlClient.SqlCommand(sUpdateSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                    'sqlCommandInsertAkcjeSpec = New SqlClient.SqlCommand(sInsertSQLAkcjeSpec, sqlConnectionAkcjeSpec)
                    sqlDataAdapterAkcjeSpec = New SqlClient.SqlDataAdapter(sqlCommandSelectAkcjeSpec)

                    'mapujemy domyslna nazwe tabeli
                    Dim MapowanieTabeliAkcjeSpec As System.Data.Common.DataTableMapping
                    MapowanieTabeliAkcjeSpec = sqlDataAdapterAkcjeSpec.TableMappings.Add("Table", "TABELA_AkcjeSpec")

                    'komendy DataAdapter
                    'Dim objParamAkcjeSpec As SqlClient.SqlParameter
                    '------------------------------------------
                    'wypelnij DATASET
                    Try
                        sqlConnectionAkcjeSpec.Open()
                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    Finally
                        ConnectionState = sqlConnectionAkcjeSpec.State
                    End Try
                End If

                If ConnectionState = ConnectionState.Open Then
                    Try
                        If ParL_DataBaseType = "MS ACCESS" Then
                            'dbDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_AkcjeSpec")
                            'dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                            'dbConnectionAkcjeSpec.Close()
                        Else
                            sqlDataAdapterAkcjeSpec.FillSchema(DataSetAkcjeSpec, SchemaType.Mapped, "TABELA_AkcjeSpec")
                            sqlDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                            sqlConnectionAkcjeSpec.Close()
                        End If

                        'zdefiniuj unikalny klucz wyszukiwania w tabeli
                        DataSetAkcjeSpec.Tables(0).PrimaryKey = New DataColumn() {DataSetAkcjeSpec.Tables(0).Columns("IDENTAKCJI"), DataSetAkcjeSpec.Tables(0).Columns("IDENTKLI")}

                        'definiuj DataView
                        DataViewAkcjeSpec = New DataView(DataSetAkcjeSpec.Tables(0))
                        DataViewAkcjeSpec.AllowDelete = False
                        DataViewAkcjeSpec.AllowNew = False

                        '-------------------------------------------------
                        'parametry DataGridView
                        '-------------------------------------------------
                        ' Initialize basic DataGridView properties.
                        dagAkcjeSpec.BackgroundColor = System.Drawing.SystemColors.Control
                        dagAkcjeSpec.GridColor = System.Drawing.SystemColors.ControlDark
                        dagAkcjeSpec.ForeColor = System.Drawing.Color.Purple
                        dagAkcjeSpec.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagAkcjeSpec.BorderStyle = BorderStyle.Fixed3D
                        'dagAkcjeSpec.Dock = DockStyle.Fill
                        dagAkcjeSpec.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                        Or System.Windows.Forms.AnchorStyles.Bottom) _
                                        Or System.Windows.Forms.AnchorStyles.Left) _
                                        Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                        ' Set property values appropriate for read-only display and limited interactivity. 
                        dagAkcjeSpec.AllowUserToAddRows = False
                        dagAkcjeSpec.AllowUserToDeleteRows = False
                        dagAkcjeSpec.AllowUserToOrderColumns = True
                        dagAkcjeSpec.AllowUserToResizeColumns = True
                        dagAkcjeSpec.AllowUserToResizeRows = True
                        dagAkcjeSpec.ReadOnly = True
                        dagAkcjeSpec.SelectionMode = DataGridViewSelectionMode.CellSelect
                        dagAkcjeSpec.MultiSelect = False

                        'UWAGA: Ustawienie AutoSizeMode naAllCells powoduje BARDZO WOLNE WYPELNIANIE DataGridView....
                        'dagAkcjeSpec.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                        dagAkcjeSpec.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
                        dagAkcjeSpec.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                        dagAkcjeSpec.ColumnHeadersVisible = True
                        dagAkcjeSpec.ColumnHeadersHeight = 40
                        dagAkcjeSpec.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
                        dagAkcjeSpec.RowHeadersVisible = True
                        dagAkcjeSpec.RowHeadersWidth = 50
                        dagAkcjeSpec.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

                        ' Set the selection background color for all the cells.
                        dagAkcjeSpec.DefaultCellStyle.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
                        dagAkcjeSpec.DefaultCellStyle.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow
                        dagAkcjeSpec.DefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagAkcjeSpec.DefaultCellStyle.ForeColor = System.Drawing.Color.Purple
                        dagAkcjeSpec.DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
                        dagAkcjeSpec.DefaultCellStyle.NullValue = ""
                        dagAkcjeSpec.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                        dagAkcjeSpec.DefaultCellStyle.WrapMode = DataGridViewTriState.True


                        ' Set RowHeadersDefaultCellStyle.SelectionBackColor so that its default 
                        ' value won't override DataGridView.DefaultCellStyle.SelectionBackColor.
                        dagAkcjeSpec.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Empty
                        dagAkcjeSpec.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.Empty

                        ' Set the background color for all rows and for alternating rows.  
                        ' The value for alternating rows overrides the value for all rows. 
                        dagAkcjeSpec.RowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.ControlLight
                        dagAkcjeSpec.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.SystemColors.Control

                        ' Set the row and column header styles.
                        dagAkcjeSpec.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagAkcjeSpec.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        dagAkcjeSpec.RowHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.Navy
                        dagAkcjeSpec.RowHeadersDefaultCellStyle.BackColor = System.Drawing.SystemColors.ScrollBar
                        '-----------------------------------
                        'wypelnienie DataGridView
                        dagAkcjeSpec.AutoGenerateColumns = False       'NIE GENERUJ AUTOMATYCZNIE KOLUMN !!!
                        'dagAkcjeSpec.DataMember = DataSetAkcjeSpec.Tables(0).TableName
                        'dagAkcjeSpec.DataSource = DataViewAkcjeSpec
                        '-----------------------------------

                        dagAkcjeSpec.Columns.Clear()

                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(0).ColumnName, "", 0, DataGridViewContentAlignment.MiddleLeft, False)
                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(1).ColumnName, "Akcja", 150, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(2).ColumnName, "Data", 80, DataGridViewContentAlignment.MiddleLeft, True)
                        TxtColumnView(dagAkcjeSpec, DataSetAkcjeSpec.Tables(0).Columns(3).ColumnName, "Opis", 300, DataGridViewContentAlignment.MiddleLeft, True)


                        Me.Cursor = Cursors.WaitCursor
                        dagAkcjeSpec.DataSource = DataViewAkcjeSpec
                        Me.Cursor = Cursors.Default

                        'ustaw sie na pierwszym zapisie
                        If DataSetAkcjeSpec.Tables(0).Rows.Count > 0 Then
                            dagAkcjeSpec.CurrentCell = dagAkcjeSpec(1, 0)
                            dagAkcjeSpec.CurrentCell.Selected = True
                        End If




                    Catch Ex As System.Exception
                        ViewInLog(Ex, Me.Name(), Nothing)
                        'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                        '    System.Windows.Forms.MessageBoxButtons.OK, _
                        '    MessageBoxIcon.Information, _
                        '    MessageBoxDefaultButton.Button1)
                    End Try
                    '---------------------------------
                    '---------------------------------
                    dagAkcjeSpec.Focus()
                    '---------------------------------
                    'dodaj StatusBar na koncu
                    stbAkcjeSpec.Panels.Clear()

                    stbAkcjeSpec.BackColor = System.Drawing.SystemColors.Control
                    stbAkcjeSpec.ForeColor = System.Drawing.Color.DarkGreen
                    stbAkcjeSpec.Anchor = CType((((System.Windows.Forms.AnchorStyles.Bottom) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                    stbAkcjeSpec.Panels.Add("stbPanelIleRec")
                    stbAkcjeSpec.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbAkcjeSpec.Panels(0).Width = 80
                    stbAkcjeSpec.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbAkcjeSpec.Panels(0).Text = "Iloœæ rec.: " & DataSetAkcjeSpec.Tables(0).Rows.Count.ToString

                    stbAkcjeSpec.Panels.Add("stbPanelWiersz")
                    stbAkcjeSpec.Panels(1).AutoSize = StatusBarPanelAutoSize.None
                    stbAkcjeSpec.Panels(1).Width = 80
                    stbAkcjeSpec.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagAkcjeSpec.CurrentCell Is Nothing Then
                        stbAkcjeSpec.Panels(1).Text = "Wi: "
                    Else
                        stbAkcjeSpec.Panels(1).Text = "Wi: " & dagAkcjeSpec.CurrentCell.RowIndex.ToString
                    End If

                    stbAkcjeSpec.Panels.Add("stbPanelKolumna")
                    stbAkcjeSpec.Panels(2).AutoSize = StatusBarPanelAutoSize.None
                    stbAkcjeSpec.Panels(2).Width = 80
                    stbAkcjeSpec.Panels(2).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    If dagAkcjeSpec.CurrentCell Is Nothing Then
                        stbAkcjeSpec.Panels(2).Text = "Ko: "
                    Else
                        stbAkcjeSpec.Panels(2).Text = "Ko: " & dagAkcjeSpec.CurrentCell.ColumnIndex.ToString
                    End If

                    stbAkcjeSpec.Panels.Add("stbPanelSort")
                    stbAkcjeSpec.Panels(3).AutoSize = StatusBarPanelAutoSize.None
                    stbAkcjeSpec.Panels(3).Width = 150
                    stbAkcjeSpec.Panels(3).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbAkcjeSpec.Panels(3).Text = "Sort : "

                    stbAkcjeSpec.Panels.Add("stbPanelFiltr")
                    stbAkcjeSpec.Panels(4).AutoSize = StatusBarPanelAutoSize.None
                    stbAkcjeSpec.Panels(4).Width = 150
                    stbAkcjeSpec.Panels(4).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbAkcjeSpec.Panels(4).Text = "Filtr : "

                    stbAkcjeSpec.Panels.Add("stbPanelReszta")
                    stbAkcjeSpec.Panels(5).AutoSize = StatusBarPanelAutoSize.Spring
                    'stbAkcjeSpec.Panels(5).Width = 150
                    stbAkcjeSpec.Panels(5).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbAkcjeSpec.Panels(5).Text = ""

                    stbAkcjeSpec.ShowPanels = True
                    '---------------------------------
                    'dodaj StatusBar na koncu
                    stbFiltryAkcjeSpec.Panels.Clear()

                    stbFiltryAkcjeSpec.BackColor = System.Drawing.SystemColors.Control
                    stbFiltryAkcjeSpec.ForeColor = System.Drawing.Color.DarkGreen
                    stbFiltryAkcjeSpec.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top) _
                                    Or System.Windows.Forms.AnchorStyles.Left) _
                                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)

                    stbFiltryAkcjeSpec.Panels.Add("stbFiltryRowHead")
                    stbFiltryAkcjeSpec.Panels(0).AutoSize = StatusBarPanelAutoSize.None
                    stbFiltryAkcjeSpec.Panels(0).Width = 50  'dagAkcjeSpec.TableStyles(0).RowHeaderWidth
                    stbFiltryAkcjeSpec.Panels(0).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbFiltryAkcjeSpec.Panels(0).Text = "Filtry"

                    stbFiltryAkcjeSpec.Panels.Add("stbFiltryReszta")
                    stbFiltryAkcjeSpec.Panels(1).AutoSize = StatusBarPanelAutoSize.Spring
                    'stbFiltryAkcjeSpec.Panels(1).Width = 150
                    stbFiltryAkcjeSpec.Panels(1).BorderStyle = StatusBarPanelBorderStyle.Sunken
                    stbFiltryAkcjeSpec.Panels(1).Text = ""

                    stbFiltryAkcjeSpec.ShowPanels = True

                    'dodaj ctrl do pobierania szablonu
                    cmdWszystkoAkcjeSpec.Size = New Size(20, 22)
                    cmdWszystkoAkcjeSpec.Location = New System.Drawing.Point(stbFiltryAkcjeSpec.Location.X + 50 - 20,
                                             stbFiltryAkcjeSpec.Location.Y)
                    cmdWszystkoAkcjeSpec.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or
                                        System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
                    cmdWszystkoAkcjeSpec.Visible = True
                End If
                InitAkcjeSpec() 'inicjuj zmienne

                '--------------------------------------------------------------------
                'set the header width to something apporpriate
                dagAkcjeSpec.RowHeadersWidth = 50       'use if tablestyle
                'Me.dagAkcjeSpec.RowHeaderWidth = 80 'use if no tablestyle
                '--------------------------------------------------------------------
                ''set topleft corner point so we can locate the toprow
                'If DataSetAkcjeSpec.Tables(0).Rows.Count > 0 Then
                '    StartPointInCell00 = New System.Drawing.Point((dagAkcjeSpec.GetCellDisplayRectangle(0, 0, True).X + 4),
                '                      (dagAkcjeSpec.GetCellDisplayRectangle(0, 0, True).Y + 4))
                '    ScrollXOffset = 10000 * dagAkcjeSpec.FirstDisplayedScrollingColumnIndex +
                '        dagAkcjeSpec.GetCellDisplayRectangle(dagAkcjeSpec.FirstDisplayedScrollingColumnIndex, 0, True).Left -
                '        StartPointInCell00.X
                'Else
                '    'brak zapisow - poczatek DataGrid
                '    StartPointInCell00 = New System.Drawing.Point((dagAkcjeSpec.Left + 4), (dagAkcjeSpec.Top + 4))
                '    ScrollXOffset = 0
                'End If
                ''--------------------------------------------------
                ''inicjowanie zegara dla scrollingu poziomego
                'HorizScrollLastTime = ""
                'HorizScrollTimer.Interval = 2 * 1000
                'HorizScrollTimer.Enabled = True
                '--------------------------------------------------
                'Stworz zmienne do filtrowania
                Dim ii As Integer = 0
                For ii = 0 To DataViewAkcjeSpec.Table.Columns.Count - 1
                    'stworz zmienn¹ TXTBOX
                    Dim CtrlT As New System.Windows.Forms.TextBox
                    CtrlT.Name = "txtFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                    CtrlT.Visible = False
                    Me.AkcjeSpec.Controls.Add(CtrlT)
                    AddHandler CtrlT.TextChanged, New EventHandler(AddressOf txtFiXXAkcjeSpec_TextChanged)
                    AddHandler CtrlT.GotFocus, New EventHandler(AddressOf txtFiXXAkcjeSpec_GotFocus)
                    AddHandler CtrlT.LostFocus, New EventHandler(AddressOf txtFiXXAkcjeSpec_LostFocus)
                    AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXXAkcjeSpec_KeyDown)

                    'stworz zmienn¹ BUTTON
                    Dim CtrlB As New System.Windows.Forms.Button
                    CtrlB.Name = "cmdFi" & Microsoft.VisualBasic.Right("000" & Trim(Str(ii)), 3)
                    CtrlB.Image = cmdFi.Image
                    CtrlB.ImageAlign = ContentAlignment.MiddleCenter
                    CtrlB.Visible = False
                    Me.AkcjeSpec.Controls.Add(CtrlB)
                    AddHandler CtrlB.Click, New EventHandler(AddressOf cmdFiXXAkcjeSpec_Click)
                    AddHandler CtrlB.GotFocus, New EventHandler(AddressOf cmdFiXXAkcjeSpec_GotFocus)
                    AddHandler CtrlB.LostFocus, New EventHandler(AddressOf cmdFiXXAkcjeSpec_LostFocus)
                Next
                Me.Refresh()
                '--------------------------------------------------
                StartFormy = False
                StartFormyAkcjeSpec = False
                dagAkcjeSpec_CurrentCellChanged(sender, e)
                InicjujFiltryColumnDGV(Me.AkcjeSpec, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltryAkcjeSpec, StartFormyAkcjeSpec)
                '--------------------------------------------------------------------
                cmdPowrot.Focus()




        End Select

    End Sub












    '***************************************************
    '* procedury pobierania/zapisywania danych
    '***************************************************

    Private Sub PobierzDane()
        txtIdent.Text = oIdentKli
        txtTofi.Text = oNrTOFITxtKli

        chkKartyPayback.Checked = oKartaPaybackKli  'IIf(oKartaPaybackKli, System.Windows.Forms.CheckState.Checked, System.Windows.Forms.CheckState.Unchecked)
        If chkKartyPayback.Checked Then
            lblNryKartPayback.Visible = True
            txtNryKartPayBack.Visible = True
            cmdEdytujNumeryKartPayBack.Visible = True
        Else
            lblNryKartPayback.Visible = False
            txtNryKartPayBack.Visible = False
            cmdEdytujNumeryKartPayBack.Visible = False
        End If


        Dim txt As String = Trim(oNryKartPaybackKli)
        Dim NrKarty As String = ""
        Dim poz As Integer = 0
        poz = InStr(txt, ";")
        txtNryKartPayBack.Text = ""
        Do While poz > 0
            NrKarty = Mid(txt, 1, poz - 1)
            txt = Mid(txt, poz + 1)
            If Len(Trim(NrKarty)) > 0 Then
                txtNryKartPayBack.Text = NrKarty
            End If
            poz = InStr(txt, ";")
        Loop
        'ostatni nr karty
        NrKarty = txt
        txt = ""
        If Len(Trim(NrKarty)) > 0 Then
            txtNryKartPayBack.Text = NrKarty
        End If

        txtNIP.Text = oNIPKli
        txtNazwa1.Text = oNazwa1Kli
        txtKodP.Text = oKodPoczKli
        txtMiejscowosc.Text = oMiejscKli
        txtUlica.Text = oUlicaKli
        txtNrDomu.Text = oNumNrDomuKli.ToString("F0")
        txtIDDomu.Text = oIDDomuKli
        txtRejon.Text = oRejonKli
        txtTelefon.Text = oTelefonKli
        txtFax.Text = oFaxKli
        txteMail.Text = oeMailKli
        txtKtoWpisal.Text = oWpisalKli
        '------------------------------------
        txtPKontakt.Text = oPKontaktKli

        chbRodzL.CheckState = IIf(oRodzSprzetuLKli, System.Windows.Forms.CheckState.Checked, System.Windows.Forms.CheckState.Unchecked)
        chbRodzA.CheckState = IIf(oRodzSprzetuAKli, System.Windows.Forms.CheckState.Checked, System.Windows.Forms.CheckState.Unchecked)
        chbRodzT.CheckState = IIf(oRodzSprzetuTKli, System.Windows.Forms.CheckState.Checked, System.Windows.Forms.CheckState.Unchecked)

        txtWykazUrzadzen.Text = oWykazUrzadzenKli

        txtUwagiCharakterystyka.Text = oUwagikli
        txtTZlozenia.Text = oOfertaTZlozeniaKli
        lblOfertaPlik.Text = oOfertaPlikKli
        '------------------------------------
        txtOpisPotencjalu.Text = oSkupUwagikli

        txtPlanDlugookresowy.Text = oPlanDlugookresowyKli.ToString("F0")
        txtPlanKrotkookresowy.Text = oPlanKrotkookresowyKli.ToString("F0")

        txtLiczbaUrzadzen.Text = oLiczbaUrzadzenKli.ToString("F0")
        txtLiczbaWydrukow.Text = oLiczbaWydrukowKli.ToString("F0")
        chkZainstalowanyMonitor.Checked = oZainstalowanoMonitorKli

        chkAktZakupMatEksp.Checked = oAktZakupMatEkspKli
        chkAktRozliczaStrony.Checked = oAktRozliczaStronyDrukuKli
        chkAktKorzystaZNajmu.Checked = oAktKorzystaZNajmuKli
        chkZaintRozliczaStrony.Checked = oZaintRozliczaniemStronDrukuKli
        chkZaintKorzystaZNajmu.Checked = oZaintKorzystaniemZNajmuKli

        cbxSposobWyboruDostawcy.Text = oSposobWyboruDostawcyKli
        lblPotencjalDruku.Text = oPotencjalDrukuKli
        WeryfikujPotencjalDruku()

        dtpDataWeryfikKrytSegmentacji.Value = IIf(Len(Trim(oDataWeryfSegmentacjiKli)) > 0, oDataWeryfSegmentacjiKli, "2000-01-01")
        lblOstTransak.Text = OstTransakcje(oIdentKli, True)
        '------------------------------------
        txtSprzedazOpiekun.Text = Trim(oSprzedOpiekunkli)
        cbxSprzedazOKontR.Text = oSprzedOKontaktRKli
        txtSprzedazOKonT.Text = oSprzedOKontaktTKli
        cbxSprzedazOKonD.Text = oSprzedOKontaktDKli
        cbxSprzedazNKontR.Text = oSprzedNKontaktRKli
        txtSprzedazNKonT.Text = oSprzedNKontaktTKli
        cbxSprzedazNKonD.Text = oSprzedNKontaktDKli
        txtSprzedazUwagi.Text = oSprzedUwagiKli

        txtzakupyPopRok.Text = oZakupPopRokKli.ToString("N2")
        txtWartGran.Text = Par_WartGranMatWartosc.ToString("N2")

        'chkU1.Checked = True
        'chkU2.Checked = True
        'chkU3.Checked = True
        'chkU4.Checked = True
        'chkU5.Checked = True
        'chkU6.Checked = True
        'chkU7.Checked = True
        'chkU8.Checked = True
        'chkU9.Checked = True

        chkU1.Checked = (Mid(oUslugiDodKli, 1, 1) = "T")
        chkU2.Checked = (Mid(oUslugiDodKli, 2, 1) = "T")
        chkU3.Checked = (Mid(oUslugiDodKli, 3, 1) = "T")
        chkU4.Checked = (Mid(oUslugiDodKli, 4, 1) = "T")
        chkU5.Checked = (Mid(oUslugiDodKli, 5, 1) = "T")
        chkU6.Checked = (Mid(oUslugiDodKli, 6, 1) = "T")
        chkU7.Checked = (Mid(oUslugiDodKli, 7, 1) = "T")
        chkU8.Checked = (Mid(oUslugiDodKli, 8, 1) = "T")
        chkU9.Checked = (Mid(oUslugiDodKli, 9, 1) = "T")

        dtpU1.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 11, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 11, 10))
        dtpU2.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 21, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 21, 10))
        dtpU3.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 31, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 31, 10))
        dtpU4.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 41, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 41, 10))
        dtpU5.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 51, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 51, 10))
        dtpU6.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 61, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 61, 10))
        dtpU7.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 71, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 71, 10))
        dtpU8.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 81, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 81, 10))
        dtpU9.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 91, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 91, 10))

        dtpDataOb.Value = IIf(Len(Trim(Mid(oUslugiDodKli, 101, 10))) = 0, Mid(SysData(), 1, 4) & "-12-31", Mid(oUslugiDodKli, 101, 10))

        'If oZakupPopRokKli <= 1000 Then
        '    chkU1.Enabled = False
        '    chkU2.Enabled = False
        '    chkU3.Enabled = False
        '    chkU4.Enabled = False
        '    chkU5.Enabled = False
        '    chkU6.Enabled = False
        '    chkU7.Enabled = False
        '    chkU8.Enabled = False
        '    chkU9.Enabled = False

        '    dtpU1.Enabled = False
        '    dtpU2.Enabled = False
        '    dtpU3.Enabled = False
        '    dtpU4.Enabled = False
        '    dtpU5.Enabled = False
        '    dtpU6.Enabled = False
        '    dtpU7.Enabled = False
        '    dtpU8.Enabled = False
        '    dtpU9.Enabled = False

        'ElseIf oZakupPopRokKli <= 3000 Then
        '    chkU1.Checked = (Mid(oUslugiDodKli, 1, 1) = "T")
        '    chkU2.Checked = (Mid(oUslugiDodKli, 2, 1) = "T")
        '    chkU3.Checked = (Mid(oUslugiDodKli, 3, 1) = "T")
        '    chkU4.Checked = (Mid(oUslugiDodKli, 4, 1) = "T")
        '    'chkU5.Checked = (Mid(oUslugiDodKli, 5, 1) = "T")
        '    'chkU6.Checked = (Mid(oUslugiDodKli, 6, 1) = "T")
        '    'chkU7.Checked = (Mid(oUslugiDodKli, 7, 1) = "T")
        '    'chkU8.Checked = (Mid(oUslugiDodKli, 8, 1) = "T")
        '    'chkU9.Checked = (Mid(oUslugiDodKli, 9, 1) = "T")

        '    chkU1.Enabled = True
        '    chkU2.Enabled = True
        '    chkU3.Enabled = True
        '    chkU4.Enabled = True
        '    chkU5.Enabled = False
        '    chkU6.Enabled = False
        '    chkU7.Enabled = False
        '    chkU8.Enabled = False
        '    chkU9.Enabled = False

        '    dtpU1.Enabled = True
        '    dtpU2.Enabled = True
        '    dtpU3.Enabled = True
        '    dtpU4.Enabled = True
        '    dtpU5.Enabled = False
        '    dtpU6.Enabled = False
        '    dtpU7.Enabled = False
        '    dtpU8.Enabled = False
        '    dtpU9.Enabled = False


        'ElseIf oZakupPopRokKli <= 6000 Then
        '    chkU1.Checked = (Mid(oUslugiDodKli, 1, 1) = "T")
        '    chkU2.Checked = (Mid(oUslugiDodKli, 2, 1) = "T")
        '    chkU3.Checked = (Mid(oUslugiDodKli, 3, 1) = "T")
        '    chkU4.Checked = (Mid(oUslugiDodKli, 4, 1) = "T")
        '    chkU5.Checked = (Mid(oUslugiDodKli, 5, 1) = "T")
        '    chkU6.Checked = (Mid(oUslugiDodKli, 6, 1) = "T")
        '    'chkU7.Checked = (Mid(oUslugiDodKli, 7, 1) = "T")
        '    'chkU8.Checked = (Mid(oUslugiDodKli, 8, 1) = "T")
        '    'chkU9.Checked = (Mid(oUslugiDodKli, 9, 1) = "T")

        '    chkU1.Enabled = True
        '    chkU2.Enabled = True
        '    chkU3.Enabled = True
        '    chkU4.Enabled = True
        '    chkU5.Enabled = True
        '    chkU6.Enabled = True
        '    chkU7.Enabled = False
        '    chkU8.Enabled = False
        '    chkU9.Enabled = False

        '    dtpU1.Enabled = True
        '    dtpU2.Enabled = True
        '    dtpU3.Enabled = True
        '    dtpU4.Enabled = True
        '    dtpU5.Enabled = True
        '    dtpU6.Enabled = True
        '    dtpU7.Enabled = False
        '    dtpU8.Enabled = False
        '    dtpU9.Enabled = False

        'Else
        '    chkU1.Checked = (Mid(oUslugiDodKli, 1, 1) = "T")
        '    chkU2.Checked = (Mid(oUslugiDodKli, 2, 1) = "T")
        '    chkU3.Checked = (Mid(oUslugiDodKli, 3, 1) = "T")
        '    chkU4.Checked = (Mid(oUslugiDodKli, 4, 1) = "T")
        '    chkU5.Checked = (Mid(oUslugiDodKli, 5, 1) = "T")
        '    chkU6.Checked = (Mid(oUslugiDodKli, 6, 1) = "T")
        '    chkU7.Checked = (Mid(oUslugiDodKli, 7, 1) = "T")
        '    chkU8.Checked = (Mid(oUslugiDodKli, 8, 1) = "T")
        '    chkU9.Checked = (Mid(oUslugiDodKli, 9, 1) = "T")

        '    chkU1.Enabled = True
        '    chkU2.Enabled = True
        '    chkU3.Enabled = True
        '    chkU4.Enabled = True
        '    chkU5.Enabled = True
        '    chkU6.Enabled = True
        '    chkU7.Enabled = True
        '    chkU8.Enabled = True
        '    chkU9.Enabled = True

        '    dtpU1.Enabled = True
        '    dtpU2.Enabled = True
        '    dtpU3.Enabled = True
        '    dtpU4.Enabled = True
        '    dtpU5.Enabled = True
        '    dtpU6.Enabled = True
        '    dtpU7.Enabled = True
        '    dtpU8.Enabled = True
        '    dtpU9.Enabled = True
        'End If

        txtRaport01.Text = oRZURap01
        txtRaport02.Text = oRZURap02
        txtRaport03.Text = oRZURap03
        txtRaport04.Text = oRZURap04
        txtRaport05.Text = oRZURap05
        txtRaport06.Text = oRZURap06
        txtRaport07.Text = oRZURap07
        txtRaport08.Text = oRZURap08
        txtRaport09.Text = oRZURap09




        'Public oBranzaKli As String        'BRANZA         Text, 100
        'Public oPodBranzaKli As String     'PODBRANZA      Text, 100
        'Public oLiczbaZaktrudnionychKli As Integer   'LICZBAZATRUDNIONYCH       INTEGER
        'Public oLiczbaPracZdalnieKli As Integer      'LICZBAPRACZDALNIE         INTEGER
        txtBranza.Text = oBranzaKli
        txtPodbranza.Text = oPodBranzaKli
        txtLiczbaZatrudnionych.Text = oLiczbaZaktrudnionychKli.ToString("N0")
        txtLiczbaPracZdalnie.Text = oLiczbaPracZdalnieKli.ToString("N0")








    End Sub


    Private Sub ZapiszDane()
        oIdentKli = txtIdent.Text
        oNrTOFITxtKli = txtTofi.Text

        oKartaPaybackKli = chkKartyPayback.Checked
        'oNryKartPaybackKli = ""

        oNIPKli = txtNIP.Text
        oNazwa1Kli = txtNazwa1.Text
        oKodPoczKli = txtKodP.Text
        oMiejscKli = txtMiejscowosc.Text
        oUlicaKli = txtUlica.Text
        oNrDomuKli = txtNrDomu.Text
        oNumNrDomuKli = CInt(txtNrDomu.Text)
        oIDDomuKli = txtIDDomu.Text
        oRejonKli = txtRejon.Text
        oTelefonKli = txtTelefon.Text
        oFaxKli = txtFax.Text
        oeMailKli = txteMail.Text
        oWpisalKli = txtKtoWpisal.Text
        '------------------------------------
        oPKontaktKli = txtPKontakt.Text

        oRodzSprzetuLKli = chbRodzL.Checked
        oRodzSprzetuAKli = chbRodzA.Checked
        oRodzSprzetuTKli = chbRodzT.Checked

        oWykazUrzadzenKli = txtWykazUrzadzen.Text

        oUwagikli = txtUwagiCharakterystyka.Text
        oOfertaTZlozeniaKli = txtTZlozenia.Text
        oOfertaPlikKli = lblOfertaPlik.Text
        '------------------------------------
        oSkupUwagikli = txtOpisPotencjalu.Text
        oPlanDlugookresowyKli = GetIntFromText(txtPlanDlugookresowy.Text)
        oPlanKrotkookresowyKli = GetIntFromText(txtPlanKrotkookresowy.Text)

        oLiczbaUrzadzenKli = GetIntFromText(txtLiczbaUrzadzen.Text)
        oLiczbaWydrukowKli = GetIntFromText(txtLiczbaWydrukow.Text)
        oSposobWyboruDostawcyKli = cbxSposobWyboruDostawcy.Text
        oZainstalowanoMonitorKli = chkZainstalowanyMonitor.Checked
        oPotencjalDrukuKli = lblPotencjalDruku.Text

        oAktZakupMatEkspKli = chkAktZakupMatEksp.Checked
        oAktRozliczaStronyDrukuKli = chkAktRozliczaStrony.Checked
        oAktKorzystaZNajmuKli = chkAktKorzystaZNajmu.Checked
        oZaintRozliczaniemStronDrukuKli = chkZaintRozliczaStrony.Checked
        oZaintKorzystaniemZNajmuKli = chkZaintKorzystaZNajmu.Checked

        oDataWeryfSegmentacjiKli = IIf(dtpDataWeryfikKrytSegmentacji.Value.ToString("yyyy-MM-dd") = "2000-01-01", "", dtpDataWeryfikKrytSegmentacji.Value.ToString("yyyy-MM-dd"))
        '------------------------------------
        oSprzedOpiekunkli = Trim(txtSprzedazOpiekun.Text)
        oSprzedOKontaktRKli = cbxSprzedazOKontR.Text
        oSprzedOKontaktTKli = txtSprzedazOKonT.Text
        oSprzedOKontaktDKli = cbxSprzedazOKonD.Text
        oSprzedNKontaktRKli = cbxSprzedazNKontR.Text
        oSprzedNKontaktTKli = txtSprzedazNKonT.Text
        oSprzedNKontaktDKli = cbxSprzedazNKonD.Text
        oSprzedUwagiKli = txtSprzedazUwagi.Text
        '------------------------------------
        oUslugiDodKli = IIf(chkU1.Checked, "T", "N") &
                        IIf(chkU2.Checked, "T", "N") &
                        IIf(chkU3.Checked, "T", "N") &
                        IIf(chkU4.Checked, "T", "N") &
                        IIf(chkU5.Checked, "T", "N") &
                        IIf(chkU6.Checked, "T", "N") &
                        IIf(chkU7.Checked, "T", "N") &
                        IIf(chkU8.Checked, "T", "N") &
                        IIf(chkU9.Checked, "T", "N") & " " &
                        dtpU1.Value.ToString("yyyy-MM-dd") &
                        dtpU2.Value.ToString("yyyy-MM-dd") &
                        dtpU3.Value.ToString("yyyy-MM-dd") &
                        dtpU4.Value.ToString("yyyy-MM-dd") &
                        dtpU5.Value.ToString("yyyy-MM-dd") &
                        dtpU6.Value.ToString("yyyy-MM-dd") &
                        dtpU7.Value.ToString("yyyy-MM-dd") &
                        dtpU8.Value.ToString("yyyy-MM-dd") &
                        dtpU9.Value.ToString("yyyy-MM-dd") &
                        dtpDataOb.Value.ToString("yyyy-MM-dd")


        oRZURap01 = txtRaport01.Text
        oRZURap02 = txtRaport02.Text
        oRZURap03 = txtRaport03.Text
        oRZURap04 = txtRaport04.Text
        oRZURap05 = txtRaport05.Text
        oRZURap06 = txtRaport06.Text
        oRZURap07 = txtRaport07.Text
        oRZURap08 = txtRaport08.Text
        oRZURap09 = txtRaport09.Text

        'oZakupPopRokKli = GetDblFromText(txtzakupyPopRok.Text)

        'oMarkerKli = False
        'oMarkerPKli = False
        '------------------------------------

        'Public oBranzaKli As String        'BRANZA         Text, 100
        'Public oPodBranzaKli As String     'PODBRANZA      Text, 100
        'Public oLiczbaZaktrudnionychKli As Integer   'LICZBAZATRUDNIONYCH       INTEGER
        'Public oLiczbaPracZdalnieKli As Integer      'LICZBAPRACZDALNIE         INTEGER
        oBranzaKli = txtBranza.Text
        oPodBranzaKli = txtPodbranza.Text
        oLiczbaZaktrudnionychKli = GetIntFromText(txtLiczbaZatrudnionych.Text)
        oLiczbaPracZdalnieKli = GetIntFromText(txtLiczbaPracZdalnie.Text)
    End Sub






    '*******************************************
    '** procedury obslugi edycji
    '*******************************************

    Private Sub txtIdent_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIdent.TextChanged
        If TestIntLen(txtIdent, "IDENTYFIKATOR", 6) Then
        End If
        '-------------------------
        If Not StartFormy Then
            oIdentKli = txtIdent.Text
        End If
        Dim dbSelectSQLOsoby As String = "SELECT * FROM Osoby WHERE IDENTKLIENTA='" & oIdentKli & "'"
        Dim dbSelectSQLObrotyMies As String = baseSelectSQLObrotyMies & " WHERE IDENTKLIENTA='" & oIdentKli & "' ORDER BY ObrotyMies.MIESIAC DESC"
        Dim dbSelectSQLObroty As String = baseSelectSQLObroty & " WHERE IDENTKLIENTA='" & oIdentKli & "' ORDER BY Obroty.DATA DESC"
    End Sub
    Private Sub txtIdent_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIdent.GotFocus
        StartEdycjiTxt(txtIdent)
    End Sub
    Private Sub txtIdent_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIdent.LostFocus
        KoniecEdycjiTxt(txtIdent)
        '-------------------
        'If TestIntLen(txtIdent, "IDENTYFIKATOR", 6) Then
        '    txtIdent.Focus()
        'End If
    End Sub








    Private Sub txtTofi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTofi.TextChanged
        TestLen(txtTofi, "IDENT. TOFI", 500)
    End Sub

    Private Sub txtNIP_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNIP.TextChanged
        TestLen(txtNIP, "NIP", 15)
    End Sub

    Private Sub txtNazwa1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNazwa1.TextChanged
        TestLen(txtNazwa1, "NAZWA KLIENTA", 100)
    End Sub

    Private Sub txtKodP_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKodP.TextChanged
        TestLen(txtKodP, "KOD POCZTOWY", 10)
    End Sub

    Private Sub txtMiejscowosc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMiejscowosc.TextChanged
        TestLen(txtMiejscowosc, "MIEJSCOWOSC", 50)
    End Sub

    Private Sub txtUlica_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUlica.TextChanged
        TestLen(txtUlica, "ADRES #1", 50)
    End Sub

    Private Sub txtNrDomu_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNrDomu.TextChanged
        If TestNum(txtNrDomu, "NR DOMU") Then
        End If
    End Sub

    Private Sub txtIDDomu_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIDDomu.TextChanged
        TestLen(txtIDDomu, "ID DOMU", 10)
    End Sub

    Private Sub txtRejon_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TestLen(txtRejon, "REJON", 20)
    End Sub

    Private Sub txtTelefon_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTelefon.TextChanged
        TestLen(txtTelefon, "TELEFON", 50)
    End Sub

    Private Sub txtFax_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFax.TextChanged
        TestLen(txtFax, "FAX", 50)
    End Sub

    Private Sub txteMail_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txteMail.TextChanged
        TestLen(txteMail, "ADRES EMAIL", 50)
    End Sub


    Private Sub txtKtoWpisal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKtoWpisal.TextChanged
        TestLen(txtKtoWpisal, "KTO WPISA£", 10)
    End Sub

    '-----------------------------------------------

    Private Sub txtPKontakt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPKontakt.TextChanged
        TestLen(txtPKontakt, "PIERWSZY KONTAKT", 10)
    End Sub

    Private Sub txtTZlozenia_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTZlozenia.TextChanged
        TestLen(txtTZlozenia, "TERMIN ZLOZENIA", 4)
    End Sub

    Private Sub txtUwagiCharakterystyka_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUwagiCharakterystyka.TextChanged
        'TestLen(txtUwagiCharakterystyka, "UWAGI", 300)
    End Sub

    '-----------------------------------------------


    Private Sub txtOpisPotencjalu_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOpisPotencjalu.TextChanged
        'TestLen(txtOpisPotencjalu, "SKUP UWAGI", 300)
    End Sub

    '------------------------------------------------------------

    Private Sub txtSprzedazOpiekun_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSprzedazOpiekun.TextChanged
        TestLen(txtSprzedazOpiekun, "SPRZEDA¯ OPIEKUN", 10)
    End Sub

    Private Sub cbxSprzedazOKontR_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSprzedazOKontR.SelectedIndexChanged
        TestCbxLen(cbxSprzedazOKontR, "SPRZEDA¯ OSTATNI KONTAKT - ROK", 4)
    End Sub


    Private Sub txtSprzedazOKonT_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSprzedazOKonT.TextChanged
        'TestLen(txtSprzedazOKonT, "SPRZEDA¯ OSTATNI KONTAKT - TYDZIEÑ", 2)
        If Not TestNum(txtSprzedazOKonT, "SPRZEDA¯ OSTATNI KONTAKT - TYDZIEÑ") Then
            txtSprzedazOKonT.Text = MakeNum(txtSprzedazOKonT.Text)
        End If
    End Sub
    Private Sub txtSprzedazOKonT_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSprzedazOKonT.GotFocus
        StartEdycjiTxt(txtSprzedazOKonT)
    End Sub
    Private Sub txtSprzedazOKonT_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSprzedazOKonT.LostFocus
        KoniecEdycjiTxt(txtSprzedazOKonT)
        '------------------------
        If TestNum(txtSprzedazOKonT, "SPRZEDA¯ OSTATNI KONTAKT - TYDZIEÑ") Then
            Dim NRTy As Integer = GetIntFromText(txtSprzedazOKonT.Text)
            If (NRTy < 0) Or (NRTy > 52) Then
                MessageBox.Show("Numer tygodnia mo¿e byæ liczb¹ z zakresu 0..52.",
                            "Uwaga :",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1)
                txtSprzedazOKonT.Focus()
            End If
        End If
    End Sub




    Private Sub cbxSprzedazOKonD_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSprzedazOKonD.SelectedIndexChanged
        TestCbxLen(cbxSprzedazOKonD, "SPRZEDA¯ OSTATNI KONTAKT - DZIEÑ TYGODNIA", 12)
    End Sub

    Private Sub cbxSprzedazNKonR_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSprzedazNKontR.SelectedIndexChanged
        TestCbxLen(cbxSprzedazNKontR, "SPRZEDA¯ KOLEJNY KONTAKT - ROK", 4)
    End Sub

    Private Sub txtSprzedazNKonT_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSprzedazNKonT.TextChanged
        'TestLen(txtSprzedazNKonT, "SPRZEDA¯ NASTÊPNY KONTAKT - TYDZIEÑ", 2)
        If Not TestNum(txtSprzedazNKonT, "SPRZEDA¯ NASTÊPNY KONTAKT - TYDZIEÑ") Then
            txtSprzedazNKonT.Text = MakeNum(txtSprzedazNKonT.Text)
        End If
    End Sub
    Private Sub txtSprzedazNKonT_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSprzedazNKonT.GotFocus
        StartEdycjiTxt(txtSprzedazNKonT)
    End Sub
    Private Sub txtSprzedazNKonT_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSprzedazNKonT.LostFocus
        KoniecEdycjiTxt(txtSprzedazNKonT)
        '------------------------
        If TestNum(txtSprzedazNKonT, "SPRZEDA¯ NASTÊPNY KONTAKT - TYDZIEÑ") Then
            Dim NRTy As Integer = GetIntFromText(txtSprzedazNKonT.Text)
            If (NRTy < 0) Or (NRTy > 52) Then
                MessageBox.Show("Numer tygodnia mo¿e byæ liczb¹ z zakresu 0..52.",
                            "Uwaga :",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1)
                txtSprzedazNKonT.Focus()
            End If
        End If
    End Sub

    Private Sub cbxSprzedazNKonD_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSprzedazNKonD.SelectedIndexChanged
        TestCbxLen(cbxSprzedazNKonD, "SPRZEDA¯ KOLEJNY KONTAKT - DZIEÑ TYGODNIA", 12)
    End Sub


    Private Sub txtSprzedazUwagi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'TestLen(txtSprzedazUwagi, "SPRZEDA¯ UWAGI", 300)
    End Sub







    Private Sub txtPlanDlugookresowy_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPlanDlugookresowy.TextChanged
        If TestNum(txtPlanDlugookresowy, "PLAN D£UGOOKRESOWY") Then
        End If
    End Sub
    Private Sub txtPlanDlugookresowy_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPlanDlugookresowy.GotFocus
        StartEdycjiTxt(txtPlanDlugookresowy)
    End Sub
    Private Sub txtPlanDlugookresowy_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPlanDlugookresowy.LostFocus
        KoniecEdycjiTxt(txtPlanDlugookresowy)
    End Sub

    Private Sub txtPlanKrotkookresowy_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPlanKrotkookresowy.TextChanged
        If TestNum(txtPlanKrotkookresowy, "PLAN KRÓTKOOKRESOWY") Then
        End If
    End Sub
    Private Sub txtPlanKrotkookresowy_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPlanKrotkookresowy.GotFocus
        StartEdycjiTxt(txtPlanKrotkookresowy)
    End Sub
    Private Sub txtPlanKrotkookresowy_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPlanKrotkookresowy.LostFocus
        KoniecEdycjiTxt(txtPlanKrotkookresowy)
    End Sub






    'sprawdz czy zmieniono parametry segmentacji
    'txtLiczbaUrzadzen.Text = oLiczbaUrzadzenKli.ToString("F0")
    'txtLiczbaWydrukow.Text = oLiczbaWydrukowKli.ToString("F0")
    'cbxSposobWyboruDostawcy.Text = oSposobWyboruDostawcyKli
    'chkAktZakupMatEksp.Checked = oAktZakupMatEkspKli
    'chkAktRozliczaStrony.Checked = oAktRozliczaStronyDrukuKli
    'chkAktKorzystaZNajmu.Checked = oAktKorzystaZNajmuKli
    'chkZaintRozliczaStrony.Checked = oZaintRozliczaniemStronDrukuKli
    'chkZaintKorzystaZNajmu.Checked = oZaintKorzystaniemZNajmuKli
    Private Sub CzyZmienionoKryteriaSegmentacji()
        If (txtLiczbaUrzadzen.Text = oLiczbaUrzadzenKli.ToString("F0")) And
           (txtLiczbaWydrukow.Text = oLiczbaWydrukowKli.ToString("F0")) And
           (cbxSposobWyboruDostawcy.Text = oSposobWyboruDostawcyKli) And
           (chkAktZakupMatEksp.Checked = oAktZakupMatEkspKli) And
           (chkAktRozliczaStrony.Checked = oAktRozliczaStronyDrukuKli) And
           (chkAktKorzystaZNajmu.Checked = oAktKorzystaZNajmuKli) And
           (chkZaintRozliczaStrony.Checked = oZaintRozliczaniemStronDrukuKli) And
           (chkZaintKorzystaZNajmu.Checked = oZaintKorzystaniemZNajmuKli) Then
            ' nie zmieni³y sie - nic nie modyfikuj
        Else
            'ZMIENI£Y SIE - ustaw date weryfikacji 
            dtpDataWeryfikKrytSegmentacji.Value = SysData()
        End If
    End Sub





    Private Sub txtLiczbaUrzadzen_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLiczbaUrzadzen.TextChanged
        If TestNum(txtLiczbaUrzadzen, "LICZBA URZ¥DZEN") Then
        End If
    End Sub
    Private Sub txtLiczbaUrzadzen_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLiczbaUrzadzen.GotFocus
        StartEdycjiTxt(txtLiczbaUrzadzen)
    End Sub
    Private Sub txtLiczbaUrzadzen_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLiczbaUrzadzen.LostFocus
        KoniecEdycjiTxt(txtLiczbaUrzadzen)
        CzyZmienionoKryteriaSegmentacji()
        WeryfikujPotencjalDruku()
    End Sub

    Private Sub txtLiczbaWydrukow_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLiczbaWydrukow.TextChanged
        If TestNum(txtLiczbaWydrukow, "LICZBA WYDRUKÓW") Then
        End If
    End Sub
    Private Sub txtLiczbaWydrukow_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLiczbaWydrukow.GotFocus
        StartEdycjiTxt(txtLiczbaWydrukow)
    End Sub
    Private Sub txtLiczbaWydrukow_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLiczbaWydrukow.LostFocus
        KoniecEdycjiTxt(txtLiczbaWydrukow)
        CzyZmienionoKryteriaSegmentacji()
        WeryfikujPotencjalDruku()
    End Sub

    Private Sub WeryfikujPotencjalDruku()
        Dim LiUrzadzen As Integer = GetIntFromText(txtLiczbaUrzadzen.Text)
        Dim LiWydrukow As Integer = GetIntFromText(txtLiczbaWydrukow.Text)

        If (LiWydrukow = 0) And (LiUrzadzen = 0) Then
            oPotencjalDrukuKli = ""
        ElseIf LiWydrukow > 0 Then
            oPotencjalDrukuKli = IIf(LiWydrukow < 15, "W0", IIf(LiWydrukow < 50, "W1", IIf(LiWydrukow < 100, "W2", IIf(LiWydrukow < 200, "W3", "W4"))))
        Else
            oPotencjalDrukuKli = IIf(LiUrzadzen <= 2, "U1", IIf(LiUrzadzen <= 5, "U2", IIf(LiUrzadzen <= 24, "U3", "U4")))
        End If
        lblPotencjalDruku.Text = oPotencjalDrukuKli
    End Sub


    Private Sub cbxSposobWyboruDostawcy_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxSposobWyboruDostawcy.SelectedIndexChanged
    End Sub
    Private Sub cbxSposobWyboruDostawcy_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSposobWyboruDostawcy.GotFocus
        StartEdycjiCbx(cbxSposobWyboruDostawcy)
    End Sub
    Private Sub cbxSposobWyboruDostawcy_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSposobWyboruDostawcy.LostFocus
        KoniecEdycjiCbx(cbxSposobWyboruDostawcy)
        CzyZmienionoKryteriaSegmentacji()
        'WeryfikujPotencjalDruku()
    End Sub


    Private Sub chkAktZakupMatEksp_lostfocus(sender As Object, e As EventArgs) Handles chkAktZakupMatEksp.LostFocus
        CzyZmienionoKryteriaSegmentacji()
    End Sub

    Private Sub chkAktRozliczaStrony_lostfocus(sender As Object, e As EventArgs) Handles chkAktRozliczaStrony.LostFocus
        CzyZmienionoKryteriaSegmentacji()
    End Sub

    Private Sub chkAktKorzystaZNajmu_lostfocus(sender As Object, e As EventArgs) Handles chkAktKorzystaZNajmu.LostFocus
        CzyZmienionoKryteriaSegmentacji()
    End Sub

    Private Sub chkZaintRozliczaStrony_lostfocus(sender As Object, e As EventArgs) Handles chkZaintRozliczaStrony.LostFocus
        CzyZmienionoKryteriaSegmentacji()
    End Sub

    Private Sub chkZaintKorzystaZNajmu_lostfocus(sender As Object, e As EventArgs) Handles chkZaintKorzystaZNajmu.LostFocus
        CzyZmienionoKryteriaSegmentacji()
    End Sub


    '**************************************************************
    '** edycja pliku Raportu RZU
    '**************************************************************


    Private Sub txtRaport01_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport01.TextChanged
        TestLen(txtRaport01, "RAPORT RZU", 100)
    End Sub
    Private Sub txtRaport01_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport01.GotFocus
        StartEdycjiTxt(txtRaport01)
    End Sub
    Private Sub txtRaport01_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport01.LostFocus
        KoniecEdycjiTxt(txtRaport01)
    End Sub

    Private Sub txtRaport02_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport02.TextChanged
        TestLen(txtRaport02, "RAPORT RZU", 100)
    End Sub
    Private Sub txtRaport02_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport02.GotFocus
        StartEdycjiTxt(txtRaport02)
    End Sub
    Private Sub txtRaport02_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport02.LostFocus
        KoniecEdycjiTxt(txtRaport02)
    End Sub

    Private Sub txtRaport03_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport03.TextChanged
        TestLen(txtRaport03, "RAPORT RZU", 100)
    End Sub
    Private Sub txtRaport03_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport03.GotFocus
        StartEdycjiTxt(txtRaport03)
    End Sub
    Private Sub txtRaport03_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport03.LostFocus
        KoniecEdycjiTxt(txtRaport03)
    End Sub

    Private Sub txtRaport04_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport04.TextChanged
        TestLen(txtRaport04, "RAPORT RZU", 100)
    End Sub
    Private Sub txtRaport04_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport04.GotFocus
        StartEdycjiTxt(txtRaport04)
    End Sub
    Private Sub txtRaport04_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport04.LostFocus
        KoniecEdycjiTxt(txtRaport04)
    End Sub

    Private Sub txtRaport05_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport05.TextChanged
        TestLen(txtRaport05, "RAPORT RZU", 100)
    End Sub
    Private Sub txtRaport05_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport05.GotFocus
        StartEdycjiTxt(txtRaport05)
    End Sub
    Private Sub txtRaport05_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport05.LostFocus
        KoniecEdycjiTxt(txtRaport05)
    End Sub

    Private Sub txtRaport06_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport06.TextChanged
        TestLen(txtRaport06, "RAPORT RZU", 100)
    End Sub
    Private Sub txtRaport06_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport06.GotFocus
        StartEdycjiTxt(txtRaport06)
    End Sub
    Private Sub txtRaport06_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport06.LostFocus
        KoniecEdycjiTxt(txtRaport06)
    End Sub

    Private Sub txtRaport07_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport07.TextChanged
        TestLen(txtRaport07, "RAPORT RZU", 100)
    End Sub
    Private Sub txtRaport07_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport07.GotFocus
        StartEdycjiTxt(txtRaport07)
    End Sub
    Private Sub txtRaport07_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport07.LostFocus
        KoniecEdycjiTxt(txtRaport07)
    End Sub

    Private Sub txtRaport08_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport08.TextChanged
        TestLen(txtRaport08, "RAPORT RZU", 100)
    End Sub
    Private Sub txtRaport08_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport08.GotFocus
        StartEdycjiTxt(txtRaport08)
    End Sub
    Private Sub txtRaport08_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport08.LostFocus
        KoniecEdycjiTxt(txtRaport08)
    End Sub

    Private Sub txtRaport09_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport09.TextChanged
        TestLen(txtRaport09, "RAPORT RZU", 9100)
    End Sub
    Private Sub txtRaport09_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport09.GotFocus
        StartEdycjiTxt(txtRaport09)
    End Sub
    Private Sub txtRaport09_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRaport09.LostFocus
        KoniecEdycjiTxt(txtRaport09)
    End Sub



















    '**************************************************************
    '** obsluga obiektu w trakcie edycji
    '**************************************************************


    Private Sub txtTofi_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTofi.GotFocus
        StartEdycjiTxt(txtTofi)
    End Sub
    Private Sub txtTofi_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTofi.LostFocus
        KoniecEdycjiTxt(txtTofi)
    End Sub

    Private Sub txtNIP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNIP.GotFocus
        StartEdycjiTxt(txtNIP)
    End Sub
    Private Sub txtNIP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNIP.LostFocus
        KoniecEdycjiTxt(txtNIP)
    End Sub

    Private Sub txtNazwa1_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNazwa1.GotFocus
        StartEdycjiTxt(txtNazwa1)
    End Sub
    Private Sub txtNazwa1_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNazwa1.LostFocus
        KoniecEdycjiTxt(txtNazwa1)
    End Sub

    Private Sub txtKodP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKodP.GotFocus
        StartEdycjiTxt(txtKodP)
    End Sub
    Private Sub txtKodP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtKodP.LostFocus
        KoniecEdycjiTxt(txtKodP)
    End Sub

    Private Sub txtMiejscowosc_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMiejscowosc.GotFocus
        StartEdycjiTxt(txtMiejscowosc)
    End Sub
    Private Sub txtMiejscowosc_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMiejscowosc.LostFocus
        KoniecEdycjiTxt(txtMiejscowosc)
    End Sub

    Private Sub txtUlica_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUlica.GotFocus
        StartEdycjiTxt(txtUlica)
    End Sub
    Private Sub txtUlica_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUlica.LostFocus
        KoniecEdycjiTxt(txtUlica)
    End Sub

    Private Sub txtNrDomu_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNrDomu.GotFocus
        StartEdycjiTxt(txtNrDomu)
    End Sub
    Private Sub txtNrDomu_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNrDomu.LostFocus
        KoniecEdycjiTxt(txtNrDomu)
    End Sub

    Private Sub txtIDDomu_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIDDomu.GotFocus
        StartEdycjiTxt(txtIDDomu)
    End Sub
    Private Sub txtIDDomu_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIDDomu.LostFocus
        KoniecEdycjiTxt(txtIDDomu)
    End Sub

    Private Sub txtRejon_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(txtRejon)
    End Sub
    Private Sub txtRejon_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)
        KoniecEdycjiTxt(txtRejon)
    End Sub

    Private Sub txtTelefon_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTelefon.GotFocus
        StartEdycjiTxt(txtTelefon)
    End Sub
    Private Sub txtTelefon_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTelefon.LostFocus
        KoniecEdycjiTxt(txtTelefon)
    End Sub

    Private Sub txtFax_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFax.GotFocus
        StartEdycjiTxt(txtFax)
    End Sub
    Private Sub txtFax_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFax.LostFocus
        KoniecEdycjiTxt(txtFax)
    End Sub

    Private Sub txteMail_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txteMail.GotFocus
        StartEdycjiTxt(txteMail)
    End Sub
    Private Sub txteMail_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txteMail.LostFocus
        KoniecEdycjiTxt(txteMail)
        '------------------
        If String.IsNullOrEmpty(Trim(txteMail.Text)) Then
            'adres eMail niezdef - jesk OK
        Else
            'adres eMail zdef - czy jest poprawny
            If Not IsValidEmail(Trim(txteMail.Text)) Then
                MessageBox.Show("Adres email klienta nie jest poprawny" & vbCrLf &
                                Trim(txteMail.Text),
                                "Uwaga.",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation)
                txteMail.Focus()
            Else
            End If
        End If
    End Sub

    Private Sub txtKtoWpisal_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKtoWpisal.GotFocus
        StartEdycjiTxt(txtKtoWpisal)
    End Sub
    Private Sub txtKtoWpisal_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKtoWpisal.LostFocus
        KoniecEdycjiTxt(txtKtoWpisal)
    End Sub

    '---------------------------------------------------

    Private Sub txtPKontakt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPKontakt.GotFocus
        StartEdycjiTxt(txtPKontakt)
    End Sub
    Private Sub txtPKontakt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPKontakt.LostFocus
        KoniecEdycjiTxt(txtPKontakt)
    End Sub


    Private Sub chbRodzL_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbRodzL.GotFocus
        StartEdycjiChb(chbRodzL)
    End Sub
    Private Sub chbRodzL_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbRodzL.LostFocus
        KoniecEdycjiChb(chbRodzL)
    End Sub

    Private Sub chbRodzA_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbRodzA.GotFocus
        StartEdycjiChb(chbRodzA)
    End Sub
    Private Sub chbRodzA_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbRodzA.LostFocus
        KoniecEdycjiChb(chbRodzA)
    End Sub

    Private Sub chbRodzT_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbRodzT.GotFocus
        StartEdycjiChb(chbRodzT)
    End Sub
    Private Sub chbRodzT_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbRodzT.LostFocus
        KoniecEdycjiChb(chbRodzT)
    End Sub




    Private Sub txtWykazUrzadzen_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWykazUrzadzen.TextChanged
        'TestLen(txtIloscSprzetu, "ILOŒÆ SPRZÊTU", 100)
    End Sub
    Private Sub txtWykazUrzadzen_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWykazUrzadzen.GotFocus
        StartEdycjiTxt(txtWykazUrzadzen)
    End Sub
    Private Sub txtWykazUrzadzen_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWykazUrzadzen.LostFocus
        KoniecEdycjiTxt(txtWykazUrzadzen)
    End Sub

    Private Sub txtTZlozenia_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTZlozenia.GotFocus
        StartEdycjiTxt(txtTZlozenia)
    End Sub
    Private Sub txtTZlozenia_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTZlozenia.LostFocus
        KoniecEdycjiTxt(txtTZlozenia)
    End Sub

    Private Sub txtUwagiCharakterystyka_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUwagiCharakterystyka.GotFocus
        StartEdycjiTxt(txtUwagiCharakterystyka)
    End Sub
    Private Sub txtUwagiCharakterystyka_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUwagiCharakterystyka.LostFocus
        KoniecEdycjiTxt(txtUwagiCharakterystyka)
    End Sub

    '-----------------------------------------------

    Private Sub txtOpisPotencjalu_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpisPotencjalu.GotFocus
        StartEdycjiTxt(txtOpisPotencjalu)
    End Sub
    Private Sub txtOpisPotencjalu_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOpisPotencjalu.LostFocus
        KoniecEdycjiTxt(txtOpisPotencjalu)
    End Sub

    '-----------------------------------------------

    Private Sub txtSprzedazOpiekun_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSprzedazOpiekun.GotFocus
        StartEdycjiTxt(txtSprzedazOpiekun)
    End Sub
    Private Sub txtSprzedazOpiekun_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSprzedazOpiekun.LostFocus
        KoniecEdycjiTxt(txtSprzedazOpiekun)
    End Sub

    Private Sub cbxSprzedazOKontr_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxSprzedazOKontR.GotFocus
        StartEdycjiCbx(cbxSprzedazOKontR)
    End Sub
    Private Sub cbxSprzedazOKontr_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxSprzedazOKontR.LostFocus
        KoniecEdycjiCbx(cbxSprzedazOKontR)
    End Sub


    Private Sub cbxSprzedazOKonD_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxSprzedazOKonD.GotFocus
        StartEdycjiCbx(cbxSprzedazOKonD)
    End Sub
    Private Sub cbxSprzedazOKonD_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxSprzedazOKonD.LostFocus
        KoniecEdycjiCbx(cbxSprzedazOKonD)
    End Sub

    Private Sub cbxSprzedazNKontr_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxSprzedazNKontR.GotFocus
        StartEdycjiCbx(cbxSprzedazNKontR)
    End Sub
    Private Sub cbxSprzedazNKontr_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxSprzedazNKontR.LostFocus
        KoniecEdycjiCbx(cbxSprzedazNKontR)
    End Sub

    Private Sub cbxSprzedazNKonD_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxSprzedazNKonD.GotFocus
        StartEdycjiCbx(cbxSprzedazNKonD)
    End Sub
    Private Sub cbxSprzedazNKonD_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxSprzedazNKonD.LostFocus
        KoniecEdycjiCbx(cbxSprzedazNKonD)
    End Sub

    Private Sub txtSprzedazUwagi_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSprzedazUwagi.GotFocus
        StartEdycjiTxt(txtSprzedazUwagi)
    End Sub
    Private Sub txtSprzedazUwagi_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSprzedazUwagi.LostFocus
        KoniecEdycjiTxt(txtSprzedazUwagi)
    End Sub












    Private Sub cmdWysylaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWysylaj.Click
        If TestujParametryeMail() Then
            Select Case ParL_eMailService
                Case "Wbudowany"
                    Dim Form As New SendeMail(txteMail.Text, "", "", "")
                    Form.Show()
                Case "MS Outlook"
                    SendByOutlook(txteMail.Text, "", "", "", "Edytuj")
                Case "Mozilla Thunderbird"
                    SendByThunderbird(txteMail.Text, "", "", "", "", "")

                Case Else
            End Select

            'Dim Form As New SendeMail(txteMail.Text, "", "", "")
            'Form.ShowDialog()
        End If
    End Sub


    Private Sub btnPop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPop.Click
        ZapiszDane()
        oCoDalej = "POPRZEDNI"
        oAktualizuj = True
        Me.Close()
    End Sub

    Private Sub btnNast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNast.Click
        ZapiszDane()
        oCoDalej = "NASTÊPNY"
        oAktualizuj = True
        Me.Close()
    End Sub

    Private Sub cmdNrKarty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNrKarty.Click
        oCoMamRobic = "WYBIERAJ"
        oNumer = Trim(txtIdent.Text)
        Dim form As New WybierzNrKarty
        form.ShowDialog()
        If Len(Trim(oNumer)) > 0 Then
            txtIdent.Text = oNumer
        End If
    End Sub

    Private Sub cmdNowyNrKarty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNowyNrKarty.Click
        Dim LastIdent As Long = 0
        If Len(Trim(txtIdent.Text)) = 0 Then
            LastIdent = Val(OstatniIdKlienta())
            txtIdent.Text = Mid(Trim(Str(1000000 + LastIdent + 1)), 2, 6)
        Else
        End If
    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TestujParametryeMail() Then
            Select Case ParL_eMailService
                Case "Wbudowany"
                    Dim Form As New SendeMail(txteMail.Text, "", "", "")
                    Form.Show()
                Case "MS Outlook"
                    SendByOutlook(txteMail.Text, "", "", "", "Edytuj")
                Case "Mozilla Thunderbird"
                    SendByThunderbird(txteMail.Text, "", "", "", "", "")

                Case Else
            End Select
            'Dim Form As New SendeMail(txteMail.Text, "", "", "")
            'Form.ShowDialog()
        End If
    End Sub



    Private Sub chkKartyPayback_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkKartyPayback.CheckedChanged
        If chkKartyPayback.Checked Then
            lblNryKartPayback.Visible = True
            txtNryKartPayBack.Visible = True
            cmdEdytujNumeryKartPayBack.Visible = True
        Else
            lblNryKartPayback.Visible = False
            txtNryKartPayBack.Visible = False
            cmdEdytujNumeryKartPayBack.Visible = False
        End If
    End Sub




    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        'System.Diagnostics.Process.Start( _
        '       "C:\Program Files\Internet Explorer\IExplore.exe", _
        '       "https://www.payback.pl/pb/id/36624/index.html/?utm_source=pryzmatwww&utm_medium=homepage&utm_campaign=ekupony")

        System.Diagnostics.Process.Start(
               "C:\Program Files\Internet Explorer\IExplore.exe",
               Trim(Par_wwwPAYBACK))
    End Sub

    Private Sub cmdHarmonogramWizyt_Click(sender As Object, e As EventArgs) Handles cmdHarmonogramWizyt.Click
        Dim Form As New KatalogHarmonogramWizyt(txtSprzedazOpiekun.Text)
        Form.Show()
    End Sub










    Private Sub chkU1_CheckedChanged(sender As Object, e As EventArgs) Handles chkU1.CheckedChanged
        If chkU1.Checked Then
            chkU1.ForeColor = Color.Red
        Else
            chkU1.ForeColor = Color.Navy
        End If
    End Sub

    Private Sub chkU2_CheckedChanged(sender As Object, e As EventArgs) Handles chkU2.CheckedChanged
        If chkU2.Checked Then
            chkU2.ForeColor = Color.Red
        Else
            chkU2.ForeColor = Color.Navy
        End If
    End Sub

    Private Sub chkU3_CheckedChanged(sender As Object, e As EventArgs) Handles chkU3.CheckedChanged
        If chkU3.Checked Then
            chkU3.ForeColor = Color.Red
        Else
            chkU3.ForeColor = Color.Navy
        End If
    End Sub

    Private Sub chkU4_CheckedChanged(sender As Object, e As EventArgs) Handles chkU4.CheckedChanged
        If chkU4.Checked Then
            chkU4.ForeColor = Color.Red
        Else
            chkU4.ForeColor = Color.Navy
        End If
    End Sub

    Private Sub chkU5_CheckedChanged(sender As Object, e As EventArgs) Handles chkU5.CheckedChanged
        If chkU5.Checked Then
            chkU5.ForeColor = Color.Red
        Else
            chkU5.ForeColor = Color.Navy
        End If
    End Sub

    Private Sub chkU6_CheckedChanged(sender As Object, e As EventArgs) Handles chkU6.CheckedChanged
        If chkU6.Checked Then
            chkU6.ForeColor = Color.Red
        Else
            chkU6.ForeColor = Color.Navy
        End If
    End Sub

    Private Sub chkU7_CheckedChanged(sender As Object, e As EventArgs) Handles chkU7.CheckedChanged
        If chkU7.Checked Then
            chkU7.ForeColor = Color.Red
        Else
            chkU7.ForeColor = Color.Navy
        End If
    End Sub

    Private Sub chkU8_CheckedChanged(sender As Object, e As EventArgs) Handles chkU8.CheckedChanged
        If chkU8.Checked Then
            chkU8.ForeColor = Color.Red
        Else
            chkU8.ForeColor = Color.Navy
        End If
    End Sub

    Private Sub chkU9_CheckedChanged(sender As Object, e As EventArgs) Handles chkU9.CheckedChanged
        If chkU9.Checked Then
            chkU9.ForeColor = Color.Red
        Else
            chkU9.ForeColor = Color.Navy
        End If
    End Sub




























    Private Sub cmdKontaktPromotor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKontaktPromotor.Click
        'cmdDodajKontakty_Click(sender, e)
        Dim Form As New KatalogKontaktow(txtIdent.Text)
        Form.ShowDialog()
        '---------------------------
        ' odswierz info o kontaktach...
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionKontakty.Open()
            '    If dbConnectionKontakty.State = ConnectionState.Open Then
            '        dbDataAdapterKontakty.Fill(DataSetKontakty)
            '        dbConnectionKontakty.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnectionKontakty.Open()
                If sqlConnectionKontakty.State = ConnectionState.Open Then
                    sqlDataAdapterKontakty.Fill(DataSetKontakty)
                    sqlConnectionKontakty.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub

    'Ten sam efekt uzyskamy po klikniêciu na tabele KONTAKTY
    Private Sub dagKontakty_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagKontakty.DoubleClick
        cmdKontaktPromotor_Click(sender, e)
    End Sub




























    '*********************************************************
    '** CoDalej
    '*********************************************************

    'Public cdUNIQID As String      'UNIQID Text, 40
    'Public cdIDENTKLIENTA As String      'IDENTKLIENTA Text, 40
    'Public cdIDENT As String             'IDENT        Text, 40
    'Public cdOPIS As String              'OPIS         memo
    'Public cdWersja As Integer           'WERSJA       Integer

    Private Sub InitCoDalej()
        cdUNIQID = Guid.NewGuid.ToString
        cdIDENTKLIENTA = oIdentKli
        cdIDENT = ""
        cdOPIS = ""
        cdWersja = 1          'WERSJA         Integer, 2
    End Sub

    Private Sub PobierzCoDalej()
        'pobierz wartosci przed aktualizacja
        cdUNIQID = GetTxtField(dagCoDalej, 0)
        cdIDENTKLIENTA = GetTxtField(dagCoDalej, 1)
        cdIDENT = GetTxtField(dagCoDalej, 2)
        cdOPIS = GetTxtField(dagCoDalej, 3)
        cdWersja = GetTxtField(dagCoDalej, 4)
    End Sub

    Private Sub menuSingleLineCoDalej_Click(sender As Object, e As EventArgs) Handles menuSingleLineCoDalej.Click
        dagCoDalej.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagCoDalej.Refresh()
    End Sub

    Private Sub menuMultiLineCoDalej_Click(sender As Object, e As EventArgs) Handles menuMultiLineCoDalej.Click
        dagCoDalej.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagCoDalej.Refresh()
    End Sub

    Private Sub menuOdswiezCoDalej_Click(sender As Object, e As EventArgs) Handles menuOdswiezCoDalej.Click
        OdswiezBazeCoDalej()
    End Sub

    Private Sub cmdCoDalej_Click(sender As Object, e As EventArgs) Handles cmdCoDalej.Click
        'cmdDodajCoDalej_Click(sender, e)
        Dim Form As New KatalogCoDalej(txtIdent.Text)
        Form.ShowDialog()
        '---------------------------
        ' odswierz info o kontaktach...
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionCoDalej.Open()
            '    If dbConnectionCoDalej.State = ConnectionState.Open Then
            '        dbDataAdapterCoDalej.Fill(DataSetCoDalej)
            '        dbConnectionCoDalej.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnectionCoDalej.Open()
                If sqlConnectionCoDalej.State = ConnectionState.Open Then
                    sqlDataAdapterCoDalej.Fill(DataSetCoDalej)
                    sqlConnectionCoDalej.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub

    'Ten sam efekt uzyskamy po klikniêciu na tabele CoDalej
    Private Sub dagCoDalej_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagCoDalej.DoubleClick
        cmdCoDalej_Click(sender, e)
    End Sub


    Private Sub OdswiezBazeCoDalej()
        DataSetCoDalej.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionCoDalej.Open()
            '    If dbConnectionCoDalej.State = ConnectionState.Open Then
            '        dbDataAdapterCoDalej.Fill(DataSetCoDalej)
            '        dbConnectionCoDalej.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnectionCoDalej.Open()
                If sqlConnectionCoDalej.State = ConnectionState.Open Then
                    sqlDataAdapterCoDalej.Fill(DataSetCoDalej)
                    sqlConnectionCoDalej.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub








    '*********************************************************
    '** Kontakty
    '*********************************************************

    'Public oUniqIdKon As String           'UNIQID            varchar(40)
    'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
    'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
    'Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
    'Public oTematKon As String            'TEMAT            Text, 50
    'Public oUwagiKon As String            'UWAGI            Memo
    'Public oWersjaKon As Integer          'WERSJA           Integer

    Private Sub InitKontakty()
        oUniqIdKon = Guid.NewGuid.ToString
        oIdentKon = oIdentKli
        oDataKon = SysData()
        oPracownikKon = ""
        oTematKon = ""
        oUwagiKon = ""
        oWersjaKon = 1          'WERSJA         Integer, 2
    End Sub

    Private Sub PobierzKontakty()
        'pobierz wartosci przed aktualizacja
        oUniqIdKon = GetTxtField(dagKontakty, 0)
        oIdentKon = GetTxtField(dagKontakty, 1)
        oDataKon = GetTxtField(dagKontakty, 2)
        oPracownikKon = GetTxtField(dagKontakty, 3)
        oTematKon = GetTxtField(dagKontakty, 4)
        oUwagiKon = GetTxtField(dagKontakty, 5)
        oWersjaKon = GetTxtField(dagKontakty, 6)
    End Sub


    Private Sub menuSingleLine_Click(sender As Object, e As EventArgs) Handles menuSingleLine.Click
        dagKontakty.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagKontakty.Refresh()
    End Sub

    Private Sub menuMultiLine_Click(sender As Object, e As EventArgs) Handles menuMultiLine.Click
        dagKontakty.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagKontakty.Refresh()
    End Sub

    Private Sub menuOdswiezkontakty_Click(sender As Object, e As EventArgs) Handles menuOdswiezKontakty.Click
        OdswiezBazekontakty()
    End Sub

    Private Sub OdswiezBazekontakty()
        DataSetKontakty.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionKontakty.Open()
            '    If dbConnectionKontakty.State = ConnectionState.Open Then
            '        dbDataAdapterKontakty.Fill(DataSetKontakty)
            '        dbConnectionKontakty.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnectionKontakty.Open()
                If sqlConnectionKontakty.State = ConnectionState.Open Then
                    sqlDataAdapterKontakty.Fill(DataSetKontakty)
                    sqlConnectionKontakty.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub









    '*********************************************************
    '** RZUKontakty
    '*********************************************************

    Private Sub dagRZUKontakty_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormyRZUKontakty Then
            If dagRZUKontakty.CurrentCell Is Nothing Then
                stbRZUKontakty.Panels(1).Text = "Wi: "
                stbRZUKontakty.Panels(2).Text = "Ko: "
            Else
                stbRZUKontakty.Panels(1).Text = "Wi: " & dagRZUKontakty.CurrentCell.RowIndex.ToString
                stbRZUKontakty.Panels(2).Text = "Ko: " & dagRZUKontakty.CurrentCell.ColumnIndex.ToString
            End If
        End If
    End Sub

    '-=------------------------------------------
    'obsluga zegara scrollingu poziomego
    '-=------------------------------------------
    'Private Sub HorizScrollTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HorizScrollTimer.Tick
    '    HorizScrollTimer.Enabled = False
    '    If Len(HorizScrollLastTime) = 0 Then
    '        'nie zainicjowano scrollingu
    '    Else
    '        'zainicjowano scrolling - sprawdz czy juz trwa 2 sec
    '        If IleSekundMinelo(HorizScrollLastTime, SysGodz()) >= 0 Then
    '            If dagRZUKontakty.RowCount > 0 Then
    '                If ScrollXOffset <> 10000 * dagRZUKontakty.FirstDisplayedScrollingColumnIndex +
    '                                dagRZUKontakty.GetCellDisplayRectangle(dagRZUKontakty.FirstDisplayedScrollingColumnIndex, 0, True).Left -
    '                                StartPointInCell00.X Then

    '                    PokazFiltryColumnDGV(me.uslugidod, dagRZUKontakty, DataViewRZUKontakty.Table.Columns.Count, stbFiltryRZUKontakty, StartFormyRZUKontakty)
    '                    '-----------------------------
    '                    ScrollXOffset = 10000 * dagRZUKontakty.FirstDisplayedScrollingColumnIndex +
    '                                dagRZUKontakty.GetCellDisplayRectangle(dagRZUKontakty.FirstDisplayedScrollingColumnIndex, 0, True).Left -
    '                                StartPointInCell00.X
    '                    HorizScrollLastTime = ""
    '                End If
    '            Else
    '                PokazFiltryColumnDGV(me.uslugidod, dagRZUKontakty, DataViewRZUKontakty.Table.Columns.Count, stbFiltryRZUKontakty, StartFormyRZUKontakty)
    '            End If
    '        Else
    '        End If
    '    End If
    '    HorizScrollTimer.Enabled = True
    'End Sub


    'Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
    '    Dim rc As DataGridViewSelectedRowCollection = dagRZUKontakty.SelectedRows
    '    If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
    '        If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
    '            Dim nextrow As Integer = rc(0).Index + 1
    '            dagRZUKontakty.Rows(nextrow).Selected = True
    '        End If
    '    End If
    'End Sub



    '*********************************************************
    '***** Filtrowanie wg poszczególnych pól
    '*********************************************************

    'AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

    Private Sub txtFiXXRZUKontakty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Not StartFormyRZUKontakty Then
            Select Case e.KeyCode
                'Case Keys.Enter
                'Case Keys.Insert, Keys.Add
                'Case Keys.Delete, Keys.Subtract

                Case Keys.Down
                    PrzejdzDoDGV(dagRZUKontakty, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXXRZUKontakty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormyRZUKontakty Then
            FiltrujDataViewDGV(Me.UslugiDod, dagRZUKontakty, DataViewRZUKontakty.Table.Columns.Count, DataViewRZUKontakty, stbRZUKontakty)
        End If
    End Sub
    Private Sub txtFiXXRZUKontakty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXXRZUKontakty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXXRZUKontakty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me.UslugiDod, dagRZUKontakty, Mid(sender.name, 6, 3), "RZUKontakty")
    End Sub
    Private Sub cmdFiXXRZUKontakty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiFiltraCmd(Me.UslugiDod, sender)
    End Sub
    Private Sub cmdFiXXRZUKontakty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraCmd(Me.UslugiDod, sender)
    End Sub


    '-------------------------------------------------
    ' filtrowanie 
    '-------------------------------------------------

    Private Sub cmdWszystkoRZUKontakty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWszystkoRZUKontakty.Click
        Me.Cursor = Cursors.WaitCursor
        oLokalnyFiltr = ""
        Try
            StartFormyRZUKontakty = True
            CzyscFiltryDGV(Me.UslugiDod, dagRZUKontakty, DataViewRZUKontakty.Table.Columns.Count, DataViewRZUKontakty, stbRZUKontakty)
            StartFormyRZUKontakty = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbRZUKontakty.Panels(0).Text = "Il.zap.: " & DataViewRZUKontakty.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagRZUKontakty_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs)
        Using b As SolidBrush = New SolidBrush(dagRZUKontakty.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagRZUKontakty.DefaultCellStyle.Font,
                      b,
                      e.RowBounds.Location.X + 15,
                      e.RowBounds.Location.Y + 4)
        End Using
    End Sub



    '*************************************************
    '*** Odœwiezanie filtrow kolumnowych 
    '*************************************************

    Private Sub stbFiltryRZUKontakty_PanelClick(sender As Object, e As StatusBarPanelClickEventArgs) Handles stbFiltryRZUKontakty.PanelClick
        If Not StartFormyRZUKontakty Then
            PokazFiltryColumnDGV(Me.UslugiDod, dagRZUKontakty, DataViewRZUKontakty.Table.Columns.Count, stbFiltryRZUKontakty, StartFormyRZUKontakty)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagRZUKontakty_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagRZUKontakty.CellFormatting
    '    If Not StartFormyRZUKontakty Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagRZUKontakty, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagRZUKontakty, e.RowIndex, 4)

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

    Private Sub dagRZUKontakty_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs)
        If Not StartFormyRZUKontakty Then
            PokazFiltryColumnDGV(Me.UslugiDod, dagRZUKontakty, DataViewRZUKontakty.Table.Columns.Count, stbFiltryRZUKontakty, StartFormyRZUKontakty)
        End If
    End Sub

    Private Sub dagRZUKontakty_Resize(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormyRZUKontakty Then
            PokazFiltryColumnDGV(Me.UslugiDod, dagRZUKontakty, DataViewRZUKontakty.Table.Columns.Count, stbFiltryRZUKontakty, StartFormyRZUKontakty)
        End If
    End Sub

    Private Sub dagRZUKontakty_Scroll(sender As Object, e As ScrollEventArgs)
        'If (Not StartFormyRZUKontakty) And (DataViewRZUKontakty.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(me.uslugidod, dagRZUKontakty, DataViewRZUKontakty.Table.Columns.Count, stbFiltryRZUKontakty, StartFormyRZUKontakty)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(me.uslugidod, dagRZUKontakty, DataViewRZUKontakty.Table.Columns.Count, stbFiltryRZUKontakty, StartFormyRZUKontakty)
        '        End If
        '    End If
        'End If

        If (Not StartFormyRZUKontakty) And (DataViewRZUKontakty.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagRZUKontakty_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormyRZUKontakty Then
            PokazFiltryColumnDGV(Me.UslugiDod, dagRZUKontakty, DataViewRZUKontakty.Table.Columns.Count, stbFiltryRZUKontakty, StartFormyRZUKontakty)
        End If
    End Sub

    Private Sub dagRZUKontakty_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None     '"the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                'If Not StartFormyRZUKontakty Then
                '    PokazFiltryColumnDGV(me.uslugidod, dagRZUKontakty, DataViewRZUKontakty.Table.Columns.Count, stbFiltryRZUKontakty, StartFormyRZUKontakty)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormyRZUKontakty Then
                    PokazFiltryColumnDGV(Me.UslugiDod, dagRZUKontakty, DataViewRZUKontakty.Table.Columns.Count, stbFiltryRZUKontakty, StartFormyRZUKontakty)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagRZUKontakty_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
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
                stbRZUKontakty.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbRZUKontakty.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagRZUKontakty(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbRZUKontakty.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbRZUKontakty.Panels(3).Text = "Sort: " & DataSetRZUKontakty.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbRZUKontakty.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagRZUKontakty(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbRZUKontakty.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbRZUKontakty.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub




    Private Sub dagRZUKontakty_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagRZUKontakty.DoubleClick
        cmdKontaktPromotor_Click(sender, e)
    End Sub

    'Private Sub dagRZUKontakty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dagRZUKontakty.KeyDown
    '    Select Case e.KeyCode
    '        Case Keys.Enter
    '            'If _CoMamRobic = "WYBIERAJ" Then
    '            '    cmdStop_Click(sender, e)
    '            'Else
    '            '    cmdEdytuj_Click(sender, e)
    '            'End If
    '        Case Keys.Insert
    '            cmdDodaj_Click(sender, e)
    '        Case Keys.Delete
    '            cmdUsuñ_Click(sender, e)
    '        Case Else
    '    End Select
    'End Sub






    'Public oUniqIdKon As String           'UNIQID            varchar(40)
    'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
    'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
    'Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
    'Public oTematKon As String            'TEMAT            Text, 50
    'Public oUwagiKon As String            'UWAGI            Memo
    'Public oWersjaKon As Integer          'WERSJA           Integer

    Private Sub InitRZUKontakty()
        oUniqIdKon = Guid.NewGuid.ToString
        oIdentKon = oIdentKli
        oDataKon = SysData()
        oPracownikKon = ""
        oTematKon = ""
        oUwagiKon = ""
        oWersjaKon = 1          'WERSJA         Integer, 2
    End Sub

    Private Sub PobierzRZUKontakty()
        'pobierz wartosci przed aktualizacja
        oUniqIdKon = GetTxtField(dagRZUKontakty, 0)
        oIdentKon = GetTxtField(dagRZUKontakty, 1)
        oDataKon = GetTxtField(dagRZUKontakty, 2)
        oPracownikKon = GetTxtField(dagRZUKontakty, 3)
        oTematKon = GetTxtField(dagRZUKontakty, 4)
        oUwagiKon = GetTxtField(dagRZUKontakty, 5)
        oWersjaKon = GetTxtField(dagRZUKontakty, 6)
    End Sub


    Private Sub menuRZUSingleLine_Click(sender As Object, e As EventArgs) Handles menuRZUSingleLine.Click
        dagRZUKontakty.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagRZUKontakty.Refresh()
    End Sub

    Private Sub menuRZUMultiLine_Click(sender As Object, e As EventArgs) Handles menuRZUMultiLine.Click
        dagRZUKontakty.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagRZUKontakty.Refresh()
    End Sub

    Private Sub menuOdswiezRZUKontakty_Click(sender As Object, e As EventArgs) Handles menuOdswiezRZUKontakty.Click
        OdswiezBazeRZUKontakty()
    End Sub

    Private Sub OdswiezBazeRZUKontakty()
        DataSetRZUKontakty.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionRZUKontakty.Open()
            '    If dbConnectionRZUKontakty.State = ConnectionState.Open Then
            '        dbDataAdapterRZUKontakty.Fill(DataSetRZUKontakty)
            '        dbConnectionRZUKontakty.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnectionRZUKontakty.Open()
                If sqlConnectionRZUKontakty.State = ConnectionState.Open Then
                    sqlDataAdapterRZUKontakty.Fill(DataSetRZUKontakty)
                    sqlConnectionRZUKontakty.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub











    '*********************************************************
    '** Katalog Osob
    '*********************************************************

    Private Sub dagOsoby_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagOsoby.CurrentCellChanged
        If Not StartFormyOsoby Then
            If dagOsoby.CurrentCell Is Nothing Then
                stbOsoby.Panels(1).Text = "Wi: "
                stbOsoby.Panels(2).Text = "Ko: "
            Else
                stbOsoby.Panels(1).Text = "Wi: " & dagOsoby.CurrentCell.RowIndex.ToString
                stbOsoby.Panels(2).Text = "Ko: " & dagOsoby.CurrentCell.ColumnIndex.ToString
            End If
        End If
    End Sub

    '-=------------------------------------------
    'obsluga zegara scrollingu poziomego
    '-=------------------------------------------
    'Private Sub HorizScrollTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HorizScrollTimer.Tick
    '    HorizScrollTimer.Enabled = False
    '    If Len(HorizScrollLastTime) = 0 Then
    '        'nie zainicjowano scrollingu
    '    Else
    '        'zainicjowano scrolling - sprawdz czy juz trwa 2 sec
    '        If IleSekundMinelo(HorizScrollLastTime, SysGodz()) >= 0 Then
    '            If dagOsoby.RowCount > 0 Then
    '                If ScrollXOffset <> 10000 * dagOsoby.FirstDisplayedScrollingColumnIndex +
    '                                dagOsoby.GetCellDisplayRectangle(dagOsoby.FirstDisplayedScrollingColumnIndex, 0, True).Left -
    '                                StartPointInCell00.X Then

    '                    PokazFiltryColumnDGV(Me.Identyfikacja, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltryOsoby, StartFormyOsoby)
    '                    '-----------------------------
    '                    ScrollXOffset = 10000 * dagOsoby.FirstDisplayedScrollingColumnIndex +
    '                                dagOsoby.GetCellDisplayRectangle(dagOsoby.FirstDisplayedScrollingColumnIndex, 0, True).Left -
    '                                StartPointInCell00.X
    '                    HorizScrollLastTime = ""
    '                End If
    '            Else
    '                PokazFiltryColumnDGV(Me.Identyfikacja, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltryOsoby, StartFormyOsoby)
    '            End If
    '        Else
    '        End If
    '    End If
    '    HorizScrollTimer.Enabled = True
    'End Sub


    'Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
    '    Dim rc As DataGridViewSelectedRowCollection = dagOsoby.SelectedRows
    '    If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
    '        If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
    '            Dim nextrow As Integer = rc(0).Index + 1
    '            dagOsoby.Rows(nextrow).Selected = True
    '        End If
    '    End If
    'End Sub



    '*********************************************************
    '***** Filtrowanie wg poszczególnych pól
    '*********************************************************

    'AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

    Private Sub txtFiXXOsoby_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Not StartFormyOsoby Then
            Select Case e.KeyCode
                'Case Keys.Enter
                'Case Keys.Insert, Keys.Add
                'Case Keys.Delete, Keys.Subtract

                Case Keys.Down
                    PrzejdzDoDGV(dagOsoby, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXXOsoby_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormyOsoby Then
            FiltrujDataViewDGV(Me.Identyfikacja, dagOsoby, DataViewOsoby.Table.Columns.Count, DataViewOsoby, stbOsoby)
        End If
    End Sub
    Private Sub txtFiXXOsoby_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXXOsoby_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXXOsoby_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me.Identyfikacja, dagOsoby, Mid(sender.name, 6, 3), "Osoby")
    End Sub
    Private Sub cmdFiXXOsoby_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiFiltraCmd(Me.Identyfikacja, sender)
    End Sub
    Private Sub cmdFiXXOsoby_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraCmd(Me.Identyfikacja, sender)
    End Sub


    '-------------------------------------------------
    ' filtrowanie 
    '-------------------------------------------------

    Private Sub cmdWszystkoOsoby_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWszystkoOsoby.Click
        Me.Cursor = Cursors.WaitCursor
        oLokalnyFiltr = ""
        Try
            StartFormyOsoby = True
            CzyscFiltryDGV(Me.Identyfikacja, dagOsoby, DataViewOsoby.Table.Columns.Count, DataViewOsoby, stbOsoby)
            StartFormyOsoby = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbOsoby.Panels(0).Text = "Il.zap.: " & DataViewOsoby.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagOsoby_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagOsoby.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagOsoby.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagOsoby.DefaultCellStyle.Font,
                      b,
                      e.RowBounds.Location.X + 15,
                      e.RowBounds.Location.Y + 4)
        End Using
    End Sub



    '*************************************************
    '*** Odœwiezanie filtrow kolumnowych 
    '*************************************************

    Private Sub stbFiltryOsoby_PanelClick(sender As Object, e As StatusBarPanelClickEventArgs) Handles stbFiltryOsoby.PanelClick
        If Not StartFormyOsoby Then
            PokazFiltryColumnDGV(Me.Identyfikacja, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltryOsoby, StartFormyOsoby)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    Private Sub dagOsoby_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagOsoby.CellFormatting
        If Not StartFormyOsoby Then
            '-----------------------
            'zmieniamy ForeColor
            Dim vip As Boolean = GetLogField(dagOsoby, e.RowIndex, 2)

            If vip Then
                e.CellStyle.ForeColor = Color.Red
            Else
                e.CellStyle.ForeColor = Color.Navy
            End If
            '-----------------------
            ''zmieniamy BackColor
            'Dim Wal As String = GetTxtField(dagKursyWalut, e.RowIndex, 0)
            'Select Case Wal
            '    Case "EUR"
            '        e.CellStyle.BackColor = Color.White
            '    Case Else
            'End Select
            '-----------------------
        End If
    End Sub




    '*********************************************************
    '***** obsluga dag
    '*********************************************************

    Private Sub dagOsoby_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagOsoby.ColumnWidthChanged
        If Not StartFormyOsoby Then
            PokazFiltryColumnDGV(Me.Identyfikacja, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltryOsoby, StartFormyOsoby)
        End If
    End Sub

    Private Sub dagOsoby_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagOsoby.Resize
        If Not StartFormyOsoby Then
            PokazFiltryColumnDGV(Me.Identyfikacja, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltryOsoby, StartFormyOsoby)
        End If
    End Sub

    Private Sub dagOsoby_Scroll(sender As Object, e As ScrollEventArgs) Handles dagOsoby.Scroll
        'If (Not StartFormyOsoby) And (DataViewOsoby.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me.Identyfikacja, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltryOsoby, StartFormyOsoby)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me.Identyfikacja, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltryOsoby, StartFormyOsoby)
        '        End If
        '    End If
        'End If

        If (Not StartFormyOsoby) And (DataViewOsoby.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagOsoby_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagOsoby.SizeChanged
        If Not StartFormyOsoby Then
            PokazFiltryColumnDGV(Me.Identyfikacja, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltryOsoby, StartFormyOsoby)
        End If
    End Sub

    Private Sub dagOsoby_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagOsoby.MouseMove
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None     '"the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                'If Not StartFormyOsoby Then
                '    PokazFiltryColumnDGV(Me.Identyfikacja, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltryOsoby, StartFormyOsoby)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormyOsoby Then
                    PokazFiltryColumnDGV(Me.Identyfikacja, dagOsoby, DataViewOsoby.Table.Columns.Count, stbFiltryOsoby, StartFormyOsoby)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagOsoby_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagOsoby.MouseUp
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
                stbOsoby.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbOsoby.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagOsoby(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbOsoby.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbOsoby.Panels(3).Text = "Sort: " & DataSetOsoby.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbOsoby.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagOsoby(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbOsoby.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbOsoby.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub






    Private Sub dagOsoby_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagOsoby.DoubleClick
        'If _CoMamRobic = "WYBIERAJ" Then
        '    cmdStop_Click(sender, e)
        'Else
        '    cmdEdytuj_Click(sender, e)
        'End If
        cmdEdytujOsoby_Click(sender, e)
    End Sub

    Private Sub dagOsoby_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case e.KeyCode
            Case Keys.Enter
                'If oCoMamRobic = "WYBIERAJ" Then
                'cmdStopOsoby_Click(sender, e)
                'Else
                cmdEdytujOsoby_Click(sender, e)
                'End If
            Case Keys.Insert
                cmdDodajOsoby_Click(sender, e)
            Case Keys.Delete
                cmdUsuñOsoby_Click(sender, e)
            Case Else
        End Select
    End Sub



    Private Sub InitOsoby()
        oIdentOso = oIdentKli
        oOsobaOso = ""
        oVIPOso = False
        oeMailOso = ""
        oTelefonOso = ""
        oTelefonKomOso = ""
        oStanowiskoOso = ""
        oKompetencjeOso = ""
        oWersjaOso = 1          'WERSJA         Integer, 2
    End Sub

    Private Sub PobierzOsoby()
        'pobierz wartosci przed aktualizacja
        oIdentOso = GetTxtField(dagOsoby, 0)
        oOsobaOso = GetTxtField(dagOsoby, 1)
        oVIPOso = GetLogField(dagOsoby, 2)
        oeMailOso = GetTxtField(dagOsoby, 3)
        oTelefonOso = GetTxtField(dagOsoby, 4)
        oTelefonKomOso = GetTxtField(dagOsoby, 5)
        oStanowiskoOso = GetTxtField(dagOsoby, 6)
        oKompetencjeOso = GetTxtField(dagOsoby, 7)
        oWersjaOso = GetNumField(dagOsoby, 8)
    End Sub



    '***********************************************************
    '*** obsluga komend dla Osob kontaktowych
    '***********************************************************

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


    Private Sub OdswiezBazeOsoby()
        'DataSetOsoby.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionOsoby.Open()
            '    If dbConnectionOsoby.State = ConnectionState.Open Then
            '        dbDataAdapterOsoby.Fill(DataSetOsoby)
            '        dbConnectionOsoby.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnectionOsoby.Open()
                If sqlConnectionOsoby.State = ConnectionState.Open Then
                    sqlDataAdapterOsoby.Fill(DataSetOsoby)
                    sqlConnectionOsoby.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub

    Private Sub cmdDodajOsoby_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodajOsoby.Click
        oOperacja = "DODAJ"
        InitOsoby()
        '----------------------------------------------
        Do While True
            oAktualizuj = False
            Dim Form As New EdytujKatalogOsob
            Form.ShowDialog()
            Form.Dispose()
            If Not oAktualizuj Then
                Exit Do
            Else
                Dim foundRow As DataRow
                Dim NewRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentKli
                findTheseVals(1) = oOsobaOso
                foundRow = DataSetOsoby.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "Klient = " & oIdentKli & vbCrLf &
                        "Osoba = " & oOsobaOso,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetOsoby.Tables(0).NewRow()
                    NewRow("IDENTKLIENTA") = oIdentKli
                    NewRow("OSOBA") = oOsobaOso
                    NewRow("VIP") = oVIPOso
                    NewRow("EMAIL") = oeMailOso
                    NewRow("TELEFON") = oTelefonOso
                    NewRow("TELEFONKOM") = oTelefonKomOso
                    NewRow("STANOWISKO") = oStanowiskoOso
                    NewRow("KOMPETENCJE") = oKompetencjeOso
                    NewRow("WERSJA") = 1     'Wersja         Integer, 2

                    'dodaj rekord do DataSet
                    DataSetOsoby.Tables(0).Rows.Add(NewRow)
                    'wyswietl ilosc rekordow
                    AktualizujOsoby()
                    'aktualizuj DataGrid
                    dagOsoby.Update()
                    Exit Do
                End If
            End If
        Loop
    End Sub

    Private Sub cmdUsuñOsoby_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñOsoby.Click
        If DataSetOsoby.Tables(0).Rows.Count <= 0 Then
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
                PobierzOsoby()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentKli
                findTheseVals(1) = oOsobaOso
                foundRow = DataSetOsoby.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    AktualizujOsoby()
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdEdytujOsoby_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytujOsoby.Click
        If DataSetOsoby.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzOsoby()
            Dim Form As New EdytujKatalogOsob
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentKli
                findTheseVals(1) = oOsobaOso
                foundRow = DataSetOsoby.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'foundRow("IDENTKLIENTA") = oIdentKli
                    foundRow("OSOBA") = oOsobaOso
                    foundRow("VIP") = oVIPOso
                    foundRow("EMAIL") = oeMailOso
                    foundRow("TELEFON") = oTelefonOso
                    foundRow("TELEFONKOM") = oTelefonKomOso
                    foundRow("STANOWISKO") = oStanowiskoOso
                    foundRow("KOMPETENCJE") = oKompetencjeOso
                    foundRow("WERSJA") = ZmienWersje(oWersjaOso) 'Wersja         Integer, 2

                    'wyswietl ilosc rekordow
                    AktualizujOsoby()
                    'aktualizuj DataGrid
                    dagOsoby.Update()
                End If
            End If
        End If
    End Sub


    Private Sub menuOsobyEdytuj_Click(sender As Object, e As EventArgs) Handles menuOsobyEdytuj.Click
        cmdEdytujOsoby_Click(sender, e)
    End Sub

    Private Sub menuOsobyDodaj_Click(sender As Object, e As EventArgs) Handles menuOsobyDodaj.Click
        cmdDodajOsoby_Click(sender, e)
    End Sub

    Private Sub menuOsobyUsun_Click(sender As Object, e As EventArgs) Handles menuOsobyUsun.Click
        cmdUsuñOsoby_Click(sender, e)
    End Sub

    Private Sub MenuOsobySingleL_Click(sender As Object, e As EventArgs) Handles MenuOsobySingleL.Click
        dagOsoby.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagOsoby.Refresh()
    End Sub

    Private Sub menuOsobyMultiL_Click(sender As Object, e As EventArgs) Handles menuOsobyMultiL.Click
        dagOsoby.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagOsoby.Refresh()
    End Sub

    Private Sub menuOsobyOdswiez_Click(sender As Object, e As EventArgs) Handles menuOsobyOdswiez.Click
        OdswiezBazeOsoby()
    End Sub













    '*********************************************************
    '** Katalog AnalizyZakupu
    '*********************************************************

    Private Sub dagAnalizyZakupu_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAnalizyZakupu.CurrentCellChanged
        If Not StartFormyAnalizyZakupu Then
            If dagAnalizyZakupu.CurrentCell Is Nothing Then
                stbAnalizyZakupu.Panels(1).Text = "Wi: "
                stbAnalizyZakupu.Panels(2).Text = "Ko: "
            Else
                stbAnalizyZakupu.Panels(1).Text = "Wi: " & dagAnalizyZakupu.CurrentCell.RowIndex.ToString
                stbAnalizyZakupu.Panels(2).Text = "Ko: " & dagAnalizyZakupu.CurrentCell.ColumnIndex.ToString
            End If
        End If
    End Sub

    '-=------------------------------------------
    'obsluga zegara scrollingu poziomego
    '-=------------------------------------------
    'Private Sub HorizScrollTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HorizScrollTimer.Tick
    '    HorizScrollTimer.Enabled = False
    '    If Len(HorizScrollLastTime) = 0 Then
    '        'nie zainicjowano scrollingu
    '    Else
    '        'zainicjowano scrolling - sprawdz czy juz trwa 2 sec
    '        If IleSekundMinelo(HorizScrollLastTime, SysGodz()) >= 0 Then
    '            If dagAnalizyZakupu.RowCount > 0 Then
    '                If ScrollXOffset <> 10000 * dagAnalizyZakupu.FirstDisplayedScrollingColumnIndex +
    '                                dagAnalizyZakupu.GetCellDisplayRectangle(dagAnalizyZakupu.FirstDisplayedScrollingColumnIndex, 0, True).Left -
    '                                StartPointInCell00.X Then

    '                    PokazFiltryColumnDGV(Me.AnalizyZakupu, dagAnalizyZakupu, DataViewAnalizyZakupu.Table.Columns.Count, stbFiltryAnalizyZakupu, StartFormyAnalizyZakupu)
    '                    '-----------------------------
    '                    ScrollXOffset = 10000 * dagAnalizyZakupu.FirstDisplayedScrollingColumnIndex +
    '                                dagAnalizyZakupu.GetCellDisplayRectangle(dagAnalizyZakupu.FirstDisplayedScrollingColumnIndex, 0, True).Left -
    '                                StartPointInCell00.X
    '                    HorizScrollLastTime = ""
    '                End If
    '            Else
    '                PokazFiltryColumnDGV(Me.AnalizyZakupu, dagAnalizyZakupu, DataViewAnalizyZakupu.Table.Columns.Count, stbFiltryAnalizyZakupu, StartFormyAnalizyZakupu)
    '            End If
    '        Else
    '        End If
    '    End If
    '    HorizScrollTimer.Enabled = True
    'End Sub


    'Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
    '    Dim rc As DataGridViewSelectedRowCollection = dagAnalizyZakupu.SelectedRows
    '    If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
    '        If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
    '            Dim nextrow As Integer = rc(0).Index + 1
    '            dagAnalizyZakupu.Rows(nextrow).Selected = True
    '        End If
    '    End If
    'End Sub



    '*********************************************************
    '***** Filtrowanie wg poszczególnych pól
    '*********************************************************

    'AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

    Private Sub txtFiXXAnalizyZakupu_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Not StartFormyAnalizyZakupu Then
            Select Case e.KeyCode
                'Case Keys.Enter
                'Case Keys.Insert, Keys.Add
                'Case Keys.Delete, Keys.Subtract

                Case Keys.Down
                    PrzejdzDoDGV(dagAnalizyZakupu, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXXAnalizyZakupu_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormyAnalizyZakupu Then
            FiltrujDataViewDGV(Me.AnalizyZakupu, dagAnalizyZakupu, DataViewAnalizyZakupu.Table.Columns.Count, DataViewAnalizyZakupu, stbAnalizyZakupu)
        End If
    End Sub
    Private Sub txtFiXXAnalizyZakupu_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXXAnalizyZakupu_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXXAnalizyZakupu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me.AnalizyZakupu, dagAnalizyZakupu, Mid(sender.name, 6, 3), "AnalizyZakupu")
    End Sub
    Private Sub cmdFiXXAnalizyZakupu_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiFiltraCmd(Me.AnalizyZakupu, sender)
    End Sub
    Private Sub cmdFiXXAnalizyZakupu_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraCmd(Me.AnalizyZakupu, sender)
    End Sub


    '-------------------------------------------------
    ' filtrowanie 
    '-------------------------------------------------

    Private Sub cmdWszystkoAnalizyZakupu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWszystkoAnalizyZakupu.Click
        Me.Cursor = Cursors.WaitCursor
        oLokalnyFiltr = ""
        Try
            StartFormyAnalizyZakupu = True
            CzyscFiltryDGV(Me.AnalizyZakupu, dagAnalizyZakupu, DataViewAnalizyZakupu.Table.Columns.Count, DataViewAnalizyZakupu, stbAnalizyZakupu)
            StartFormyAnalizyZakupu = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbAnalizyZakupu.Panels(0).Text = "Il.zap.: " & DataViewAnalizyZakupu.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagAnalizyZakupu_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagAnalizyZakupu.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagAnalizyZakupu.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagAnalizyZakupu.DefaultCellStyle.Font,
                      b,
                      e.RowBounds.Location.X + 15,
                      e.RowBounds.Location.Y + 4)
        End Using
    End Sub



    '*************************************************
    '*** Odœwiezanie filtrow kolumnowych 
    '*************************************************

    Private Sub stbFiltryAnalizyZakupu_PanelClick(sender As Object, e As StatusBarPanelClickEventArgs) Handles stbFiltryAnalizyZakupu.PanelClick
        If Not StartFormyAnalizyZakupu Then
            PokazFiltryColumnDGV(Me.AnalizyZakupu, dagAnalizyZakupu, DataViewAnalizyZakupu.Table.Columns.Count, stbFiltryAnalizyZakupu, StartFormyAnalizyZakupu)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagAnalizyZakupu_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagAnalizyZakupu.CellFormatting
    '    If Not StartFormyAnalizyZakupu Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagAnalizyZakupu, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagAnalizyZakupu, e.RowIndex, 4)

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

    Private Sub dagAnalizyZakupu_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagAnalizyZakupu.ColumnWidthChanged
        If Not StartFormyAnalizyZakupu Then
            PokazFiltryColumnDGV(Me.AnalizyZakupu, dagAnalizyZakupu, DataViewAnalizyZakupu.Table.Columns.Count, stbFiltryAnalizyZakupu, StartFormyAnalizyZakupu)
        End If
    End Sub

    Private Sub dagAnalizyZakupu_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAnalizyZakupu.Resize
        If Not StartFormyAnalizyZakupu Then
            PokazFiltryColumnDGV(Me.AnalizyZakupu, dagAnalizyZakupu, DataViewAnalizyZakupu.Table.Columns.Count, stbFiltryAnalizyZakupu, StartFormyAnalizyZakupu)
        End If
    End Sub

    Private Sub dagAnalizyZakupu_Scroll(sender As Object, e As ScrollEventArgs) Handles dagAnalizyZakupu.Scroll
        'If (Not StartFormyAnalizyZakupu) And (DataViewAnalizyZakupu.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me.AnalizyZakupu, dagAnalizyZakupu, DataViewAnalizyZakupu.Table.Columns.Count, stbFiltryAnalizyZakupu, StartFormyAnalizyZakupu)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me.AnalizyZakupu, dagAnalizyZakupu, DataViewAnalizyZakupu.Table.Columns.Count, stbFiltryAnalizyZakupu, StartFormyAnalizyZakupu)
        '        End If
        '    End If
        'End If

        If (Not StartFormyAnalizyZakupu) And (DataViewAnalizyZakupu.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagAnalizyZakupu_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAnalizyZakupu.SizeChanged
        If Not StartFormyAnalizyZakupu Then
            PokazFiltryColumnDGV(Me.AnalizyZakupu, dagAnalizyZakupu, DataViewAnalizyZakupu.Table.Columns.Count, stbFiltryAnalizyZakupu, StartFormyAnalizyZakupu)
        End If
    End Sub

    Private Sub dagAnalizyZakupu_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagAnalizyZakupu.MouseMove
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None     '"the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                'If Not StartFormyAnalizyZakupu Then
                '    PokazFiltryColumnDGV(Me.AnalizyZakupu, dagAnalizyZakupu, DataViewAnalizyZakupu.Table.Columns.Count, stbFiltryAnalizyZakupu, StartFormyAnalizyZakupu)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormyAnalizyZakupu Then
                    PokazFiltryColumnDGV(Me.AnalizyZakupu, dagAnalizyZakupu, DataViewAnalizyZakupu.Table.Columns.Count, stbFiltryAnalizyZakupu, StartFormyAnalizyZakupu)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagAnalizyZakupu_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagAnalizyZakupu.MouseUp
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
                stbAnalizyZakupu.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbAnalizyZakupu.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagAnalizyZakupu(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbAnalizyZakupu.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbAnalizyZakupu.Panels(3).Text = "Sort: " & DataSetAnalizyZakupu.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbAnalizyZakupu.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagAnalizyZakupu(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbAnalizyZakupu.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbAnalizyZakupu.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub






    Private Sub dagAnalizyZakupu_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAnalizyZakupu.DoubleClick
        'cmdEdytujAnalizyZakupu_Click(sender, e)
    End Sub

    Private Sub dagAnalizyZakupu_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        'Select Case e.KeyCode
        '    Case Keys.Enter
        '        cmdEdytujAnalizyZakupu_Click(sender, e)
        '    Case Keys.Insert
        '        cmdDodajAnalizyZakupu_Click(sender, e)
        '    Case Keys.Delete
        '        cmdUsuñAnalizyZakupu_Click(sender, e)
        '    Case Else
        'End Select
    End Sub



    Private Sub InitAnalizyZakupu()
    End Sub

    Private Sub PobierzAnalizyZakupu()
        'pobierz wartosci przed aktualizacja
    End Sub



    '***********************************************************
    '*** obsluga komend dla Osob kontaktowych
    '***********************************************************

    Private Sub AktualizujAnalizyZakupu()
        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionAnalizyZakupu.Open()
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try

            'If dbConnectionAnalizyZakupu.State = ConnectionState.Open Then
            '    oWystapilConcurrencyException = False
            '    '------------------------------------------------------------
            '    Try
            '        dbDataAdapterAnalizyZakupu.Update(DataSetAnalizyZakupu, "TABELA_AnalizyZakupu")
            '    Catch Ex As System.Exception
            '        ViewInLog(Ex, Me.Name(), Nothing)
            '    End Try
            '    '------------------------------------------------------------
            '    If oWystapilConcurrencyException = True Then
            '        dbDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
            '    End If
            '    dbConnectionAnalizyZakupu.Close()
            'End If
        Else
            Try
                sqlConnectionAnalizyZakupu.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionAnalizyZakupu.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterAnalizyZakupu.Update(DataSetAnalizyZakupu, "TABELA_AnalizyZakupu")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
                End If
                sqlConnectionAnalizyZakupu.Close()
            End If
        End If
    End Sub


    Private Sub OdswiezBazeAnalizyZakupu()
        DataSetAnalizyZakupu.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionAnalizyZakupu.Open()
            '    If dbConnectionAnalizyZakupu.State = ConnectionState.Open Then
            '        dbDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
            '        dbConnectionAnalizyZakupu.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnectionAnalizyZakupu.Open()
                If sqlConnectionAnalizyZakupu.State = ConnectionState.Open Then
                    sqlDataAdapterAnalizyZakupu.Fill(DataSetAnalizyZakupu)
                    sqlConnectionAnalizyZakupu.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub

    'Private Sub cmdDodajAnalizyZakupu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodajAnalizyZakupu.Click
    '    oOperacja = "DODAJ"
    '    InitAnalizyZakupu()
    '    '----------------------------------------------
    '    Do While True
    '        oAktualizuj = False
    '        Dim Form As New EdytujKatalogOsob
    '        Form.ShowDialog()
    '        Form.Dispose()
    '        If Not oAktualizuj Then
    '            Exit Do
    '        Else
    '            Dim foundRow As DataRow
    '            Dim NewRow As DataRow
    '            ' Create an array for the key values to find.
    '            Dim findTheseVals(1) As Object
    '            findTheseVals(0) = oIdentKli
    '            findTheseVals(1) = oOsobaOso
    '            foundRow = DataSetAnalizyZakupu.Tables(0).Rows.Find(findTheseVals)
    '            ' sprawdz czy jest...
    '            If Not (foundRow Is Nothing) Then
    '                MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
    '                    "Klient = " & oIdentKli & vbCrLf &
    '                    "Osoba = " & oOsobaOso,
    '                    "Uwaga :",
    '                    System.Windows.Forms.MessageBoxButtons.OK,
    '                    MessageBoxIcon.Information,
    '                    MessageBoxDefaultButton.Button1)
    '            Else
    '                'stworz nowy rekord
    '                NewRow = DataSetAnalizyZakupu.Tables(0).NewRow()
    '                NewRow("IDENTKLIENTA") = oIdentKli
    '                NewRow("OSOBA") = oOsobaOso
    '                NewRow("STANOWISKO") = oStanowiskoOso
    '                NewRow("KOMPETENCJE") = oKompetencjeOso
    '                NewRow("TELEFON") = oTelefonOso
    '                NewRow("EMAIL") = oeMailOso
    '                NewRow("VIP") = oVIPOso
    '                NewRow("WERSJA") = 1     'Wersja         Integer, 2

    '                'dodaj rekord do DataSet
    '                DataSetAnalizyZakupu.Tables(0).Rows.Add(NewRow)
    '                'wyswietl ilosc rekordow
    '                AktualizujAnalizyZakupu()
    '                'aktualizuj DataGrid
    '                dagAnalizyZakupu.Update()
    '                Exit Do
    '            End If
    '        End If
    '    Loop
    'End Sub

    'Private Sub cmdUsuñAnalizyZakupu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñAnalizyZakupu.Click
    '    If DataSetAnalizyZakupu.Tables(0).Rows.Count <= 0 Then
    '        MessageBox.Show("Nie potrafie usun¹æ nieistniej¹cego zapisu" & vbCrLf &
    '            "Proszê najpierw dopisaæ rekord do tabeli...",
    '            "Uwaga :",
    '            System.Windows.Forms.MessageBoxButtons.OK,
    '            MessageBoxIcon.Information,
    '            MessageBoxDefaultButton.Button1)
    '    Else
    '        If MessageBox.Show("Czy na pewno mam usun¹æ ten zapis ?",
    '                    "Proszê o potwierdzenie :",
    '                    System.Windows.Forms.MessageBoxButtons.YesNo,
    '                    MessageBoxIcon.Question,
    '                    MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then
    '            oOperacja = "USUN"
    '            PobierzAnalizyZakupu()
    '            Dim foundRow As DataRow
    '            ' Create an array for the key values to find.
    '            Dim findTheseVals(1) As Object
    '            findTheseVals(0) = oIdentKli
    '            findTheseVals(1) = oOsobaOso
    '            foundRow = DataSetAnalizyZakupu.Tables(0).Rows.Find(findTheseVals)
    '            ' sprawdz czy jest...
    '            If Not (foundRow Is Nothing) Then
    '                foundRow.Delete()
    '                'wyswietl ilosc rekordow
    '                AktualizujAnalizyZakupu()
    '            Else
    '            End If
    '        End If
    '    End If
    'End Sub

    'Private Sub cmdEdytujAnalizyZakupu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytujAnalizyZakupu.Click
    '    If DataSetAnalizyZakupu.Tables(0).Rows.Count <= 0 Then
    '        'Raiseevent(Dodaj,Click )
    '        MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
    '            "Proszê najpierw dopisaæ rekord do tabeli...",
    '            "Uwaga :",
    '            System.Windows.Forms.MessageBoxButtons.OK,
    '            MessageBoxIcon.Information,
    '            MessageBoxDefaultButton.Button1)
    '    Else
    '        oOperacja = "EDYTUJ"
    '        PobierzAnalizyZakupu()
    '        Dim Form As New EdytujKatalogOsob
    '        oAktualizuj = False
    '        Form.ShowDialog()
    '        If Not oAktualizuj Then
    '        Else
    '            Dim foundRow As DataRow
    '            ' Create an array for the key values to find.
    '            Dim findTheseVals(1) As Object
    '            findTheseVals(0) = oIdentKli
    '            findTheseVals(1) = oOsobaOso
    '            foundRow = DataSetAnalizyZakupu.Tables(0).Rows.Find(findTheseVals)
    '            ' sprawdz czy jest...
    '            If Not (foundRow Is Nothing) Then
    '                'foundRow("IDENTKLIENTA") = oIdentKli
    '                foundRow("OSOBA") = oOsobaOso
    '                foundRow("STANOWISKO") = oStanowiskoOso
    '                foundRow("KOMPETENCJE") = oKompetencjeOso
    '                foundRow("TELEFON") = oTelefonOso
    '                foundRow("EMAIL") = oeMailOso
    '                foundRow("VIP") = oVIPOso
    '                foundRow("WERSJA") = ZmienWersje(oWersjaOso) 'Wersja         Integer, 2

    '                'wyswietl ilosc rekordow
    '                AktualizujAnalizyZakupu()
    '                'aktualizuj DataGrid
    '                dagAnalizyZakupu.Update()
    '            End If
    '        End If
    '    End If
    'End Sub


    'Private Sub menuAnalizyZakupuEdytuj_Click(sender As Object, e As EventArgs) Handles menuAnalizyZakupuEdytuj.Click
    '    cmdEdytujAnalizyZakupu_Click(sender, e)
    'End Sub

    'Private Sub menuAnalizyZakupuDodaj_Click(sender As Object, e As EventArgs) Handles menuAnalizyZakupuDodaj.Click
    '    cmdDodajAnalizyZakupu_Click(sender, e)
    'End Sub

    'Private Sub menuAnalizyZakupuUsun_Click(sender As Object, e As EventArgs) Handles menuAnalizyZakupuUsun.Click
    '    cmdUsuñAnalizyZakupu_Click(sender, e)
    'End Sub

    Private Sub menuAnalizyZakupuMultiL_Click(sender As Object, e As EventArgs) Handles menuAnalizyZakupuMultiL.Click
        dagAnalizyZakupu.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagAnalizyZakupu.Refresh()
    End Sub

    Private Sub menuAnalizyZakupuOdswiez_Click(sender As Object, e As EventArgs) Handles menuAnalizyZakupuOdswiez.Click
        OdswiezBazeAnalizyZakupu()
    End Sub

    Private Sub menuAnalizyZakupuSingleL_Click_1(sender As Object, e As EventArgs) Handles menuAnalizyZakupuSingleL.Click
        dagAnalizyZakupu.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagAnalizyZakupu.Refresh()
    End Sub






















    '*********************************************************
    '** Katalog obrotow miesiecznych
    '*********************************************************

    Private Sub dagObrotyMies_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagObrotyMies.CurrentCellChanged
        If Not StartFormyObrotyMies Then
            If dagObrotyMies.CurrentCell Is Nothing Then
                stbObrotyMies.Panels(1).Text = "Wi: "
                stbObrotyMies.Panels(2).Text = "Ko: "
            Else
                stbObrotyMies.Panels(1).Text = "Wi: " & dagObrotyMies.CurrentCell.RowIndex.ToString
                stbObrotyMies.Panels(2).Text = "Ko: " & dagObrotyMies.CurrentCell.ColumnIndex.ToString
            End If
        End If
    End Sub

    '-=------------------------------------------
    'obsluga zegara scrollingu poziomego
    '-=------------------------------------------
    'Private Sub HorizScrollTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HorizScrollTimer.Tick
    '    HorizScrollTimer.Enabled = False
    '    If Len(HorizScrollLastTime) = 0 Then
    '        'nie zainicjowano scrollingu
    '    Else
    '        'zainicjowano scrolling - sprawdz czy juz trwa 2 sec
    '        If IleSekundMinelo(HorizScrollLastTime, SysGodz()) >= 0 Then
    '            If dagObrotyMies.RowCount > 0 Then
    '                If ScrollXOffset <> 10000 * dagObrotyMies.FirstDisplayedScrollingColumnIndex +
    '                                dagObrotyMies.GetCellDisplayRectangle(dagObrotyMies.FirstDisplayedScrollingColumnIndex, 0, True).Left -
    '                                StartPointInCell00.X Then

    '                    PokazFiltryColumnDGV(Me.ObrotyMies, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltryObrotyMies, StartFormyObrotyMies)
    '                    '-----------------------------
    '                    ScrollXOffset = 10000 * dagObrotyMies.FirstDisplayedScrollingColumnIndex +
    '                                dagObrotyMies.GetCellDisplayRectangle(dagObrotyMies.FirstDisplayedScrollingColumnIndex, 0, True).Left -
    '                                StartPointInCell00.X
    '                    HorizScrollLastTime = ""
    '                End If
    '            Else
    '                PokazFiltryColumnDGV(Me.ObrotyMies, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltryObrotyMies, StartFormyObrotyMies)
    '            End If
    '        Else
    '        End If
    '    End If
    '    HorizScrollTimer.Enabled = True
    'End Sub


    'Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
    '    Dim rc As DataGridViewSelectedRowCollection = dagObrotyMies.SelectedRows
    '    If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
    '        If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
    '            Dim nextrow As Integer = rc(0).Index + 1
    '            dagObrotyMies.Rows(nextrow).Selected = True
    '        End If
    '    End If
    'End Sub



    '*********************************************************
    '***** Filtrowanie wg poszczególnych pól
    '*********************************************************

    'AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

    Private Sub txtFiXXObrotyMies_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Not StartFormyObrotyMies Then
            Select Case e.KeyCode
                'Case Keys.Enter
                'Case Keys.Insert, Keys.Add
                'Case Keys.Delete, Keys.Subtract

                Case Keys.Down
                    PrzejdzDoDGV(dagObrotyMies, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXXObrotyMies_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormyObrotyMies Then
            FiltrujDataViewDGV(Me.ObrotyMies, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, DataViewObrotyMies, stbObrotyMies)
        End If
    End Sub
    Private Sub txtFiXXObrotyMies_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXXObrotyMies_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXXObrotyMies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me.ObrotyMies, dagObrotyMies, Mid(sender.name, 6, 3), "ObrotyMies")
    End Sub
    Private Sub cmdFiXXObrotyMies_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiFiltraCmd(Me.ObrotyMies, sender)
    End Sub
    Private Sub cmdFiXXObrotyMies_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraCmd(Me.ObrotyMies, sender)
    End Sub


    '-------------------------------------------------
    ' filtrowanie 
    '-------------------------------------------------

    Private Sub cmdWszystkoObrotyMies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWszystkoObrotyMies.Click
        Me.Cursor = Cursors.WaitCursor
        oLokalnyFiltr = ""
        Try
            StartFormyObrotyMies = True
            CzyscFiltryDGV(Me.ObrotyMies, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, DataViewObrotyMies, stbObrotyMies)
            StartFormyObrotyMies = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbObrotyMies.Panels(0).Text = "Il.zap.: " & DataViewObrotyMies.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagObrotyMies_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagObrotyMies.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagObrotyMies.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagObrotyMies.DefaultCellStyle.Font,
                      b,
                      e.RowBounds.Location.X + 15,
                      e.RowBounds.Location.Y + 4)
        End Using
    End Sub



    '*************************************************
    '*** Odœwiezanie filtrow kolumnowych 
    '*************************************************

    Private Sub stbFiltryObrotyMies_PanelClick(sender As Object, e As StatusBarPanelClickEventArgs) Handles stbFiltryObrotyMies.PanelClick
        If Not StartFormyObrotyMies Then
            PokazFiltryColumnDGV(Me.ObrotyMies, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltryObrotyMies, StartFormyObrotyMies)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagObrotyMies_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagObrotyMies.CellFormatting
    '    If Not StartFormyObrotyMies Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagObrotyMies, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagObrotyMies, e.RowIndex, 4)

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

    Private Sub dagObrotyMies_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagObrotyMies.ColumnWidthChanged
        If Not StartFormyObrotyMies Then
            PokazFiltryColumnDGV(Me.ObrotyMies, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltryObrotyMies, StartFormyObrotyMies)
        End If
    End Sub

    Private Sub dagObrotyMies_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagObrotyMies.Resize
        If Not StartFormyObrotyMies Then
            PokazFiltryColumnDGV(Me.ObrotyMies, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltryObrotyMies, StartFormyObrotyMies)
        End If
    End Sub

    Private Sub dagObrotyMies_Scroll(sender As Object, e As ScrollEventArgs) Handles dagObrotyMies.Scroll
        'If (Not StartFormyObrotyMies) And (DataViewObrotyMies.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me.ObrotyMies, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltryObrotyMies, StartFormyObrotyMies)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me.ObrotyMies, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltryObrotyMies, StartFormyObrotyMies)
        '        End If
        '    End If
        'End If

        If (Not StartFormyObrotyMies) And (DataViewObrotyMies.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagObrotyMies_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagObrotyMies.SizeChanged
        If Not StartFormyObrotyMies Then
            PokazFiltryColumnDGV(Me.ObrotyMies, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltryObrotyMies, StartFormyObrotyMies)
        End If
    End Sub

    Private Sub dagObrotyMies_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagObrotyMies.MouseMove
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None     '"the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                'If Not StartFormyObrotyMies Then
                '    PokazFiltryColumnDGV(Me.ObrotyMies, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltryObrotyMies, StartFormyObrotyMies)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormyObrotyMies Then
                    PokazFiltryColumnDGV(Me.ObrotyMies, dagObrotyMies, DataViewObrotyMies.Table.Columns.Count, stbFiltryObrotyMies, StartFormyObrotyMies)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagObrotyMies_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagObrotyMies.MouseUp
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
                stbObrotyMies.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbObrotyMies.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagObrotyMies(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbObrotyMies.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbObrotyMies.Panels(3).Text = "Sort: " & DataSetObrotyMies.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbObrotyMies.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagObrotyMies(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbObrotyMies.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbObrotyMies.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub






    Private Sub dagObrotyMies_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagObrotyMies.DoubleClick
        cmdEdytujObrotyMies_Click(sender, e)
    End Sub

    Private Sub dagObrotyMies_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case e.KeyCode
            Case Keys.Enter
                cmdEdytujObrotyMies_Click(sender, e)
            Case Keys.Insert
                cmdDodajObrotyMies_Click(sender, e)
            Case Keys.Delete
                cmdUsuñObrotyMies_Click(sender, e)
            Case Else
        End Select
    End Sub










    Private Sub InitObrotyMies()
        oIdentMies = oIdentKli
        'IdentIdKlienta(oIdentMies)
        oMiesiacMies = Mid(SysData(), 1, 7)

        oAIloSprzedMies = 0
        oAWartSprzedMies = 0
        oAMARSprzedMies = 0
        oLIloSprzedMies = 0
        oLWartSprzedMies = 0
        oLMARSprzedMies = 0

        oAOIloSprzedMies = 0
        oAOWartSprzedMies = 0
        oAOMARSprzedMies = 0
        oLOIloSprzedMies = 0
        oLOWartSprzedMies = 0
        oLOMARSprzedMies = 0

        oNAJEMWartSprzedMies = 0
        oNAJEMIloSprzedMies = 0
        oNAJEMMARSprzedMies = 0

        oSTRONYWartSprzedMies = 0
        oSTRONYIloSprzedMies = 0
        oSTRONYMARSprzedMies = 0

        oDRUKARKIWartSprzedMies = 0
        oDRUKARKIIloSprzedMies = 0
        oDRUKARKIMARSprzedMies = 0

        oSKUPWartSprzedMies = 0
        oSKUPIloSprzedMies = 0
        oSKUPMARSprzedMies = 0

        oWersjaMies = 1          'WERSJA         Integer, 2
    End Sub

    Private Sub PobierzObrotyMies()
        'pobierz wartosci przed aktualizacja
        oIdentMies = GetTxtField(dagObrotyMies, 0)
        'IdentIdKlienta(oIdentMies)
        oMiesiacMies = GetTxtField(dagObrotyMies, 2)

        oLMARSprzedMies = GetDblField(dagObrotyMies, 3)
        oLWartSprzedMies = GetDblField(dagObrotyMies, 4)
        oLIloSprzedMies = GetDblField(dagObrotyMies, 5)
        oAMARSprzedMies = GetDblField(dagObrotyMies, 6)
        oAWartSprzedMies = GetDblField(dagObrotyMies, 7)
        oAIloSprzedMies = GetDblField(dagObrotyMies, 8)

        oLOMARSprzedMies = GetDblField(dagObrotyMies, 9)
        oLOWartSprzedMies = GetDblField(dagObrotyMies, 10)
        oLOIloSprzedMies = GetDblField(dagObrotyMies, 11)
        oAOMARSprzedMies = GetDblField(dagObrotyMies, 12)
        oAOWartSprzedMies = GetDblField(dagObrotyMies, 13)
        oAOIloSprzedMies = GetDblField(dagObrotyMies, 14)

        oNAJEMMARSprzedMies = GetDblField(dagObrotyMies, 15)
        oNAJEMWartSprzedMies = GetDblField(dagObrotyMies, 16)
        oNAJEMIloSprzedMies = GetDblField(dagObrotyMies, 17)

        oSTRONYMARSprzedMies = GetDblField(dagObrotyMies, 18)
        oSTRONYWartSprzedMies = GetDblField(dagObrotyMies, 19)
        oSTRONYIloSprzedMies = GetDblField(dagObrotyMies, 20)

        oDRUKARKIMARSprzedMies = GetDblField(dagObrotyMies, 21)
        oDRUKARKIWartSprzedMies = GetDblField(dagObrotyMies, 22)
        oDRUKARKIIloSprzedMies = GetDblField(dagObrotyMies, 23)

        oSKUPMARSprzedMies = GetDblField(dagObrotyMies, 24)
        oSKUPWartSprzedMies = GetDblField(dagObrotyMies, 25)
        oSKUPIloSprzedMies = GetDblField(dagObrotyMies, 26)

        oWersjaMies = GetIntField(dagObrotyMies, 30)
    End Sub




    '***********************************************************
    '*** obsluga komend dla obrotow miesiecznych
    '***********************************************************

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

    Private Sub OdswiezBazeObrotyMies()
        DataSetObrotyMies.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionObrotyMies.Open()
            '    If dbConnectionObrotyMies.State = ConnectionState.Open Then
            '        dbDataAdapterObrotyMies.Fill(DataSetObrotyMies)
            '        dbConnectionObrotyMies.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnectionObrotyMies.Open()
                If sqlConnectionObrotyMies.State = ConnectionState.Open Then
                    sqlDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    sqlConnectionObrotyMies.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub



    Private Sub cmdDodajObrotyMies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodajObrotyMies.Click
        oOperacja = "DODAJ"
        InitObrotyMies()
        '----------------------------------------------
        Do While True
            oAktualizuj = False
            Dim Form As New EdytujKatalogObrotowMies
            Form.ShowDialog()
            Form.Dispose()
            If Not oAktualizuj Then
                Exit Do
            Else
                Dim foundRow As DataRow
                Dim NewRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentKli
                findTheseVals(1) = oMiesiacMies
                foundRow = DataSetObrotyMies.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "Klient = " & oIdentKli & vbCrLf &
                        "Miesi¹c = " & oMiesiacMies,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetObrotyMies.Tables(0).NewRow()

                    NewRow("IDENTKLIENTA") = oIdentKli
                    NewRow("NRTOFITXT") = oNrTOFITxtKli
                    NewRow("MIESIAC") = oMiesiacMies

                    NewRow("LILOSPRZED") = oLIloSprzedMies
                    NewRow("LWARTSPRZED") = oLWartSprzedMies
                    NewRow("LMARSPRZED") = oLMARSprzedMies
                    NewRow("AILOSPRZED") = oAIloSprzedMies
                    NewRow("AWARTSPRZED") = oAWartSprzedMies
                    NewRow("AMARSPRZED") = oAMARSprzedMies

                    NewRow("LOILOSPRZED") = oLOIloSprzedMies
                    NewRow("LOWARTSPRZED") = oLOWartSprzedMies
                    NewRow("LOMARSPRZED") = oLOMARSprzedMies
                    NewRow("AOILOSPRZED") = oAOIloSprzedMies
                    NewRow("AOWARTSPRZED") = oAOWartSprzedMies
                    NewRow("AOMARSPRZED") = oAOMARSprzedMies

                    NewRow("NAJEMILOSPRZED") = oNAJEMIloSprzedMies
                    NewRow("NAJEMWARTSPRZED") = oNAJEMWartSprzedMies
                    NewRow("NAJEMMARSPRZED") = oNAJEMMARSprzedMies

                    NewRow("STRONYILOSPRZED") = oSTRONYIloSprzedMies
                    NewRow("STRONYWARTSPRZED") = oSTRONYWartSprzedMies
                    NewRow("STRONYMARSPRZED") = oSTRONYMARSprzedMies

                    NewRow("DRUKARKIILOSPRZED") = oDRUKARKIIloSprzedMies
                    NewRow("DRUKARKIWARTSPRZED") = oDRUKARKIWartSprzedMies
                    NewRow("DRUKARKIMARSPRZED") = oDRUKARKIMARSprzedMies

                    NewRow("NAJEMILOSPRZED") = oNAJEMIloSprzedMies
                    NewRow("NAJEMWARTSPRZED") = oNAJEMWartSprzedMies
                    NewRow("NAJEMMARSPRZED") = oNAJEMMARSprzedMies

                    NewRow("SUMAILO") = oLIloSprzedMies + oAIloSprzedMies +
                                          oLOIloSprzedMies + oAOIloSprzedMies +
                                          oNAJEMIloSprzedMies + oSTRONYIloSprzedMies +
                                          oDRUKARKIIloSprzedMies + oSKUPIloSprzedMies
                    NewRow("SUMAWART") = oLWartSprzedMies + oAWartSprzedMies +
                                          oLOWartSprzedMies + oAOWartSprzedMies +
                                          oNAJEMWartSprzedMies + oSTRONYWartSprzedMies +
                                          oDRUKARKIWartSprzedMies + oSKUPWartSprzedMies
                    NewRow("SUMAMAR") = oLMARSprzedMies + oAMARSprzedMies +
                                          oLOMARSprzedMies + oAOMARSprzedMies +
                                          oNAJEMMARSprzedMies + oSTRONYMARSprzedMies +
                                          oDRUKARKIMARSprzedMies + oSKUPMARSprzedMies

                    NewRow("WERSJA") = 1     'Wersja         Integer, 2

                    'dodaj rekord do DataSet
                    DataSetObrotyMies.Tables(0).Rows.Add(NewRow)
                    'wyswietl ilosc rekordow
                    AktualizujObrotyMies()
                    'aktualizuj DataGrid
                    dagObrotyMies.Update()
                    Exit Do
                End If
            End If
        Loop
    End Sub

    Private Sub cmdUsuñObrotyMies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñObrotyMies.Click
        If DataSetObrotyMies.Tables(0).Rows.Count <= 0 Then
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
                PobierzObrotyMies()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentKli
                findTheseVals(1) = oMiesiacMies
                foundRow = DataSetObrotyMies.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    AktualizujObrotyMies()
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdEdytujObrotyMies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytujObrotyMies.Click
        If DataSetObrotyMies.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzObrotyMies()
            Dim Form As New EdytujKatalogObrotowMies
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentKli
                findTheseVals(1) = oMiesiacMies
                foundRow = DataSetObrotyMies.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'foundRow("IDENTKLIENTA") = oIdentKli
                    foundRow("NRTOFITXT") = oNrTOFITxtKli
                    'foundRow("MIESIAC") = oMiesiacMies

                    foundRow("LILOSPRZED") = oLIloSprzedMies
                    foundRow("LWARTSPRZED") = oLWartSprzedMies
                    foundRow("LMARSPRZED") = oLMARSprzedMies
                    foundRow("AILOSPRZED") = oAIloSprzedMies
                    foundRow("AWARTSPRZED") = oAWartSprzedMies
                    foundRow("AMARSPRZED") = oAMARSprzedMies

                    foundRow("LOILOSPRZED") = oLOIloSprzedMies
                    foundRow("LOWARTSPRZED") = oLOWartSprzedMies
                    foundRow("LOMARSPRZED") = oLOMARSprzedMies
                    foundRow("AOILOSPRZED") = oAOIloSprzedMies
                    foundRow("AOWARTSPRZED") = oAOWartSprzedMies
                    foundRow("AOMARSPRZED") = oAOMARSprzedMies

                    foundRow("NAJEMILOSPRZED") = oNAJEMIloSprzedMies
                    foundRow("NAJEMWARTSPRZED") = oNAJEMWartSprzedMies
                    foundRow("NAJEMMARSPRZED") = oNAJEMMARSprzedMies

                    foundRow("STRONYILOSPRZED") = oSTRONYIloSprzedMies
                    foundRow("STRONYWARTSPRZED") = oSTRONYWartSprzedMies
                    foundRow("STRONYMARSPRZED") = oSTRONYMARSprzedMies

                    foundRow("DRUKARKIILOSPRZED") = oDRUKARKIIloSprzedMies
                    foundRow("DRUKARKIWARTSPRZED") = oDRUKARKIWartSprzedMies
                    foundRow("DRUKARKIMARSPRZED") = oDRUKARKIMARSprzedMies

                    foundRow("NAJEMILOSPRZED") = oNAJEMIloSprzedMies
                    foundRow("NAJEMWARTSPRZED") = oNAJEMWartSprzedMies
                    foundRow("NAJEMMARSPRZED") = oNAJEMMARSprzedMies

                    'foundRow("SUMAILO") = oLIloSprzedMies + oAIloSprzedMies +
                    '                      oLOIloSprzedMies + oAOIloSprzedMies +
                    '                      oNAJEMIloSprzedMies + oSTRONYIloSprzedMies +
                    '                      oDRUKARKIIloSprzedMies + oSKUPIloSprzedMies
                    'foundRow("SUMAWART") = oLWartSprzedMies + oAWartSprzedMies +
                    '                      oLOWartSprzedMies + oAOWartSprzedMies +
                    '                      oNAJEMWartSprzedMies + oSTRONYWartSprzedMies +
                    '                      oDRUKARKIWartSprzedMies + oSKUPWartSprzedMies
                    'foundRow("SUMAMAR") = oLMARSprzedMies + oAMARSprzedMies +
                    '                      oLOMARSprzedMies + oAOMARSprzedMies +
                    '                      oNAJEMMARSprzedMies + oSTRONYMARSprzedMies +
                    '                      oDRUKARKIMARSprzedMies + oSKUPMARSprzedMies

                    foundRow("WERSJA") = ZmienWersje(oWersjaMies) 'Wersja         Integer, 2

                    'wyswietl ilosc rekordow
                    AktualizujObrotyMies()
                    'aktualizuj DataGrid
                    dagObrotyMies.Update()
                End If
            End If
        End If
    End Sub





    Private Sub menuObrotyMiesEdytuj_Click_1(sender As Object, e As EventArgs) Handles menuObrotyMiesEdytuj.Click
        cmdEdytujObrotyMies_Click(sender, e)
    End Sub

    Private Sub menuObrotyMiesDodaj_Click(sender As Object, e As EventArgs) Handles menuObrotyMiesDodaj.Click
        cmdDodajObrotyMies_Click(sender, e)
    End Sub

    Private Sub menuObrotyMiesUsun_Click(sender As Object, e As EventArgs) Handles menuObrotyMiesUsun.Click
        cmdUsuñObrotyMies_Click(sender, e)
    End Sub

    Private Sub MenuObrotyMiesSingleL_Click(sender As Object, e As EventArgs) Handles menuObrotyMiesSingleL.Click
        dagObrotyMies.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagObrotyMies.Refresh()
    End Sub

    Private Sub menuObrotyMiesMultiL_Click(sender As Object, e As EventArgs) Handles menuObrotyMiesMultiL.Click
        dagObrotyMies.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagObrotyMies.Refresh()
    End Sub

    Private Sub menuObrotyMiesOdswiez_Click(sender As Object, e As EventArgs) Handles menuObrotyMiesOdswiez.Click
        OdswiezBazeObrotyMies()
    End Sub























    '*********************************************************
    '** Katalog obrotow
    '*********************************************************

    Private Sub dagObroty_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagObroty.CurrentCellChanged
        If Not StartFormyObroty Then
            If dagObroty.CurrentCell Is Nothing Then
                stbObroty.Panels(1).Text = "Wi: "
                stbObroty.Panels(2).Text = "Ko: "
            Else
                stbObroty.Panels(1).Text = "Wi: " & dagObroty.CurrentCell.RowIndex.ToString
                stbObroty.Panels(2).Text = "Ko: " & dagObroty.CurrentCell.ColumnIndex.ToString
            End If
        End If
    End Sub

    '-=------------------------------------------
    'obsluga zegara scrollingu poziomego
    '-=------------------------------------------
    'Private Sub HorizScrollTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HorizScrollTimer.Tick
    '    HorizScrollTimer.Enabled = False
    '    If Len(HorizScrollLastTime) = 0 Then
    '        'nie zainicjowano scrollingu
    '    Else
    '        'zainicjowano scrolling - sprawdz czy juz trwa 2 sec
    '        If IleSekundMinelo(HorizScrollLastTime, SysGodz()) >= 0 Then
    '            If dagObroty.RowCount > 0 Then
    '                If ScrollXOffset <> 10000 * dagObroty.FirstDisplayedScrollingColumnIndex +
    '                                dagObroty.GetCellDisplayRectangle(dagObroty.FirstDisplayedScrollingColumnIndex, 0, True).Left -
    '                                StartPointInCell00.X Then

    '                    PokazFiltryColumnDGV(Me.Obroty, dagObroty, DataViewObroty.Table.Columns.Count, stbFiltryObroty, StartFormyObroty)
    '                    '-----------------------------
    '                    ScrollXOffset = 10000 * dagObroty.FirstDisplayedScrollingColumnIndex +
    '                                dagObroty.GetCellDisplayRectangle(dagObroty.FirstDisplayedScrollingColumnIndex, 0, True).Left -
    '                                StartPointInCell00.X
    '                    HorizScrollLastTime = ""
    '                End If
    '            Else
    '                PokazFiltryColumnDGV(Me.Obroty, dagObroty, DataViewObroty.Table.Columns.Count, stbFiltryObroty, StartFormyObroty)
    '            End If
    '        Else
    '        End If
    '    End If
    '    HorizScrollTimer.Enabled = True
    'End Sub


    'Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
    '    Dim rc As DataGridViewSelectedRowCollection = dagObroty.SelectedRows
    '    If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
    '        If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
    '            Dim nextrow As Integer = rc(0).Index + 1
    '            dagObroty.Rows(nextrow).Selected = True
    '        End If
    '    End If
    'End Sub



    '*********************************************************
    '***** Filtrowanie wg poszczególnych pól
    '*********************************************************

    'AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

    Private Sub txtFiXXObroty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Not StartFormyObroty Then
            Select Case e.KeyCode
                'Case Keys.Enter
                'Case Keys.Insert, Keys.Add
                'Case Keys.Delete, Keys.Subtract

                Case Keys.Down
                    PrzejdzDoDGV(dagObroty, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXXObroty_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormyObroty Then
            FiltrujDataViewDGV(Me.Obroty, dagObroty, DataViewObroty.Table.Columns.Count, DataViewObroty, stbObroty)
        End If
    End Sub
    Private Sub txtFiXXObroty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXXObroty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXXObroty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me.Obroty, dagObroty, Mid(sender.name, 6, 3), "Obroty")
    End Sub
    Private Sub cmdFiXXObroty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiFiltraCmd(Me.Obroty, sender)
    End Sub
    Private Sub cmdFiXXObroty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraCmd(Me.Obroty, sender)
    End Sub


    '-------------------------------------------------
    ' filtrowanie 
    '-------------------------------------------------

    Private Sub cmdWszystkoObroty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWszystkoObroty.Click
        Me.Cursor = Cursors.WaitCursor
        oLokalnyFiltr = ""
        Try
            StartFormyObroty = True
            CzyscFiltryDGV(Me.Obroty, dagObroty, DataViewObroty.Table.Columns.Count, DataViewObroty, stbObroty)
            StartFormyObroty = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbObroty.Panels(0).Text = "Il.zap.: " & DataViewObroty.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagObroty_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagObroty.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagObroty.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagObroty.DefaultCellStyle.Font,
                      b,
                      e.RowBounds.Location.X + 15,
                      e.RowBounds.Location.Y + 4)
        End Using
    End Sub



    '*************************************************
    '*** Odœwiezanie filtrow kolumnowych 
    '*************************************************

    Private Sub stbFiltryObroty_PanelClick(sender As Object, e As StatusBarPanelClickEventArgs) Handles stbFiltryObroty.PanelClick
        If Not StartFormyObroty Then
            PokazFiltryColumnDGV(Me.Obroty, dagObroty, DataViewObroty.Table.Columns.Count, stbFiltryObroty, StartFormyObroty)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagObroty_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagObroty.CellFormatting
    '    If Not StartFormyObroty Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagObroty, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagObroty, e.RowIndex, 4)

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

    Private Sub dagObroty_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagObroty.ColumnWidthChanged
        If Not StartFormyObroty Then
            PokazFiltryColumnDGV(Me.Obroty, dagObroty, DataViewObroty.Table.Columns.Count, stbFiltryObroty, StartFormyObroty)
        End If
    End Sub

    Private Sub dagObroty_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagObroty.Resize
        If Not StartFormyObroty Then
            PokazFiltryColumnDGV(Me.Obroty, dagObroty, DataViewObroty.Table.Columns.Count, stbFiltryObroty, StartFormyObroty)
        End If
    End Sub

    Private Sub dagObroty_Scroll(sender As Object, e As ScrollEventArgs) Handles dagObroty.Scroll
        'If (Not StartFormyObroty) And (DataViewObroty.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me.Obroty, dagObroty, DataViewObroty.Table.Columns.Count, stbFiltryObroty, StartFormyObroty)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me.Obroty, dagObroty, DataViewObroty.Table.Columns.Count, stbFiltryObroty, StartFormyObroty)
        '        End If
        '    End If
        'End If

        If (Not StartFormyObroty) And (DataViewObroty.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagObroty_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagObroty.SizeChanged
        If Not StartFormyObroty Then
            PokazFiltryColumnDGV(Me.Obroty, dagObroty, DataViewObroty.Table.Columns.Count, stbFiltryObroty, StartFormyObroty)
        End If
    End Sub

    Private Sub dagObroty_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagObroty.MouseMove
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None     '"the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                'If Not StartFormyObroty Then
                '    PokazFiltryColumnDGV(Me.Obroty, dagObroty, DataViewObroty.Table.Columns.Count, stbFiltryObroty, StartFormyObroty)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormyObroty Then
                    PokazFiltryColumnDGV(Me.Obroty, dagObroty, DataViewObroty.Table.Columns.Count, stbFiltryObroty, StartFormyObroty)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagObroty_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagObroty.MouseUp
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
                stbObroty.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbObroty.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagObroty(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbObroty.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbObroty.Panels(3).Text = "Sort: " & DataSetObroty.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbObroty.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagObroty(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbObroty.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbObroty.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub






    Private Sub dagObroty_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagObroty.DoubleClick
        cmdEdytujObroty_Click(sender, e)
    End Sub

    Private Sub dagObroty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case e.KeyCode
            Case Keys.Enter
                cmdEdytujObroty_Click(sender, e)
            Case Keys.Insert
                cmdDodajObroty_Click(sender, e)
            Case Keys.Delete
                cmdusuñObroty_Click(sender, e)
            Case Else
        End Select
    End Sub








    Private Sub InitObroty()
        oIdentObr = oIdentKli
        'IdentIdKlienta(oIdentObr)   'pobierz NrTOFITxt
        oDataObr = SysData()

        oAIloSprzedObr = 0
        oAWartSprzedObr = 0
        oAMARSprzedObr = 0
        oLIloSprzedObr = 0
        oLWartSprzedObr = 0
        oLMarSprzedObr = 0

        oAOIloSprzedObr = 0
        oAOWartSprzedObr = 0
        oAOMARSprzedObr = 0
        oLOIloSprzedObr = 0
        oLOWartSprzedObr = 0
        oLOMARSprzedObr = 0

        oNAJEMWartSprzedObr = 0
        oNAJEMIloSprzedObr = 0
        oNAJEMMARSprzedObr = 0

        oSTRONYWartSprzedObr = 0
        oSTRONYIloSprzedObr = 0
        oSTRONYMARSprzedObr = 0

        oDRUKARKIWartSprzedObr = 0
        oDRUKARKIIloSprzedObr = 0
        oDRUKARKIMARSprzedObr = 0

        oSKUPWartSprzedObr = 0
        oSKUPIloSprzedObr = 0
        oSKUPMARSprzedObr = 0

        oWersjaObr = 1          'WERSJA         Integer, 2
    End Sub


    Private Sub PobierzObroty()
        'pobierz wartosci przed aktualizacja
        oIdentObr = GetTxtField(dagObroty, 0)
        'IdentIdKlienta(oIdentobr)
        oDataObr = GetTxtField(dagObroty, 2)

        oLMarSprzedObr = GetDblField(dagObroty, 3)
        oLWartSprzedObr = GetDblField(dagObroty, 4)
        oLIloSprzedObr = GetDblField(dagObroty, 5)
        oAMARSprzedObr = GetDblField(dagObroty, 6)
        oAWartSprzedObr = GetDblField(dagObroty, 7)
        oAIloSprzedObr = GetDblField(dagObroty, 8)

        oLOMARSprzedObr = GetDblField(dagObroty, 9)
        oLOWartSprzedObr = GetDblField(dagObroty, 10)
        oLOIloSprzedObr = GetDblField(dagObroty, 11)
        oAOMARSprzedObr = GetDblField(dagObroty, 12)
        oAOWartSprzedObr = GetDblField(dagObroty, 13)
        oAOIloSprzedObr = GetDblField(dagObroty, 14)

        oNAJEMMARSprzedObr = GetDblField(dagObroty, 15)
        oNAJEMWartSprzedObr = GetDblField(dagObroty, 16)
        oNAJEMIloSprzedObr = GetDblField(dagObroty, 17)

        oSTRONYMARSprzedObr = GetDblField(dagObroty, 18)
        oSTRONYWartSprzedObr = GetDblField(dagObroty, 19)
        oSTRONYIloSprzedObr = GetDblField(dagObroty, 20)

        oDRUKARKIMARSprzedObr = GetDblField(dagObroty, 21)
        oDRUKARKIWartSprzedObr = GetDblField(dagObroty, 22)
        oDRUKARKIIloSprzedObr = GetDblField(dagObroty, 23)

        oSKUPMARSprzedObr = GetDblField(dagObroty, 24)
        oSKUPWartSprzedObr = GetDblField(dagObroty, 25)
        oSKUPIloSprzedObr = GetDblField(dagObroty, 26)


        oWersjaObr = GetIntField(dagObroty, 30)
    End Sub




    '***********************************************************
    '*** obsluga komend dla obrotow miesiecznych
    '***********************************************************

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

    Private Sub OdswiezBazeObroty()
        DataSetObroty.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionObroty.Open()
            '    If dbConnectionObroty.State = ConnectionState.Open Then
            '        dbDataAdapterObroty.Fill(DataSetObroty)
            '        dbConnectionObroty.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnectionObroty.Open()
                If sqlConnectionObroty.State = ConnectionState.Open Then
                    sqlDataAdapterObroty.Fill(DataSetObroty)
                    sqlConnectionObroty.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub







    Private Sub cmdDodajObroty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodajObroty.Click
        oOperacja = "DODAJ"
        InitObroty()
        '----------------------------------------------
        Do While True
            oAktualizuj = False
            Dim Form As New EdytujKatalogObrotow
            Form.ShowDialog()
            Form.Dispose()
            If Not oAktualizuj Then
                Exit Do
            Else
                Dim foundRow As DataRow
                Dim NewRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentKli
                findTheseVals(1) = oDataObr
                foundRow = DataSetObroty.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
                        "Klient = " & oIdentKli & vbCrLf &
                        "Data = " & oDataObr,
                        "Uwaga :",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    'stworz nowy rekord
                    NewRow = DataSetObroty.Tables(0).NewRow()

                    NewRow("IDENTKLIENTA") = oIdentKli
                    NewRow("NRTOFITXT") = oNrTOFITxtKli
                    NewRow("DATA") = oDataObr

                    NewRow("LILOSPRZED") = oLIloSprzedObr
                    NewRow("LWARTSPRZED") = oLWartSprzedObr
                    NewRow("LMARSPRZED") = oLMarSprzedObr
                    NewRow("AILOSPRZED") = oAIloSprzedObr
                    NewRow("AWARTSPRZED") = oAWartSprzedObr
                    NewRow("AMARSPRZED") = oAMARSprzedObr

                    NewRow("LOILOSPRZED") = oLOIloSprzedObr
                    NewRow("LOWARTSPRZED") = oLOWartSprzedObr
                    NewRow("LOMARSPRZED") = oLOMARSprzedObr
                    NewRow("AOILOSPRZED") = oAOIloSprzedObr
                    NewRow("AOWARTSPRZED") = oAOWartSprzedObr
                    NewRow("AOMARSPRZED") = oAOMARSprzedObr

                    NewRow("NAJEMILOSPRZED") = oNAJEMIloSprzedObr
                    NewRow("NAJEMWARTSPRZED") = oNAJEMWartSprzedObr
                    NewRow("NAJEMMARSPRZED") = oNAJEMMARSprzedObr

                    NewRow("STRONYILOSPRZED") = oSTRONYIloSprzedObr
                    NewRow("STRONYWARTSPRZED") = oSTRONYWartSprzedObr
                    NewRow("STRONYMARSPRZED") = oSTRONYMARSprzedObr

                    NewRow("DRUKARKIILOSPRZED") = oDRUKARKIIloSprzedObr
                    NewRow("DRUKARKIWARTSPRZED") = oDRUKARKIWartSprzedObr
                    NewRow("DRUKARKIMARSPRZED") = oDRUKARKIMARSprzedObr

                    NewRow("NAJEMILOSPRZED") = oNAJEMIloSprzedObr
                    NewRow("NAJEMWARTSPRZED") = oNAJEMWartSprzedObr
                    NewRow("NAJEMMARSPRZED") = oNAJEMMARSprzedObr

                    NewRow("SUMAILO") = oLIloSprzedObr + oAIloSprzedObr +
                                          oLOIloSprzedObr + oAOIloSprzedObr +
                                          oNAJEMIloSprzedObr + oSTRONYIloSprzedObr +
                                          oDRUKARKIIloSprzedObr + oSKUPIloSprzedObr
                    NewRow("SUMAWART") = oLWartSprzedObr + oAWartSprzedObr +
                                          oLOWartSprzedObr + oAOWartSprzedObr +
                                          oNAJEMWartSprzedObr + oSTRONYWartSprzedObr +
                                          oDRUKARKIWartSprzedObr + oSKUPWartSprzedObr
                    NewRow("SUMAMAR") = oLMarSprzedObr + oAMARSprzedObr +
                                          oLOMARSprzedObr + oAOMARSprzedObr +
                                          oNAJEMMARSprzedObr + oSTRONYMARSprzedObr +
                                          oDRUKARKIMARSprzedObr + oSKUPMARSprzedObr

                    NewRow("WERSJA") = 1     'Wersja         Integer, 2

                    'dodaj rekord do DataSet
                    DataSetObroty.Tables(0).Rows.Add(NewRow)
                    'wyswietl ilosc rekordow
                    AktualizujObroty()
                    'aktualizuj DataGrid
                    dagObroty.Update()
                    Exit Do
                End If
            End If
        Loop
    End Sub

    Private Sub cmdusuñObroty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdusuñObroty.Click
        If DataSetObroty.Tables(0).Rows.Count <= 0 Then
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
                PobierzObroty()
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentKli
                findTheseVals(1) = oDataObr
                foundRow = DataSetObroty.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    foundRow.Delete()
                    'wyswietl ilosc rekordow
                    AktualizujObroty()
                Else
                End If
            End If
        End If
    End Sub

    Private Sub cmdEdytujObroty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytujObroty.Click
        If DataSetObroty.Tables(0).Rows.Count <= 0 Then
            'Raiseevent(Dodaj,Click )
            MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
                "Proszê najpierw dopisaæ rekord do tabeli...",
                "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            oOperacja = "EDYTUJ"
            PobierzObroty()
            Dim Form As New EdytujKatalogObrotow
            oAktualizuj = False
            Form.ShowDialog()
            If Not oAktualizuj Then
            Else
                Dim foundRow As DataRow
                ' Create an array for the key values to find.
                Dim findTheseVals(1) As Object
                findTheseVals(0) = oIdentKli
                findTheseVals(1) = oDataObr
                foundRow = DataSetObroty.Tables(0).Rows.Find(findTheseVals)
                ' sprawdz czy jest...
                If Not (foundRow Is Nothing) Then
                    'foundRow("IDENTKLIENTA") = oIdentKli
                    foundRow("NRTOFITXT") = oNrTOFITxtKli
                    'foundRow("DATA") = oDataObr

                    foundRow("LILOSPRZED") = oLIloSprzedObr
                    foundRow("LWARTSPRZED") = oLWartSprzedObr
                    foundRow("LMARSPRZED") = oLMarSprzedObr
                    foundRow("AILOSPRZED") = oAIloSprzedObr
                    foundRow("AWARTSPRZED") = oAWartSprzedObr
                    foundRow("AMARSPRZED") = oAMARSprzedObr

                    foundRow("LOILOSPRZED") = oLOIloSprzedObr
                    foundRow("LOWARTSPRZED") = oLOWartSprzedObr
                    foundRow("LOMARSPRZED") = oLOMARSprzedObr
                    foundRow("AOILOSPRZED") = oAOIloSprzedObr
                    foundRow("AOWARTSPRZED") = oAOWartSprzedObr
                    foundRow("AOMARSPRZED") = oAOMARSprzedObr

                    foundRow("NAJEMILOSPRZED") = oNAJEMIloSprzedObr
                    foundRow("NAJEMWARTSPRZED") = oNAJEMWartSprzedObr
                    foundRow("NAJEMMARSPRZED") = oNAJEMMARSprzedObr

                    foundRow("STRONYILOSPRZED") = oSTRONYIloSprzedObr
                    foundRow("STRONYWARTSPRZED") = oSTRONYWartSprzedObr
                    foundRow("STRONYMARSPRZED") = oSTRONYMARSprzedObr

                    foundRow("DRUKARKIILOSPRZED") = oDRUKARKIIloSprzedObr
                    foundRow("DRUKARKIWARTSPRZED") = oDRUKARKIWartSprzedObr
                    foundRow("DRUKARKIMARSPRZED") = oDRUKARKIMARSprzedObr

                    foundRow("NAJEMILOSPRZED") = oNAJEMIloSprzedObr
                    foundRow("NAJEMWARTSPRZED") = oNAJEMWartSprzedObr
                    foundRow("NAJEMMARSPRZED") = oNAJEMMARSprzedObr

                    'foundRow("SUMAILO") = oLIloSprzedObr + oAIloSprzedObr +
                    '                      oLOIloSprzedObr + oAOIloSprzedObr +
                    '                      oNAJEMIloSprzedObr + oSTRONYIloSprzedObr +
                    '                      oDRUKARKIIloSprzedObr + oSKUPIloSprzedObr
                    'foundRow("SUMAWART") = oLWartSprzedObr + oAWartSprzedObr +
                    '                      oLOWartSprzedObr + oAOWartSprzedObr +
                    '                      oNAJEMWartSprzedObr + oSTRONYWartSprzedObr +
                    '                      oDRUKARKIWartSprzedObr + oSKUPWartSprzedObr
                    'foundRow("SUMAMAR") = oLMarSprzedObr + oAMARSprzedObr +
                    '                      oLOMARSprzedObr + oAOMARSprzedObr +
                    '                      oNAJEMMARSprzedObr + oSTRONYMARSprzedObr +
                    '                      oDRUKARKIMARSprzedObr + oSKUPMARSprzedObr

                    foundRow("WERSJA") = ZmienWersje(oWersjaObr) 'Wersja         Integer, 2

                    'wyswietl ilosc rekordow
                    AktualizujObroty()
                    'aktualizuj DataGrid
                    dagObroty.Update()
                End If
            End If
        End If
    End Sub







    Private Sub menuObrotyEdytuj_Click(sender As Object, e As EventArgs) Handles menuObrotyEdytuj.Click
        cmdEdytujObroty_Click(sender, e)
    End Sub

    Private Sub menuObrotyDodaj_Click(sender As Object, e As EventArgs) Handles menuObrotyDodaj.Click
        cmdDodajObroty_Click(sender, e)
    End Sub

    Private Sub menuObrotyUsun_Click(sender As Object, e As EventArgs) Handles menuObrotyUsun.Click
        cmdusuñObroty_Click(sender, e)
    End Sub

    Private Sub MenuObrotySingleL_Click(sender As Object, e As EventArgs) Handles menuObrotySingleL.Click
        dagObroty.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagObroty.Refresh()
    End Sub

    Private Sub menuObrotyMultiL_Click(sender As Object, e As EventArgs) Handles menuObrotyMultiL.Click
        dagObroty.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagObroty.Refresh()
    End Sub

    Private Sub menuObrotyOdswiez_Click(sender As Object, e As EventArgs) Handles menuObrotyOdswiez.Click
        OdswiezBazeObroty()
    End Sub













    '*********************************************************
    '** Katalog AkcjeSpec
    '*********************************************************

    Private Sub dagAkcjeSpec_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAkcjeSpec.CurrentCellChanged
        If Not StartFormyAkcjeSpec Then
            If dagAkcjeSpec.CurrentCell Is Nothing Then
                stbAkcjeSpec.Panels(1).Text = "Wi: "
                stbAkcjeSpec.Panels(2).Text = "Ko: "
            Else
                stbAkcjeSpec.Panels(1).Text = "Wi: " & dagAkcjeSpec.CurrentCell.RowIndex.ToString
                stbAkcjeSpec.Panels(2).Text = "Ko: " & dagAkcjeSpec.CurrentCell.ColumnIndex.ToString
            End If
        End If
    End Sub

    '-=------------------------------------------
    'obsluga zegara scrollingu poziomego
    '-=------------------------------------------
    'Private Sub HorizScrollTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HorizScrollTimer.Tick
    '    HorizScrollTimer.Enabled = False
    '    If Len(HorizScrollLastTime) = 0 Then
    '        'nie zainicjowano scrollingu
    '    Else
    '        'zainicjowano scrolling - sprawdz czy juz trwa 2 sec
    '        If IleSekundMinelo(HorizScrollLastTime, SysGodz()) >= 0 Then
    '            If dagAkcjeSpec.RowCount > 0 Then
    '                If ScrollXOffset <> 10000 * dagAkcjeSpec.FirstDisplayedScrollingColumnIndex +
    '                                dagAkcjeSpec.GetCellDisplayRectangle(dagAkcjeSpec.FirstDisplayedScrollingColumnIndex, 0, True).Left -
    '                                StartPointInCell00.X Then

    '                    PokazFiltryColumnDGV(Me.AkcjeSpec, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltryAkcjeSpec, StartFormyAkcjeSpec)
    '                    '-----------------------------
    '                    ScrollXOffset = 10000 * dagAkcjeSpec.FirstDisplayedScrollingColumnIndex +
    '                                dagAkcjeSpec.GetCellDisplayRectangle(dagAkcjeSpec.FirstDisplayedScrollingColumnIndex, 0, True).Left -
    '                                StartPointInCell00.X
    '                    HorizScrollLastTime = ""
    '                End If
    '            Else
    '                PokazFiltryColumnDGV(Me.AkcjeSpec, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltryAkcjeSpec, StartFormyAkcjeSpec)
    '            End If
    '        Else
    '        End If
    '    End If
    '    HorizScrollTimer.Enabled = True
    'End Sub


    'Private Sub frmKatalogUzytkownikow_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
    '    Dim rc As DataGridViewSelectedRowCollection = dagAkcjeSpec.SelectedRows
    '    If e.ScrollOrientation = ScrollOrientation.VerticalScroll Then
    '        If e.NewValue > e.OldValue AndAlso rc.Count > 0 Then
    '            Dim nextrow As Integer = rc(0).Index + 1
    '            dagAkcjeSpec.Rows(nextrow).Selected = True
    '        End If
    '    End If
    'End Sub



    '*********************************************************
    '***** Filtrowanie wg poszczególnych pól
    '*********************************************************

    'AddHandler CtrlT.KeyDown, New KeyEventHandler(AddressOf txtFiXX_KeyDown)

    Private Sub txtFiXXAkcjeSpec_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Not StartFormyAkcjeSpec Then
            Select Case e.KeyCode
                'Case Keys.Enter
                'Case Keys.Insert, Keys.Add
                'Case Keys.Delete, Keys.Subtract

                Case Keys.Down
                    PrzejdzDoDGV(dagAkcjeSpec, Mid(sender.name, 6, 3))

                Case Else
            End Select
        End If
    End Sub



    Private Sub txtFiXXAkcjeSpec_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not StartFormyAkcjeSpec Then
            FiltrujDataViewDGV(Me.AkcjeSpec, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, DataViewAkcjeSpec, stbAkcjeSpec)
        End If
    End Sub
    Private Sub txtFiXXAkcjeSpec_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiTxt(sender)
    End Sub
    Private Sub txtFiXXAkcjeSpec_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraTxt(sender)
    End Sub
    Private Sub cmdFiXXAkcjeSpec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PobierzWartoscFiltraDGV(Me.AkcjeSpec, dagAkcjeSpec, Mid(sender.name, 6, 3), "AkcjeSpec")
    End Sub
    Private Sub cmdFiXXAkcjeSpec_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        StartEdycjiFiltraCmd(Me.AkcjeSpec, sender)
    End Sub
    Private Sub cmdFiXXAkcjeSpec_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        KoniecEdycjiFiltraCmd(Me.AkcjeSpec, sender)
    End Sub


    '-------------------------------------------------
    ' filtrowanie 
    '-------------------------------------------------

    Private Sub cmdWszystkoAkcjeSpec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWszystkoAkcjeSpec.Click
        Me.Cursor = Cursors.WaitCursor
        oLokalnyFiltr = ""
        Try
            StartFormyAkcjeSpec = True
            CzyscFiltryDGV(Me.AkcjeSpec, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, DataViewAkcjeSpec, stbAkcjeSpec)
            StartFormyAkcjeSpec = False

        Catch Ex As System.Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Finally
            stbAkcjeSpec.Panels(0).Text = "Il.zap.: " & DataViewAkcjeSpec.Count.ToString
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    '*************************************************
    '*** Numerowanie wierszy w DataGridView
    '*************************************************

    Private Sub dagAkcjeSpec_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dagAkcjeSpec.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dagAkcjeSpec.RowHeadersDefaultCellStyle.ForeColor)
            e.Graphics.DrawString((e.RowIndex + 1).ToString(System.Globalization.CultureInfo.CurrentUICulture),
                      dagAkcjeSpec.DefaultCellStyle.Font,
                      b,
                      e.RowBounds.Location.X + 15,
                      e.RowBounds.Location.Y + 4)
        End Using
    End Sub



    '*************************************************
    '*** Odœwiezanie filtrow kolumnowych 
    '*************************************************

    Private Sub stbFiltryAkcjeSpec_PanelClick(sender As Object, e As StatusBarPanelClickEventArgs) Handles stbFiltryAkcjeSpec.PanelClick
        If Not StartFormyAkcjeSpec Then
            PokazFiltryColumnDGV(Me.AkcjeSpec, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltryAkcjeSpec, StartFormyAkcjeSpec)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*************************************************
    'kolorowanie komorek w tabeli
    '*************************************************

    'Private Sub dagAkcjeSpec_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dagAkcjeSpec.CellFormatting
    '    If Not StartFormyAkcjeSpec Then
    '        '-----------------------
    '        'zmieniamy ForeColor
    '        Dim UserName As String = GetTxtField(dagAkcjeSpec, e.RowIndex, 0)
    '        Dim Upr As String = GetTxtField(dagAkcjeSpec, e.RowIndex, 4)

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

    Private Sub dagAkcjeSpec_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles dagAkcjeSpec.ColumnWidthChanged
        If Not StartFormyAkcjeSpec Then
            PokazFiltryColumnDGV(Me.AkcjeSpec, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltryAkcjeSpec, StartFormyAkcjeSpec)
        End If
    End Sub

    Private Sub dagAkcjeSpec_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAkcjeSpec.Resize
        If Not StartFormyAkcjeSpec Then
            PokazFiltryColumnDGV(Me.AkcjeSpec, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltryAkcjeSpec, StartFormyAkcjeSpec)
        End If
    End Sub

    Private Sub dagAkcjeSpec_Scroll(sender As Object, e As ScrollEventArgs) Handles dagAkcjeSpec.Scroll
        'If (Not StartFormyAkcjeSpec) And (DataViewAkcjeSpec.Count > 0) Then
        '    If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
        '        If e.NewValue > e.OldValue Then
        '            'w prawo
        '            PokazFiltryColumnDGV(Me.AkcjeSpec, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltryAkcjeSpec, StartFormyAkcjeSpec)
        '        Else
        '            'w lewo
        '            PokazFiltryColumnDGV(Me.AkcjeSpec, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltryAkcjeSpec, StartFormyAkcjeSpec)
        '        End If
        '    End If
        'End If

        If (Not StartFormyAkcjeSpec) And (DataViewAkcjeSpec.Count > 0) Then
            If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
                HorizScrollLastTime = SysGodz()
            End If
        End If

    End Sub

    Private Sub dagAkcjeSpec_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAkcjeSpec.SizeChanged
        If Not StartFormyAkcjeSpec Then
            PokazFiltryColumnDGV(Me.AkcjeSpec, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltryAkcjeSpec, StartFormyAkcjeSpec)
        End If
    End Sub

    Private Sub dagAkcjeSpec_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagAkcjeSpec.MouseMove
        'przypisz biezacy DataGrid do myGrid
        Dim myGridView As DataGridView = CType(sender, DataGridView)
        Dim hti As System.Windows.Forms.DataGridView.HitTestInfo
        hti = myGridView.HitTest(e.X, e.Y)

        Select Case hti.Type
            Case System.Windows.Forms.DataGrid.HitTestType.None     '"the background."
            Case System.Windows.Forms.DataGrid.HitTestType.Cell
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                'If Not StartFormyAkcjeSpec Then
                '    PokazFiltryColumnDGV(Me.AkcjeSpec, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltryAkcjeSpec, StartFormyAkcjeSpec)
                'End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                'myGrid.Select(hti.Row)
            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                If Not StartFormyAkcjeSpec Then
                    PokazFiltryColumnDGV(Me.AkcjeSpec, dagAkcjeSpec, DataViewAkcjeSpec.Table.Columns.Count, stbFiltryAkcjeSpec, StartFormyAkcjeSpec)
                End If
            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
        End Select
        'MessageBox.Show(message)
    End Sub

    Private Sub dagAkcjeSpec_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dagAkcjeSpec.MouseUp
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
                stbAkcjeSpec.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                stbAkcjeSpec.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                myGridView.CurrentCell = dagAkcjeSpec(hti.ColumnIndex, hti.RowIndex)
                myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnHeader
                message &= "the column header for column " & hti.ColumnIndex
                stbAkcjeSpec.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString
                stbAkcjeSpec.Panels(3).Text = "Sort: " & DataSetAkcjeSpec.Tables(0).Columns(hti.ColumnIndex).ColumnName

            Case System.Windows.Forms.DataGrid.HitTestType.RowHeader
                message &= "the row header for row " & hti.RowIndex
                stbAkcjeSpec.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
                'myGridView.CurrentCell = dagAkcjeSpec(hti.ColumnIndex, hti.RowIndex)
                'myGridView.CurrentCell.Selected = True
                'myGridView.Select(hti.RowIndex, hti.ColumnIndex)

            Case System.Windows.Forms.DataGrid.HitTestType.ColumnResize
                message &= "the column resizer for column " & hti.ColumnIndex
                stbAkcjeSpec.Panels(2).Text = "Ko: " & hti.ColumnIndex.ToString

            Case System.Windows.Forms.DataGrid.HitTestType.RowResize
                message &= "the row resizer for row " & hti.RowIndex
                stbAkcjeSpec.Panels(1).Text = "Wi: " & hti.RowIndex.ToString
            Case System.Windows.Forms.DataGrid.HitTestType.Caption
                message &= "the caption"
            Case System.Windows.Forms.DataGrid.HitTestType.ParentRows
                message &= "the parent row"
        End Select
        'MessageBox.Show(message)
    End Sub






    Private Sub dagAkcjeSpec_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dagAkcjeSpec.DoubleClick
        'cmdEdytujAkcjeSpec_Click(sender, e)
    End Sub

    Private Sub dagAkcjeSpec_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        'Select Case e.KeyCode
        '    Case Keys.Enter
        '        cmdEdytujAkcjeSpec_Click(sender, e)
        '    Case Keys.Insert
        '        cmdDodajAkcjeSpec_Click(sender, e)
        '    Case Keys.Delete
        '        cmdUsuñAkcjeSpec_Click(sender, e)
        '    Case Else
        'End Select
    End Sub



    Private Sub InitAkcjeSpec()
    End Sub

    Private Sub PobierzAkcjeSpec()
        'pobierz wartosci przed aktualizacja
    End Sub



    '***********************************************************
    '*** obsluga komend dla Osob kontaktowych
    '***********************************************************

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
                sqlConnectionAkcjeSpec.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try

            If sqlConnectionAkcjeSpec.State = ConnectionState.Open Then
                oWystapilConcurrencyException = False
                '------------------------------------------------------------
                Try
                    sqlDataAdapterAkcjeSpec.Update(DataSetAkcjeSpec, "TABELA_AkcjeSpec")
                Catch Ex As System.Exception
                    ViewInLog(Ex, Me.Name(), Nothing)
                End Try
                '------------------------------------------------------------
                If oWystapilConcurrencyException = True Then
                    sqlDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                End If
                sqlConnectionAkcjeSpec.Close()
            End If
        End If
    End Sub


    Private Sub OdswiezBazeAkcjeSpec()
        DataSetAkcjeSpec.Tables(0).Clear()

        If ParL_DataBaseType = "MS ACCESS" Then
            'Try
            '    dbConnectionAkcjeSpec.Open()
            '    If dbConnectionAkcjeSpec.State = ConnectionState.Open Then
            '        dbDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
            '        dbConnectionAkcjeSpec.Close()
            '    End If
            'Catch Ex As System.Exception
            '    ViewInLog(Ex, Me.Name(), Nothing)
            'End Try
        Else
            Try
                sqlConnectionAkcjeSpec.Open()
                If sqlConnectionAkcjeSpec.State = ConnectionState.Open Then
                    sqlDataAdapterAkcjeSpec.Fill(DataSetAkcjeSpec)
                    sqlConnectionAkcjeSpec.Close()
                End If
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
            End Try
        End If
    End Sub

    'Private Sub cmdDodajAkcjeSpec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodajAkcjeSpec.Click
    '    oOperacja = "DODAJ"
    '    InitAkcjeSpec()
    '    '----------------------------------------------
    '    Do While True
    '        oAktualizuj = False
    '        Dim Form As New EdytujKatalogOsob
    '        Form.ShowDialog()
    '        Form.Dispose()
    '        If Not oAktualizuj Then
    '            Exit Do
    '        Else
    '            Dim foundRow As DataRow
    '            Dim NewRow As DataRow
    '            ' Create an array for the key values to find.
    '            Dim findTheseVals(1) As Object
    '            findTheseVals(0) = oIdentKli
    '            findTheseVals(1) = oOsobaOso
    '            foundRow = DataSetAkcjeSpec.Tables(0).Rows.Find(findTheseVals)
    '            ' sprawdz czy jest...
    '            If Not (foundRow Is Nothing) Then
    '                MessageBox.Show("W katalogu znajduje siê ju¿ zapis dla którego " & vbCrLf &
    '                    "Klient = " & oIdentKli & vbCrLf &
    '                    "Osoba = " & oOsobaOso,
    '                    "Uwaga :",
    '                    System.Windows.Forms.MessageBoxButtons.OK,
    '                    MessageBoxIcon.Information,
    '                    MessageBoxDefaultButton.Button1)
    '            Else
    '                'stworz nowy rekord
    '                NewRow = DataSetAkcjeSpec.Tables(0).NewRow()
    '                NewRow("IDENTKLIENTA") = oIdentKli
    '                NewRow("OSOBA") = oOsobaOso
    '                NewRow("STANOWISKO") = oStanowiskoOso
    '                NewRow("KOMPETENCJE") = oKompetencjeOso
    '                NewRow("TELEFON") = oTelefonOso
    '                NewRow("EMAIL") = oeMailOso
    '                NewRow("VIP") = oVIPOso
    '                NewRow("WERSJA") = 1     'Wersja         Integer, 2

    '                'dodaj rekord do DataSet
    '                DataSetAkcjeSpec.Tables(0).Rows.Add(NewRow)
    '                'wyswietl ilosc rekordow
    '                AktualizujAkcjeSpec()
    '                'aktualizuj DataGrid
    '                dagAkcjeSpec.Update()
    '                Exit Do
    '            End If
    '        End If
    '    Loop
    'End Sub

    'Private Sub cmdUsuñAkcjeSpec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsuñAkcjeSpec.Click
    '    If DataSetAkcjeSpec.Tables(0).Rows.Count <= 0 Then
    '        MessageBox.Show("Nie potrafie usun¹æ nieistniej¹cego zapisu" & vbCrLf &
    '            "Proszê najpierw dopisaæ rekord do tabeli...",
    '            "Uwaga :",
    '            System.Windows.Forms.MessageBoxButtons.OK,
    '            MessageBoxIcon.Information,
    '            MessageBoxDefaultButton.Button1)
    '    Else
    '        If MessageBox.Show("Czy na pewno mam usun¹æ ten zapis ?",
    '                    "Proszê o potwierdzenie :",
    '                    System.Windows.Forms.MessageBoxButtons.YesNo,
    '                    MessageBoxIcon.Question,
    '                    MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.Yes Then
    '            oOperacja = "USUN"
    '            PobierzAkcjeSpec()
    '            Dim foundRow As DataRow
    '            ' Create an array for the key values to find.
    '            Dim findTheseVals(1) As Object
    '            findTheseVals(0) = oIdentKli
    '            findTheseVals(1) = oOsobaOso
    '            foundRow = DataSetAkcjeSpec.Tables(0).Rows.Find(findTheseVals)
    '            ' sprawdz czy jest...
    '            If Not (foundRow Is Nothing) Then
    '                foundRow.Delete()
    '                'wyswietl ilosc rekordow
    '                AktualizujAkcjeSpec()
    '            Else
    '            End If
    '        End If
    '    End If
    'End Sub

    'Private Sub cmdEdytujAkcjeSpec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytujAkcjeSpec.Click
    '    If DataSetAkcjeSpec.Tables(0).Rows.Count <= 0 Then
    '        'Raiseevent(Dodaj,Click )
    '        MessageBox.Show("Nie potrafie edytowaæ nieistniej¹cego zapisu" & vbCrLf &
    '            "Proszê najpierw dopisaæ rekord do tabeli...",
    '            "Uwaga :",
    '            System.Windows.Forms.MessageBoxButtons.OK,
    '            MessageBoxIcon.Information,
    '            MessageBoxDefaultButton.Button1)
    '    Else
    '        oOperacja = "EDYTUJ"
    '        PobierzAkcjeSpec()
    '        Dim Form As New EdytujKatalogOsob
    '        oAktualizuj = False
    '        Form.ShowDialog()
    '        If Not oAktualizuj Then
    '        Else
    '            Dim foundRow As DataRow
    '            ' Create an array for the key values to find.
    '            Dim findTheseVals(1) As Object
    '            findTheseVals(0) = oIdentKli
    '            findTheseVals(1) = oOsobaOso
    '            foundRow = DataSetAkcjeSpec.Tables(0).Rows.Find(findTheseVals)
    '            ' sprawdz czy jest...
    '            If Not (foundRow Is Nothing) Then
    '                'foundRow("IDENTKLIENTA") = oIdentKli
    '                foundRow("OSOBA") = oOsobaOso
    '                foundRow("STANOWISKO") = oStanowiskoOso
    '                foundRow("KOMPETENCJE") = oKompetencjeOso
    '                foundRow("TELEFON") = oTelefonOso
    '                foundRow("EMAIL") = oeMailOso
    '                foundRow("VIP") = oVIPOso
    '                foundRow("WERSJA") = ZmienWersje(oWersjaOso) 'Wersja         Integer, 2

    '                'wyswietl ilosc rekordow
    '                AktualizujAkcjeSpec()
    '                'aktualizuj DataGrid
    '                dagAkcjeSpec.Update()
    '            End If
    '        End If
    '    End If
    'End Sub


    'Private Sub menuAkcjeSpecEdytuj_Click(sender As Object, e As EventArgs) Handles menuAkcjeSpecEdytuj.Click
    '    cmdEdytujAkcjeSpec_Click(sender, e)
    'End Sub

    'Private Sub menuAkcjeSpecDodaj_Click(sender As Object, e As EventArgs) Handles menuAkcjeSpecDodaj.Click
    '    cmdDodajAkcjeSpec_Click(sender, e)
    'End Sub

    'Private Sub menuAkcjeSpecUsun_Click(sender As Object, e As EventArgs) Handles menuAkcjeSpecUsun.Click
    '    cmdUsuñAkcjeSpec_Click(sender, e)
    'End Sub

    Private Sub menuAkcjeSpecMultiL_Click(sender As Object, e As EventArgs) Handles menuAkcjeSpecMultiL.Click
        dagAkcjeSpec.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dagAkcjeSpec.Refresh()
    End Sub

    Private Sub menuAkcjeSpecOdswiez_Click(sender As Object, e As EventArgs) Handles menuAkcjeSpecOdswiez.Click
        OdswiezBazeAkcjeSpec()
    End Sub

    Private Sub menuAkcjeSpecSingleL_Click(sender As Object, e As EventArgs) Handles menuAkcjeSpecSingleL.Click
        dagAkcjeSpec.DefaultCellStyle.WrapMode = DataGridViewTriState.False
        dagAkcjeSpec.Refresh()
    End Sub








    Private Sub txtBranza_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBranza.TextChanged
        TestLen(txtBranza, "BRAN¯A", 100)
    End Sub
    Private Sub txtBranza_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBranza.GotFocus
        StartEdycjiTxt(txtBranza)
    End Sub
    Private Sub txtBranza_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBranza.LostFocus
        KoniecEdycjiTxt(txtBranza)
    End Sub


    Private Sub txtPodBranza_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPodbranza.TextChanged
        TestLen(txtPodbranza, "PODBRAN¯A", 100)
    End Sub
    Private Sub txtPodBranza_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPodbranza.GotFocus
        StartEdycjiTxt(txtPodbranza)
    End Sub
    Private Sub txtPodBranza_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPodbranza.LostFocus
        KoniecEdycjiTxt(txtPodbranza)
    End Sub


    Private Sub txtLiczbaZatrudnionych_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLiczbaZatrudnionych.TextChanged
        If TestNum(txtLiczbaZatrudnionych, "LICZBA ZATRUDNIONYCH") Then
        End If
    End Sub
    Private Sub txtLiczbaZatrudnionych_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLiczbaZatrudnionych.GotFocus
        StartEdycjiTxt(txtLiczbaZatrudnionych)
    End Sub
    Private Sub txtLiczbaZatrudnionych_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLiczbaZatrudnionych.LostFocus
        KoniecEdycjiTxt(txtLiczbaZatrudnionych)
    End Sub



    Private Sub txtLiczbaPracZdalnie_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLiczbaPracZdalnie.TextChanged
        If TestNum(txtLiczbaPracZdalnie, "LICZBA PRACUJ¥CYCH ZDALNIE") Then
        End If
    End Sub
    Private Sub txtLiczbaPracZdalnie_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLiczbaPracZdalnie.GotFocus
        StartEdycjiTxt(txtLiczbaPracZdalnie)
    End Sub
    Private Sub txtLiczbaPracZdalnie_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLiczbaPracZdalnie.LostFocus
        KoniecEdycjiTxt(txtLiczbaPracZdalnie)
    End Sub

    Private Sub cmdWybierzBranze_Click(sender As Object, e As EventArgs) Handles cmdWybierzBranze.Click
        oCoMamRobic = "WYBIERAJ"
        oIdBranzy = Trim(txtBranza.Text)
        Dim form As New KatalogBranze
        form.ShowDialog()
        If Len(Trim(oIdBranzy)) > 0 Then
            txtBranza.Text = oIdBranzy
        End If
    End Sub

    Private Sub cmdWybierzPodBranze_Click(sender As Object, e As EventArgs) Handles cmdWybierzPodbranze.Click
        oCoMamRobic = "WYBIERAJ"
        oIdBranzy = Trim(txtBranza.Text)
        oIdPodBranzy = Trim(txtPodbranza.Text)
        Dim form As New KatalogPodBranze
        form.ShowDialog()
        If Len(Trim(oIdBranzy)) > 0 Then
            txtBranza.Text = oIdBranzy
        End If
        If Len(Trim(oIdPodBranzy)) > 0 Then
            txtPodbranza.Text = oIdPodBranzy
        End If
    End Sub

End Class
