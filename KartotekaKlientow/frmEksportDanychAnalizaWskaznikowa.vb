Imports System.Reflection ' For Missing.Value and BindingFlags
Imports System.Runtime.InteropServices ' For COMException
Imports Microsoft.Office.Interop.Excel

Imports System.Math
Imports System.Xml
Imports System.IO

Imports System.Drawing.Printing

Public Class frmEksportDanychAnalizaWskaznikowa
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
    Friend WithEvents cmdGeneruj As System.Windows.Forms.Button
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbxRokDo As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbxMiesDo As System.Windows.Forms.ComboBox
    Friend WithEvents cbxRokOd As System.Windows.Forms.ComboBox
    Friend WithEvents cbxMiesOd As System.Windows.Forms.ComboBox
    Friend WithEvents cmdWybierz As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtKatalog As System.Windows.Forms.TextBox
    Friend WithEvents lblKlient As System.Windows.Forms.Label
    Friend WithEvents dlgOpenFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents rbtTofiBazowy As RadioButton
    Friend WithEvents rbtTofiOddzial As RadioButton
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents rbtWGPobierzP As RadioButton
    Friend WithEvents rbtWGPobierzW As RadioButton
    Friend WithEvents rbtWGWylicz As RadioButton
    Friend WithEvents rbtIdentKli As RadioButton
    Friend WithEvents lblRaport As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEksportDanychAnalizaWskaznikowa))
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.cmdGeneruj = New System.Windows.Forms.Button()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.rbtIdentKli = New System.Windows.Forms.RadioButton()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.rbtTofiBazowy = New System.Windows.Forms.RadioButton()
        Me.rbtTofiOddzial = New System.Windows.Forms.RadioButton()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.rbtWGPobierzP = New System.Windows.Forms.RadioButton()
        Me.rbtWGPobierzW = New System.Windows.Forms.RadioButton()
        Me.rbtWGWylicz = New System.Windows.Forms.RadioButton()
        Me.cmdWybierz = New System.Windows.Forms.Button()
        Me.txtKatalog = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblKlient = New System.Windows.Forms.Label()
        Me.cbxRokDo = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbxMiesDo = New System.Windows.Forms.ComboBox()
        Me.cbxRokOd = New System.Windows.Forms.ComboBox()
        Me.cbxMiesOd = New System.Windows.Forms.ComboBox()
        Me.lblRaport = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dlgOpenFolder = New System.Windows.Forms.FolderBrowserDialog()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.Location = New System.Drawing.Point(8, 304)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(104, 25)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 84
        Me.picLogo.TabStop = False
        '
        'cmdGeneruj
        '
        Me.cmdGeneruj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdGeneruj.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdGeneruj.Image = CType(resources.GetObject("cmdGeneruj.Image"), System.Drawing.Image)
        Me.cmdGeneruj.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdGeneruj.Location = New System.Drawing.Point(344, 305)
        Me.cmdGeneruj.Name = "cmdGeneruj"
        Me.cmdGeneruj.Size = New System.Drawing.Size(113, 24)
        Me.cmdGeneruj.TabIndex = 83
        Me.cmdGeneruj.Text = "Generuj Raport"
        Me.cmdGeneruj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdStop
        '
        Me.cmdStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdStop.Image = CType(resources.GetObject("cmdStop.Image"), System.Drawing.Image)
        Me.cmdStop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStop.Location = New System.Drawing.Point(463, 305)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(81, 24)
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
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.cmdWybierz)
        Me.Panel1.Controls.Add(Me.txtKatalog)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.lblKlient)
        Me.Panel1.Controls.Add(Me.cbxRokDo)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.cbxMiesDo)
        Me.Panel1.Controls.Add(Me.cbxRokOd)
        Me.Panel1.Controls.Add(Me.cbxMiesOd)
        Me.Panel1.Controls.Add(Me.lblRaport)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(540, 291)
        Me.Panel1.TabIndex = 81
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.rbtIdentKli)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.rbtTofiBazowy)
        Me.Panel3.Controls.Add(Me.rbtTofiOddzial)
        Me.Panel3.Location = New System.Drawing.Point(0, 120)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(537, 80)
        Me.Panel3.TabIndex = 379
        '
        'rbtIdentKli
        '
        Me.rbtIdentKli.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtIdentKli.ForeColor = System.Drawing.Color.Navy
        Me.rbtIdentKli.Location = New System.Drawing.Point(39, 61)
        Me.rbtIdentKli.Name = "rbtIdentKli"
        Me.rbtIdentKli.Size = New System.Drawing.Size(481, 18)
        Me.rbtIdentKli.TabIndex = 385
        Me.rbtIdentKli.TabStop = True
        Me.rbtIdentKli.Text = "Wg identyfikatora klienta w Kartotece Klientów"
        Me.rbtIdentKli.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(12, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(508, 16)
        Me.Label5.TabIndex = 381
        Me.Label5.Text = "Sposób identyfikowania danych pobieranych do Raportu :"
        '
        'rbtTofiBazowy
        '
        Me.rbtTofiBazowy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtTofiBazowy.ForeColor = System.Drawing.Color.Navy
        Me.rbtTofiBazowy.Location = New System.Drawing.Point(39, 43)
        Me.rbtTofiBazowy.Name = "rbtTofiBazowy"
        Me.rbtTofiBazowy.Size = New System.Drawing.Size(481, 18)
        Me.rbtTofiBazowy.TabIndex = 383
        Me.rbtTofiBazowy.TabStop = True
        Me.rbtTofiBazowy.Text = "Wg bazowego symbolu TOFI klienta ( BBBBB )"
        Me.rbtTofiBazowy.UseVisualStyleBackColor = True
        '
        'rbtTofiOddzial
        '
        Me.rbtTofiOddzial.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtTofiOddzial.ForeColor = System.Drawing.Color.Navy
        Me.rbtTofiOddzial.Location = New System.Drawing.Point(39, 25)
        Me.rbtTofiOddzial.Name = "rbtTofiOddzial"
        Me.rbtTofiOddzial.Size = New System.Drawing.Size(481, 18)
        Me.rbtTofiOddzial.TabIndex = 382
        Me.rbtTofiOddzial.TabStop = True
        Me.rbtTofiOddzial.Text = "Wg symbolu TOFI klienta ( BBBBB + OOO )"
        Me.rbtTofiOddzial.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.rbtWGPobierzP)
        Me.Panel2.Controls.Add(Me.rbtWGPobierzW)
        Me.Panel2.Controls.Add(Me.rbtWGWylicz)
        Me.Panel2.Location = New System.Drawing.Point(0, 40)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(537, 80)
        Me.Panel2.TabIndex = 378
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(12, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(508, 16)
        Me.Label4.TabIndex = 381
        Me.Label4.Text = "Sposób wyliczania wartoœci granicznej dla potrzeb Raportu :"
        '
        'rbtWGPobierzP
        '
        Me.rbtWGPobierzP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtWGPobierzP.ForeColor = System.Drawing.Color.Navy
        Me.rbtWGPobierzP.Location = New System.Drawing.Point(39, 61)
        Me.rbtWGPobierzP.Name = "rbtWGPobierzP"
        Me.rbtWGPobierzP.Size = New System.Drawing.Size(481, 18)
        Me.rbtWGPobierzP.TabIndex = 384
        Me.rbtWGPobierzP.TabStop = True
        Me.rbtWGPobierzP.Text = "Pobierz % sumarycznej wartoœci sprzeda¿y w pop.roku z parametrów i wylicz Wart.Gr" &
    "an"
        Me.rbtWGPobierzP.UseVisualStyleBackColor = True
        '
        'rbtWGPobierzW
        '
        Me.rbtWGPobierzW.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtWGPobierzW.ForeColor = System.Drawing.Color.Navy
        Me.rbtWGPobierzW.Location = New System.Drawing.Point(39, 43)
        Me.rbtWGPobierzW.Name = "rbtWGPobierzW"
        Me.rbtWGPobierzW.Size = New System.Drawing.Size(481, 18)
        Me.rbtWGPobierzW.TabIndex = 383
        Me.rbtWGPobierzW.TabStop = True
        Me.rbtWGPobierzW.Text = "Pobierz Wart.Gran z Parametrów"
        Me.rbtWGPobierzW.UseVisualStyleBackColor = True
        '
        'rbtWGWylicz
        '
        Me.rbtWGWylicz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.rbtWGWylicz.ForeColor = System.Drawing.Color.Navy
        Me.rbtWGWylicz.Location = New System.Drawing.Point(39, 25)
        Me.rbtWGWylicz.Name = "rbtWGWylicz"
        Me.rbtWGWylicz.Size = New System.Drawing.Size(481, 18)
        Me.rbtWGWylicz.TabIndex = 382
        Me.rbtWGWylicz.TabStop = True
        Me.rbtWGWylicz.Text = "Wylicz Wart.Gran jako 80 % sumarycznej wart. sprzedazy w poprzednim roku"
        Me.rbtWGWylicz.UseVisualStyleBackColor = True
        '
        'cmdWybierz
        '
        Me.cmdWybierz.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdWybierz.Image = CType(resources.GetObject("cmdWybierz.Image"), System.Drawing.Image)
        Me.cmdWybierz.Location = New System.Drawing.Point(488, 225)
        Me.cmdWybierz.Name = "cmdWybierz"
        Me.cmdWybierz.Size = New System.Drawing.Size(32, 26)
        Me.cmdWybierz.TabIndex = 376
        '
        'txtKatalog
        '
        Me.txtKatalog.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtKatalog.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtKatalog.ForeColor = System.Drawing.Color.Purple
        Me.txtKatalog.Location = New System.Drawing.Point(91, 229)
        Me.txtKatalog.Name = "txtKatalog"
        Me.txtKatalog.Size = New System.Drawing.Size(429, 20)
        Me.txtKatalog.TabIndex = 373
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(12, 214)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(508, 16)
        Me.Label3.TabIndex = 375
        Me.Label3.Text = "Proszê wybraæ katalog dyskowy do którego wpisaæ plik Excel z Raportem"
        '
        'lblKlient
        '
        Me.lblKlient.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblKlient.ForeColor = System.Drawing.Color.Navy
        Me.lblKlient.Location = New System.Drawing.Point(12, 232)
        Me.lblKlient.Name = "lblKlient"
        Me.lblKlient.Size = New System.Drawing.Size(112, 16)
        Me.lblKlient.TabIndex = 374
        Me.lblKlient.Text = "Dysk / Katalog"
        '
        'cbxRokDo
        '
        Me.cbxRokDo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxRokDo.ForeColor = System.Drawing.Color.Purple
        Me.cbxRokDo.Location = New System.Drawing.Point(278, 13)
        Me.cbxRokDo.Name = "cbxRokDo"
        Me.cbxRokDo.Size = New System.Drawing.Size(50, 22)
        Me.cbxRokDo.TabIndex = 370
        Me.cbxRokDo.Text = "2007"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(251, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 16)
        Me.Label2.TabIndex = 372
        Me.Label2.Text = "Do . . . . . . . ."
        '
        'cbxMiesDo
        '
        Me.cbxMiesDo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxMiesDo.ForeColor = System.Drawing.Color.Purple
        Me.cbxMiesDo.Location = New System.Drawing.Point(334, 13)
        Me.cbxMiesDo.Name = "cbxMiesDo"
        Me.cbxMiesDo.Size = New System.Drawing.Size(37, 22)
        Me.cbxMiesDo.TabIndex = 371
        Me.cbxMiesDo.Text = "12"
        '
        'cbxRokOd
        '
        Me.cbxRokOd.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxRokOd.ForeColor = System.Drawing.Color.Purple
        Me.cbxRokOd.Location = New System.Drawing.Point(144, 13)
        Me.cbxRokOd.Name = "cbxRokOd"
        Me.cbxRokOd.Size = New System.Drawing.Size(50, 22)
        Me.cbxRokOd.TabIndex = 368
        Me.cbxRokOd.Text = "2007"
        '
        'cbxMiesOd
        '
        Me.cbxMiesOd.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cbxMiesOd.ForeColor = System.Drawing.Color.Purple
        Me.cbxMiesOd.Location = New System.Drawing.Point(200, 13)
        Me.cbxMiesOd.Name = "cbxMiesOd"
        Me.cbxMiesOd.Size = New System.Drawing.Size(37, 22)
        Me.cbxMiesOd.TabIndex = 369
        Me.cbxMiesOd.Text = "12"
        '
        'lblRaport
        '
        Me.lblRaport.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblRaport.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRaport.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblRaport.ForeColor = System.Drawing.Color.Navy
        Me.lblRaport.Location = New System.Drawing.Point(15, 263)
        Me.lblRaport.Name = "lblRaport"
        Me.lblRaport.Size = New System.Drawing.Size(505, 16)
        Me.lblRaport.TabIndex = 367
        Me.lblRaport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(12, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(168, 16)
        Me.Label1.TabIndex = 336
        Me.Label1.Text = "Okres do analizy . . . Od. . . . . . . . . . . . . . . . ."
        '
        'frmEksportDanychAnalizaWskaznikowa
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(555, 332)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.cmdGeneruj)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "frmEksportDanychAnalizaWskaznikowa"
        Me.Text = "Analiza Wska¿nikowa"
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
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
        InicjujListeLat(cbxRokOd)
        InicjujListeMiesiecy(cbxMiesOd)
        InicjujListeLat(cbxRokDo)
        InicjujListeMiesiecy(cbxMiesDo)

        cbxRokOd.Text = Mid(SysData, 1, 4)
        cbxMiesOd.Text = "01"
        cbxRokDo.Text = Mid(SysData, 1, 4)
        cbxMiesDo.Text = "12"

        rbtWGWylicz.Checked = True
        rbtIdentKli.Checked = True

        txtKatalog.Text = ParL_KatalogRaporty
        lblRaport.Text = ""
    End Sub

    Private Sub frmEksportDanychAnalizaWskaznikowa_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub


    Private Sub cmdWybierz_Click(sender As Object, e As EventArgs) Handles cmdWybierz.Click
        With dlgOpenFolder
            .Description = "Proszê wybraæ katalog dyskowy do eksportu danych"
            '.RootFolder = 
            .SelectedPath = txtKatalog.Text
            .ShowNewFolderButton = True
            'Me.Visible = False
            'If .ShowDialog() = Windows.Forms.DialogResult.OK Then
            If .ShowDialog() = DialogResult.OK Then
                txtKatalog.Text = .SelectedPath
            End If
        End With
    End Sub










    '-----------------------------------
    ' pobierz dane do zestawienia
    '-----------------------------------

    Dim rPozycja As Integer = 0
    Dim rIdent As String = ""
    Dim rPopRok As Double = 0
    Dim rPlan As Double = 0

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

    '---------------------------------------------
    Dim rSprzedPopRok As Double = 0
    Dim rSprzedPopRokP As Double = 0
    Dim rSprzedPopRok01 As Double = 0
    Dim rSprzedPopRok02 As Double = 0
    Dim rSprzedPopRok03 As Double = 0
    Dim rSprzedPopRok04 As Double = 0
    Dim rSprzedPopRok05 As Double = 0
    Dim rSprzedPopRok06 As Double = 0
    Dim rSprzedPopRok07 As Double = 0
    Dim rSprzedPopRok08 As Double = 0
    Dim rSprzedPopRok09 As Double = 0
    Dim rSprzedPopRok10 As Double = 0
    Dim rSprzedPopRok11 As Double = 0
    Dim rSprzedPopRok12 As Double = 0

    Dim rSprzed As Double = 0
    Dim rSprzedP As Double = 0
    Dim rSprzed01 As Double = 0
    Dim rSprzed02 As Double = 0
    Dim rSprzed03 As Double = 0
    Dim rSprzed04 As Double = 0
    Dim rSprzed05 As Double = 0
    Dim rSprzed06 As Double = 0
    Dim rSprzed07 As Double = 0
    Dim rSprzed08 As Double = 0
    Dim rSprzed09 As Double = 0
    Dim rSprzed10 As Double = 0
    Dim rSprzed11 As Double = 0
    Dim rSprzed12 As Double = 0
    '---------------------------------------------


    Dim WartGraniczna As Double = 0

    Private Sub InicjujRoboczy(ByVal pPozycja As Integer,
                       ByVal pIdent As String)
        rPozycja = pPozycja
        rIdent = pIdent
        rPopRok = 0
        rPlan = 0
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
    End Sub

    Private Sub ZapiszRoboczy()
        Dim dr As DataRow = Nothing

        dr = DataSetRoboczy.Tables(0).NewRow()

        dr("POZYCJA") = rPozycja
        dr("IDENT") = rIdent
        dr("POPROK") = rPopRok
        dr("PLAN") = rPlan
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

        DataSetRoboczy.Tables(0).Rows.Add(dr)
    End Sub













    Private Function PrzygotujDane2(ByVal CoAnalizujemy As Integer,
                                   ByVal MiesiacOd As String,
                                   ByVal MiesiacDo As String,
                                   ByVal WartoscG As Double) As Boolean
        Dim RetValue As Boolean = False
        Dim i As Integer = 0
        Dim drv As DataRowView = Nothing
        Dim dr As DataRow = Nothing

        Dim popMiesiacOd As String = WyliczMiesiac(MiesiacOd, -12)
        Dim popMiesiacDo As String = WyliczMiesiac(MiesiacDo, -12)
        Dim popRok As String = WyliczRok(Mid(MiesiacOd, 1, 4), -1)




        ''---------------------------------------------------------------------
        ''Kontakty HANDLOWE - nowe 05.09.2015
        ''-----------------------------------------
        'Public oUniqIdKon As String           'UNIQID            varchar(40)
        'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
        'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
        'Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
        'Public oTematKon As String            'TEMAT            Text, 50
        'Public oUwagiKon As String            'UWAGI            Memo
        'Public oWersjaKon As Integer          'WERSJA           Integer

        Dim War As String = ""
        Dim War2 As String = ""
        Select Case CoAnalizujemy
            Case 10
                War = " AND ((KontaktyHandlowe.UWAGI) LIKE '%WIZ%') "
            Case 11
                War = " AND ((KontaktyHandlowe.UWAGI) LIKE '%TEL%') "
            Case 12
                War = " AND ((KontaktyHandlowe.UWAGI) LIKE '%RZU%') "
            Case 13
                War2 = " AND (KL.WERYFPOTENC<>'') "
            Case 14
                War = " AND ((KontaktyHandlowe.UWAGI) LIKE '%TEST US£UGI%') "
            Case 15
                War = " AND ((KontaktyHandlowe.UWAGI) LIKE '%WIZ%') "
                War2 = " AND (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") "
            Case 16
                War = " AND  ((KontaktyHandlowe.UWAGI LIKE '%WIZ%') OR " &
                             "(KontaktyHandlowe.UWAGI LIKE '%TEL%') OR " &
                             "(KontaktyHandlowe.UWAGI LIKE '%RZU%') OR " &
                             "(KontaktyHandlowe.UWAGI LIKE '%TEST US£UGI%')) "
                War2 = " AND (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") "
        End Select



        'dbSelectSQLKontakty = "SELECT * FROM " & _
        '                      "(SELECT IDENTKLIENTA, SUBSTRING(SKUPPLAN,1,4) AS WERYFPOTENC, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & popMiesiacDo & "')" & War & ") AS ILEKONPOPROK, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')" & War & ") AS ILEKONP, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(popMiesiacOd, 0) & "')" & War & ") AS ILEKON01, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(popMiesiacOd, 1) & "')" & War & ") AS ILEKON02, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(popMiesiacOd, 2) & "')" & War & ") AS ILEKON03, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(popMiesiacOd, 3) & "')" & War & ") AS ILEKON04, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(popMiesiacOd, 4) & "')" & War & ") AS ILEKON05, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(popMiesiacOd, 5) & "')" & War & ") AS ILEKON06, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')" & War & ") AS ILEKON07, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(popMiesiacOd, 7) & "')" & War & ") AS ILEKON08, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(popMiesiacOd, 8) & "')" & War & ") AS ILEKON09, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(popMiesiacOd, 9) & "')" & War & ") AS ILEKON10, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(popMiesiacOd, 10) & "')" & War & ") AS ILEKON11, " & _
        '                         "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(popMiesiacOd, 11) & "')" & War & ") AS ILEKON12, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "')" & War & ") AS ILEKON, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')" & War & ") AS ILEKONP, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')" & War & ") AS ILEKON01, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')" & War & ") AS ILEKON02, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')" & War & ") AS ILEKON03, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')" & War & ") AS ILEKON04, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')" & War & ") AS ILEKON05, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')" & War & ") AS ILEKON06, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')" & War & ") AS ILEKON07, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')" & War & ") AS ILEKON08, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')" & War & ") AS ILEKON09, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')" & War & ") AS ILEKON10, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')" & War & ") AS ILEKON11, " & _
        '                           "(SELECT COUNT(*) AS ILEKON FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')" & War & ") AS ILEKON12, " & _
        '                " FROM Klienci) AS KL " & _
        '                " WHERE (KL.ILEKON>0) "


        Dim dbKlienci As String = ""
        dbKlienci = "SELECT A.[IDENTKLIENTA], " &
                        " A.SKUPPLAN, " &
                        "       Split.a.value('.', 'VARCHAR(100)') AS IDENTTOFI " &
                        "FROM (SELECT [IDENTKLIENTA], SKUPPLAN, " &
                        "             CAST('<M>' + REPLACE([NRTOFITXT], ',', '</M><M>') + '</M>' AS XML) AS IDENTTOFI " &
                        "      From Klienci) AS A " &
                        "CROSS APPLY IDENTTOFI.nodes ('/M') AS Split(a)"


        If rbtTofiOddzial.Checked Then
            'symbol TOFI taki jaki jest
            dbSelectSQLKontakty = "SELECT * FROM " &
                              "(SELECT IDENTKLIENTA, IDENTTOFI, SUBSTRING(SKUPPLAN,1,4) AS WERYFPOTENC, " &
                                            "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                            " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(SELECT COUNT(*) AS ILEKONT  FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & MiesiacDo & "')" & War & ") AS ILEKON, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')" & War & ") AS ILEKONP, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')" & War & ") AS ILEKON01, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')" & War & ") AS ILEKON02, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')" & War & ") AS ILEKON03, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')" & War & ") AS ILEKON04, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')" & War & ") AS ILEKON05, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')" & War & ") AS ILEKON06, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')" & War & ") AS ILEKON07, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')" & War & ") AS ILEKON08, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')" & War & ") AS ILEKON09, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')" & War & ") AS ILEKON10, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')" & War & ") AS ILEKON11, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')" & War & ") AS ILEKON12 " &
                        " FROM (" & dbKlienci & ") AS KLT ) AS KL " &
                        " WHERE (KL.ILEKON>0) " & War2


        ElseIf rbtTofiBazowy.Checked Then
            dbSelectSQLKontakty = "SELECT * FROM " &
                              "(SELECT IDENTKLIENTA, IDENTTOFI, SUBSTRING(IDENTTOFI,1,5) AS IDENTTOFIBASE, SUBSTRING(SKUPPLAN,1,4) AS WERYFPOTENC, " &
                                            "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                            " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & MiesiacDo & "')" & War & ") AS ILEKON, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')" & War & ") AS ILEKONP, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')" & War & ") AS ILEKON01, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')" & War & ") AS ILEKON02, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')" & War & ") AS ILEKON03, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')" & War & ") AS ILEKON04, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')" & War & ") AS ILEKON05, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')" & War & ") AS ILEKON06, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')" & War & ") AS ILEKON07, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')" & War & ") AS ILEKON08, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')" & War & ") AS ILEKON09, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')" & War & ") AS ILEKON10, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')" & War & ") AS ILEKON11, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')" & War & ") AS ILEKON12 " &
                        " FROM (" & dbKlienci & ") AS KLT ) AS KL " &
                        " WHERE (KL.ILEKON>0) " & War2


        ElseIf rbtIdentKli.Checked Then
            dbSelectSQLKontakty = "SELECT * FROM " &
                              "(SELECT IDENTKLIENTA, SUBSTRING(SKUPPLAN,1,4) AS WERYFPOTENC, " &
                                            "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                            " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & MiesiacDo & "')" & War & ") AS ILEKON, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')" & War & ") AS ILEKONP, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')" & War & ") AS ILEKON01, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')" & War & ") AS ILEKON02, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')" & War & ") AS ILEKON03, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')" & War & ") AS ILEKON04, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')" & War & ") AS ILEKON05, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')" & War & ") AS ILEKON06, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')" & War & ") AS ILEKON07, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')" & War & ") AS ILEKON08, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')" & War & ") AS ILEKON09, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')" & War & ") AS ILEKON10, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')" & War & ") AS ILEKON11, " &
                                   "(SELECT COUNT(*) AS ILEKONT FROM KontaktyHandlowe WHERE (KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(KontaktyHandlowe.DATAKONTAKTU,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')" & War & ") AS ILEKON12 " &
                        " FROM Klienci) AS KL " &
                        " WHERE (KL.ILEKON>0) " & War2


        Else
        End If


        DataSetKontakty = New DataSet
        DataViewKontakty = New DataView

        DataSetKontakty.EnforceConstraints = False

        DataSetKontakty.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionKontakty = New OleDb.OleDbConnection(sConnectionKontakty)
            dbCommandSelectKontakty = New OleDb.OleDbCommand(dbSelectSQLKontakty, dbConnectionKontakty)
            'dbCommandDeleteKontakty = New OleDb.OleDbCommand(sDeleteSQLKontakty, dbConnectionKontakty)
            'dbCommandUpdateKontakty = New OleDb.OleDbCommand(sUpdateSQLKontakty, dbConnectionKontakty)
            'dbCommandInsertKontakty = New OleDb.OleDbCommand(sInsertSQLKontakty, dbConnectionKontakty)
            dbDataAdapterKontakty = New OleDb.OleDbDataAdapter(dbCommandSelectKontakty)

            Dim DBMapowanieTabeliKontakty As System.Data.Common.DataTableMapping
            DBMapowanieTabeliKontakty = dbDataAdapterKontakty.TableMappings.Add("Table", "TABELA_Kontakty")

            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionKontakty.Open()
            Catch Ex As System.Exception
                ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ConnectionState = dbConnectionKontakty.State
            End Try
        Else
            sqlConnectionKontakty = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectKontakty = New SqlClient.SqlCommand(dbSelectSQLKontakty, sqlConnectionKontakty)
            'sqlCommandDeleteKontakty = New sqlclient.sqlCommand(sDeleteSQLKontakty, sqlconnectionKontakty)
            'sqlCommandUpdateKontakty = New sqlclient.sqlCommand(sUpdateSQLKontakty, sqlconnectionKontakty)
            'sqlCommandInsertKontakty = New sqlclient.sqlCommand(sInsertSQLKontakty, sqlconnectionKontakty)
            sqlDataAdapterKontakty = New SqlClient.SqlDataAdapter(sqlCommandSelectKontakty)

            Dim sqlMapowanieTabeliKontakty As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliKontakty = sqlDataAdapterKontakty.TableMappings.Add("Table", "TABELA_Kontakty")

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
                    dbDataAdapterKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
                    dbDataAdapterKontakty.Fill(DataSetKontakty)
                    dbConnectionKontakty.Close()
                Else
                    sqlDataAdapterKontakty.FillSchema(DataSetKontakty, SchemaType.Mapped, "TABELA_Kontakty")
                    sqlDataAdapterKontakty.Fill(DataSetKontakty)
                    sqlConnectionKontakty.Close()
                End If

                'DataSetKontakty.Tables(0).PrimaryKey = New DataColumn() {DataSetKontakty.Tables(0).Columns("UNIQID")}

                DataViewKontakty = New DataView(DataSetKontakty.Tables(0))
                DataViewKontakty.AllowDelete = False
                DataViewKontakty.AllowNew = False

                rSprzed = 0
                rSprzedP = 0
                rSprzed01 = 0
                rSprzed02 = 0
                rSprzed03 = 0
                rSprzed04 = 0
                rSprzed05 = 0
                rSprzed06 = 0
                rSprzed07 = 0
                rSprzed08 = 0
                rSprzed09 = 0
                rSprzed10 = 0
                rSprzed11 = 0
                rSprzed12 = 0

                If DataViewKontakty.Count > 0 Then
                    For i = 0 To DataViewKontakty.Count - 1
                        drv = DataViewKontakty.Item(i)

                        Select Case CoAnalizujemy
                            Case 10, 11, 12, 14
                                rSprzed += drv("ILEKON")
                                rSprzedP += drv("ILEKONP")
                                rSprzed01 += drv("ILEKON01")
                                rSprzed02 += drv("ILEKON02")
                                rSprzed03 += drv("ILEKON03")
                                rSprzed04 += drv("ILEKON04")
                                rSprzed05 += drv("ILEKON05")
                                rSprzed06 += drv("ILEKON06")
                                rSprzed07 += drv("ILEKON07")
                                rSprzed08 += drv("ILEKON08")
                                rSprzed09 += drv("ILEKON09")
                                rSprzed10 += drv("ILEKON10")
                                rSprzed11 += drv("ILEKON11")
                                rSprzed12 += drv("ILEKON12")

                            Case 13
                                If Len(Trim(drv("ILEKON"))) > 0 Then
                                    rSprzed += drv("ILEKON")
                                    rSprzedP += drv("ILEKONP")
                                    rSprzed01 += drv("ILEKON01")
                                    rSprzed02 += drv("ILEKON02")
                                    rSprzed03 += drv("ILEKON03")
                                    rSprzed04 += drv("ILEKON04")
                                    rSprzed05 += drv("ILEKON05")
                                    rSprzed06 += drv("ILEKON06")
                                    rSprzed07 += drv("ILEKON07")
                                    rSprzed08 += drv("ILEKON08")
                                    rSprzed09 += drv("ILEKON09")
                                    rSprzed10 += drv("ILEKON10")
                                    rSprzed11 += drv("ILEKON11")
                                    rSprzed12 += drv("ILEKON12")
                                End If

                        End Select

                    Next
                End If
                RetValue = True

            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)

            End Try
        End If

        Return RetValue
    End Function






















    'Private Function PrzygotujDane(ByVal CoAnalizujemy As Integer, _
    '                               ByVal MiesiacOd As String, _
    '                               ByVal MiesiacDo As String, _
    '                               ByVal WartoscG As Double) As Boolean
    '    Dim RetValue As Boolean = False
    '    Dim i As Integer = 0
    '    Dim drv As DataRowView = Nothing
    '    Dim dr As DataRow = Nothing

    '    Dim popMiesiacOd As String = WyliczMiesiac(MiesiacOd, -12)
    '    Dim popMiesiacDo As String = WyliczMiesiac(MiesiacDo, -12)
    '    Dim popRok As String = WyliczRok(Mid(MiesiacOd, 1, 4), -1)

    '    '-------------------------------------------
    '    Select Case CoAnalizujemy
    '        Case 1, 2, 3, 6, 8
    '            dbSelectSQLKlienci = "SELECT * FROM " & _
    '                                 "(SELECT IDENTKLIENTA, ROKOWANIABR, " & _
    '                           "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " & _
    '                           " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " & _
    '                              "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')),0) + " & _
    '                              " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')),0)) AS SPRZEDAZPOPROKP, " & _
    '                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 0) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 0) & "')),0)) AS SPRZEDAZPOPROK01, " & _
    '                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 1) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 1) & "')),0)) AS SPRZEDAZPOPROK02, " & _
    '                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 2) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 2) & "')),0)) AS SPRZEDAZPOPROK03, " & _
    '                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 3) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 3) & "')),0)) AS SPRZEDAZPOPROK04, " & _
    '                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 4) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 4) & "')),0)) AS SPRZEDAZPOPROK05, " & _
    '                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 5) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 5) & "')),0)) AS SPRZEDAZPOPROK06, " & _
    '                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')),0)) AS SPRZEDAZPOPROK07, " & _
    '                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 7) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 7) & "')),0)) AS SPRZEDAZPOPROK08, " & _
    '                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 8) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 8) & "')),0)) AS SPRZEDAZPOPROK09, " & _
    '                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 9) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 9) & "')),0)) AS SPRZEDAZPOPROK10, " & _
    '                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 10) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 10) & "')),0)) AS SPRZEDAZPOPROK11, " & _
    '                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 11) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 11) & "')),0)) AS SPRZEDAZPOPROK12, " & _
    '                                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " & _
    '                                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " & _
    '                                        "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " & _
    '                                        " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " & _
    '                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " & _
    '                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " & _
    '                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " & _
    '                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " & _
    '                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " & _
    '                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " & _
    '                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " & _
    '                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " & _
    '                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " & _
    '                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " & _
    '                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " & _
    '                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " & _
    '                            " FROM Klienci) AS KL "

    '            Select Case CoAnalizujemy
    '                Case 1, 6
    '                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
    '                Case 2
    '                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
    '                Case 3, 8
    '                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZPOPROK=0) AND (KL.ROKOWANIABR<>'') ORDER BY KL.SPRZEDAZ ASC"

    '                Case Else
    '                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") OR ((KL.SPRZEDAZPOPROK=0) AND (KL.ROKOWANIABR<>'')) ORDER BY KL.SPRZEDAZ ASC"
    '            End Select




    '        Case 4, 5
    '            dbSelectSQLKlienci = "SELECT * FROM " & _
    '                                 "(SELECT IDENTKLIENTA, ROKOWANIABR, " & _
    '                           "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " & _
    '                           " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " & _
    '                              "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')),0) + " & _
    '                              " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')),0)) AS SPRZEDAZPOPROKP, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 0) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 0) & "')),0)) AS SPRZEDAZPOPROK01, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 1) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 1) & "')),0)) AS SPRZEDAZPOPROK02, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 2) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 2) & "')),0)) AS SPRZEDAZPOPROK03, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 3) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 3) & "')),0)) AS SPRZEDAZPOPROK04, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 4) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 4) & "')),0)) AS SPRZEDAZPOPROK05, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 5) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 5) & "')),0)) AS SPRZEDAZPOPROK06, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')),0)) AS SPRZEDAZPOPROK07, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 7) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 7) & "')),0)) AS SPRZEDAZPOPROK08, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 8) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 8) & "')),0)) AS SPRZEDAZPOPROK09, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 9) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 9) & "')),0)) AS SPRZEDAZPOPROK10, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 10) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 10) & "')),0)) AS SPRZEDAZPOPROK11, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 11) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 11) & "')),0)) AS SPRZEDAZPOPROK12, " & _
    '                                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " & _
    '                                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " & _
    '                                        "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " & _
    '                                        " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " & _
    '                            " FROM Klienci) AS KL "

    '            Select Case CoAnalizujemy
    '                Case 4
    '                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"

    '                Case 5
    '                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"

    '                Case Else
    '                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") OR ((KL.SPRZEDAZPOPROK=0) AND (KL.ROKOWANIABR<>'')) ORDER BY KL.SPRZEDAZ ASC"
    '            End Select



    '        Case 7, 9
    '            dbSelectSQLKlienci = "SELECT * FROM " & _
    '                                 "(SELECT IDENTKLIENTA, ROKOWANIABR, " & _
    '                           "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " & _
    '                           " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " & _
    '                              "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')),0) + " & _
    '                              " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')),0)) AS SPRZEDAZPOPROKP, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 0) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 0) & "')),0)) AS SPRZEDAZPOPROK01, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 1) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 1) & "')),0)) AS SPRZEDAZPOPROK02, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 2) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 2) & "')),0)) AS SPRZEDAZPOPROK03, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 3) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 3) & "')),0)) AS SPRZEDAZPOPROK04, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 4) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 4) & "')),0)) AS SPRZEDAZPOPROK05, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 5) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 5) & "')),0)) AS SPRZEDAZPOPROK06, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 6) & "')),0)) AS SPRZEDAZPOPROK07, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 7) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 7) & "')),0)) AS SPRZEDAZPOPROK08, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 8) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 8) & "')),0)) AS SPRZEDAZPOPROK09, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 9) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 9) & "')),0)) AS SPRZEDAZPOPROK10, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 10) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 10) & "')),0)) AS SPRZEDAZPOPROK11, " & _
    '                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(popMiesiacOd, 11) & "')),0) + " & _
    '                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(popMiesiacOd, 11) & "')),0)) AS SPRZEDAZPOPROK12, " & _
    '                                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " & _
    '                                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " & _
    '                                        "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " & _
    '                                        " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " & _
    '                                  "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " & _
    '                                  " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " & _
    '                            " FROM Klienci) AS KL "

    '            Select Case CoAnalizujemy
    '                Case 7, 9
    '                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"

    '                Case Else
    '                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") OR ((KL.SPRZEDAZPOPROK=0) AND (KL.ROKOWANIABR<>'')) ORDER BY KL.SPRZEDAZ ASC"
    '            End Select



    '    End Select

    '    DataSetKlienci = New DataSet
    '    DataViewKlienci = New DataView

    '    DataSetKLT.Locale = New System.Globalization.CultureInfo("pl-PL")

    '    If ParL_DataBaseType = "MS ACCESS" Then
    '        dbConnectionKlienci = New OleDb.OleDbConnection(sConnectionKlienci)
    '        dbCommandSelectKlienci = New OleDb.OleDbCommand(dbSelectSQLKlienci, dbConnectionKlienci)
    '        'dbCommandDeleteKlienci = New OleDb.OleDbCommand(sDeleteSQLKlienci, dbConnectionKlienci)
    '        'dbCommandUpdateKlienci = New OleDb.OleDbCommand(sUpdateSQLKlienci, dbConnectionKlienci)
    '        'dbCommandInsertKlienci = New OleDb.OleDbCommand(sInsertSQLKlienci, dbConnectionKlienci)
    '        dbDataAdapterKlienci = New OleDb.OleDbDataAdapter(dbCommandSelectKlienci)

    '        Dim DBMapowanieTabeliKlienci As System.Data.Common.DataTableMapping
    '        DBMapowanieTabeliKlienci = dbDataAdapterKLT.TableMappings.Add("Table", "TABELA_Klienci")

    '        '------------------------------------------
    '        'wypelnij DATASET
    '        Try
    '            dbConnectionKLT.Open()
    '        Catch Ex As System.Exception
    '            ViewInLog(Ex, Me.Name(), Nothing)
    '            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
    '            '    System.Windows.Forms.MessageBoxButtons.OK, _
    '            '    MessageBoxIcon.Information, _
    '            '    MessageBoxDefaultButton.Button1)
    '        Finally
    '            ConnectionState = dbConnectionKLT.State
    '        End Try
    '    Else
    '        sqlConnectionKlienci = New SqlClient.SqlConnection(sConnectionKlienci)
    '        sqlCommandSelectKlienci = New SqlClient.SqlCommand(dbSelectSQLKlienci, sqlConnectionKlienci)
    '        'sqlCommandDeleteKlienci = New sqlclient.sqlCommand(sDeleteSQLKlienci, sqlconnectionKlienci)
    '        'sqlCommandUpdateKlienci = New sqlclient.sqlCommand(sUpdateSQLKlienci, sqlconnectionKlienci)
    '        'sqlCommandInsertKlienci = New sqlclient.sqlCommand(sInsertSQLKlienci, sqlconnectionKlienci)
    '        sqlDataAdapterKlienci = New SqlClient.SqlDataAdapter(sqlCommandSelectKlienci)

    '        Dim sqlMapowanieTabeliKlienci As System.Data.Common.DataTableMapping
    '        sqlMapowanieTabeliKlienci = sqlDataAdapterKLT.TableMappings.Add("Table", "TABELA_Klienci")

    '        '------------------------------------------
    '        'wypelnij DATASET
    '        Try
    '            sqlConnectionKLT.Open()
    '        Catch Ex As System.Exception
    '            ViewInLog(Ex, Me.Name(), Nothing)
    '            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
    '            '    System.Windows.Forms.MessageBoxButtons.OK, _
    '            '    MessageBoxIcon.Information, _
    '            '    MessageBoxDefaultButton.Button1)
    '        Finally
    '            ConnectionState = sqlConnectionKLT.State
    '        End Try
    '    End If

    '    If ConnectionState = ConnectionState.Open Then
    '        Try
    '            If ParL_DataBaseType = "MS ACCESS" Then
    '                dbDataAdapterKLT.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
    '                dbDataAdapterKLT.Fill(DataSetKlienci)
    '                dbConnectionKLT.Close()
    '            Else
    '                sqlDataAdapterKLT.FillSchema(DataSetKlienci, SchemaType.Mapped, "TABELA_Klienci")
    '                sqlDataAdapterKLT.Fill(DataSetKlienci)
    '                sqlConnectionKLT.Close()
    '            End If
    '            DataSetKLT.Tables(0).PrimaryKey = New DataColumn() {DataSetKLT.Tables(0).Columns("NRKARTY")}

    '            DataViewKlienci = New DataView(DataSetKLT.Tables(0))
    '            DataViewKLT.AllowDelete = False
    '            DataViewKLT.AllowNew = False

    '            rSprzed = 0
    '            rSprzedP = 0
    '            rSprzed01 = 0
    '            rSprzed02 = 0
    '            rSprzed03 = 0
    '            rSprzed04 = 0
    '            rSprzed05 = 0
    '            rSprzed06 = 0
    '            rSprzed07 = 0
    '            rSprzed08 = 0
    '            rSprzed09 = 0
    '            rSprzed10 = 0
    '            rSprzed11 = 0
    '            rSprzed12 = 0

    '            rSprzedPopRok = 0
    '            rSprzedPopRokP = 0
    '            rSprzedPopRok01 = 0
    '            rSprzedPopRok02 = 0
    '            rSprzedPopRok03 = 0
    '            rSprzedPopRok04 = 0
    '            rSprzedPopRok05 = 0
    '            rSprzedPopRok06 = 0
    '            rSprzedPopRok07 = 0
    '            rSprzedPopRok08 = 0
    '            rSprzedPopRok09 = 0
    '            rSprzedPopRok10 = 0
    '            rSprzedPopRok11 = 0
    '            rSprzedPopRok12 = 0

    '            If DataViewKLT.Count > 0 Then
    '                For i = 0 To DataViewKLT.Count - 1
    '                    drv = DataViewKLT.Item(i)

    '                    Select Case CoAnalizujemy
    '                        Case 1, 2, 3, 4, 5      'Wartosciowe
    '                            rSprzedPopRok += drv("SPRZEDAZPOPROK")
    '                            rSprzedPopRokP += drv("SPRZEDAZPOPROKP")
    '                            rSprzedPopRok01 += drv("SPRZEDAZPOPROK01")
    '                            rSprzedPopRok02 += drv("SPRZEDAZPOPROK02")
    '                            rSprzedPopRok03 += drv("SPRZEDAZPOPROK03")
    '                            rSprzedPopRok04 += drv("SPRZEDAZPOPROK04")
    '                            rSprzedPopRok05 += drv("SPRZEDAZPOPROK05")
    '                            rSprzedPopRok06 += drv("SPRZEDAZPOPROK06")
    '                            rSprzedPopRok07 += drv("SPRZEDAZPOPROK07")
    '                            rSprzedPopRok08 += drv("SPRZEDAZPOPROK08")
    '                            rSprzedPopRok09 += drv("SPRZEDAZPOPROK09")
    '                            rSprzedPopRok10 += drv("SPRZEDAZPOPROK10")
    '                            rSprzedPopRok11 += drv("SPRZEDAZPOPROK11")
    '                            rSprzedPopRok12 += drv("SPRZEDAZPOPROK12")

    '                            rSprzed += drv("SPRZEDAZ")
    '                            rSprzedP += drv("SPRZEDAZP")
    '                            rSprzed01 += drv("SPRZEDAZ01")
    '                            rSprzed02 += drv("SPRZEDAZ02")
    '                            rSprzed03 += drv("SPRZEDAZ03")
    '                            rSprzed04 += drv("SPRZEDAZ04")
    '                            rSprzed05 += drv("SPRZEDAZ05")
    '                            rSprzed06 += drv("SPRZEDAZ06")
    '                            rSprzed07 += drv("SPRZEDAZ07")
    '                            rSprzed08 += drv("SPRZEDAZ08")
    '                            rSprzed09 += drv("SPRZEDAZ09")
    '                            rSprzed10 += drv("SPRZEDAZ10")
    '                            rSprzed11 += drv("SPRZEDAZ11")
    '                            rSprzed12 += drv("SPRZEDAZ12")

    '                        Case 6, 7, 8, 9
    '                            rSprzed += IIf(drv("SPRZEDAZ") > 0, 1, 0)
    '                            rSprzedP += IIf(drv("SPRZEDAZP") > 0, 1, 0)
    '                            rSprzed01 += IIf(drv("SPRZEDAZ01") > 0, 1, 0)
    '                            rSprzed02 += IIf(drv("SPRZEDAZ02") > 0, 1, 0)
    '                            rSprzed03 += IIf(drv("SPRZEDAZ03") > 0, 1, 0)
    '                            rSprzed04 += IIf(drv("SPRZEDAZ04") > 0, 1, 0)
    '                            rSprzed05 += IIf(drv("SPRZEDAZ05") > 0, 1, 0)
    '                            rSprzed06 += IIf(drv("SPRZEDAZ06") > 0, 1, 0)
    '                            rSprzed07 += IIf(drv("SPRZEDAZ07") > 0, 1, 0)
    '                            rSprzed08 += IIf(drv("SPRZEDAZ08") > 0, 1, 0)
    '                            rSprzed09 += IIf(drv("SPRZEDAZ09") > 0, 1, 0)
    '                            rSprzed10 += IIf(drv("SPRZEDAZ10") > 0, 1, 0)
    '                            rSprzed11 += IIf(drv("SPRZEDAZ11") > 0, 1, 0)
    '                            rSprzed12 += IIf(drv("SPRZEDAZ12") > 0, 1, 0)

    '                            rSprzedPopRok += IIf(drv("SPRZEDAZPOPROK") > 0, 1, 0)
    '                            rSprzedPopRokP += IIf(drv("SPRZEDAZPOPROKP") > 0, 1, 0)
    '                            rSprzedPopRok01 += IIf(drv("SPRZEDAZPOPROK01") > 0, 1, 0)
    '                            rSprzedPopRok02 += IIf(drv("SPRZEDAZPOPROK02") > 0, 1, 0)
    '                            rSprzedPopRok03 += IIf(drv("SPRZEDAZPOPROK03") > 0, 1, 0)
    '                            rSprzedPopRok04 += IIf(drv("SPRZEDAZPOPROK04") > 0, 1, 0)
    '                            rSprzedPopRok05 += IIf(drv("SPRZEDAZPOPROK05") > 0, 1, 0)
    '                            rSprzedPopRok06 += IIf(drv("SPRZEDAZPOPROK06") > 0, 1, 0)
    '                            rSprzedPopRok07 += IIf(drv("SPRZEDAZPOPROK07") > 0, 1, 0)
    '                            rSprzedPopRok08 += IIf(drv("SPRZEDAZPOPROK08") > 0, 1, 0)
    '                            rSprzedPopRok09 += IIf(drv("SPRZEDAZPOPROK09") > 0, 1, 0)
    '                            rSprzedPopRok10 += IIf(drv("SPRZEDAZPOPROK10") > 0, 1, 0)
    '                            rSprzedPopRok11 += IIf(drv("SPRZEDAZPOPROK11") > 0, 1, 0)
    '                            rSprzedPopRok12 += IIf(drv("SPRZEDAZPOPROK12") > 0, 1, 0)


    '                    End Select







    '                Next
    '            End If
    '            RetValue = True

    '        Catch Ex As System.Exception
    '            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
    '                System.Windows.Forms.MessageBoxButtons.OK, _
    '                MessageBoxIcon.Information, _
    '                MessageBoxDefaultButton.Button1)

    '        End Try
    '    End If

    '    Return RetValue
    'End Function












    Private Function PrzygotujDane(ByVal CoAnalizujemy As Integer,
                               ByVal MiesiacOd As String,
                               ByVal MiesiacDo As String,
                               ByVal WartoscG As Double) As Boolean
        Dim RetValue As Boolean = False
        Dim i As Integer = 0
        Dim drv As DataRowView = Nothing
        Dim dr As DataRow = Nothing

        Dim popMiesiacOd As String = WyliczMiesiac(MiesiacOd, -12)
        Dim popMiesiacDo As String = WyliczMiesiac(MiesiacDo, -12)
        Dim popRok As String = WyliczRok(Mid(MiesiacOd, 1, 4), -1)

        '-------------------------------------------
        'Zamiana multivalue w polu NRTOFITXT na kolumny z kodem TOFI
        'http://sqljason.com/2010/05/converting-single-comma-separated-row.html
        '
        'SELECT A.[IDENTKLIENTA], 
        '       Split.a.value('.', 'VARCHAR(100)') AS IDENTTOFI
        'FROM(SELECT [IDENTKLIENTA],
        '             CAST('<M>' + REPLACE([NRTOFITXT], ',', '</M><M>') + '</M>' AS XML) AS IDENTTOFI
        '     From Klienci) AS A 
        'CROSS APPLY IDENTTOFI.nodes ('/M') AS Split(a); 
        '-------------------------------------------


        Dim dbKlienci As String = ""
        dbKlienci = "SELECT A.[IDENTKLIENTA], " &
                        " A.ROKOWANIABR, " &
                        "       Split.a.value('.', 'VARCHAR(100)') AS IDENTTOFI " &
                        "FROM (SELECT [IDENTKLIENTA], ROKOWANIABR, " &
                        "             CAST('<M>' + REPLACE([NRTOFITXT], ',', '</M><M>') + '</M>' AS XML) AS IDENTTOFI " &
                        "      From Klienci) AS A " &
                        "CROSS APPLY IDENTTOFI.nodes ('/M') AS Split(a)"





        '-------------------------------------------
        Select Case CoAnalizujemy
            Case 1, 3, 6, 8
                'sumujemy wartosc sprzeda¿y LASER+ATRAMENT+ObcyLASER+ObcyATRAMENT
                If rbtTofiOddzial.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM (" & dbKlienci & ") AS KLT ) AS KL "
                ElseIf rbtTofiBazowy.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, SUBSTRING(IDENTTOFI,1,5) AS IDENTTOFIBASE, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM (" & dbKlienci & ") AS KLT ) AS KL "
                ElseIf rbtIdentKli.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM Klienci) AS KL "
                Else
                End If



                Select Case CoAnalizujemy
                    Case 1, 6       'obrót sumarycznie
                        'dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                        dbSelectSQLKlienci &= " ORDER BY KL.SPRZEDAZ ASC"
                    Case 3, 8       'obrót sumarycznie na nowych (popeok=0) klientach rokuj¹cych (rokowania<>"")
                        dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZPOPROK=0) AND (KL.ROKOWANIABR<>'') ORDER BY KL.SPRZEDAZ ASC"

                    Case Else
                        dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") OR ((KL.SPRZEDAZPOPROK=0) AND (KL.ROKOWANIABR<>'')) ORDER BY KL.SPRZEDAZ ASC"
                End Select



            Case 2
                'sumujemy wartosc sprzeda¿y LASER+ATRAMENT PRYZMAT
                If rbtTofiOddzial.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                     "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                     " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM (" & dbKlienci & ") AS KLT ) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                ElseIf rbtTofiBazowy.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, SUBSTRING(IDENTTOFI,1,5) AS IDENTTOFIBASE, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                     "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                     " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM (" & dbKlienci & ") AS KLT ) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                ElseIf rbtIdentKli.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                  "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                  " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                     "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                     " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                       "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                       " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM Klienci) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                Else
                End If






            Case 4
                'sumujemy marze LASER+ATRAMENT+ObcyLASER+ObcyATRAMENT
                If rbtTofiOddzial.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM (" & dbKlienci & ") AS KLT ) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                ElseIf rbtTofiBazowy.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, SUBSTRING(IDENTTOFI,1,5) AS IDENTTOFIBASE, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM (" & dbKlienci & ") AS KLT ) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                ElseIf rbtIdentKli.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED+LOMARSPRZED+AOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM Klienci) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                Else
                End If





            Case 5
                'sumujemy marze LASER+ATRAMENT PRYZMAT
                If rbtTofiOddzial.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM (" & dbKlienci & ") AS KLT ) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                ElseIf rbtTofiBazowy.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, SUBSTRING(IDENTTOFI,1,5) AS IDENTTOFIBASE, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM (" & dbKlienci & ") AS KLT ) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                ElseIf rbtIdentKli.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+AMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM Klienci) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                Else
                End If




            Case 7      'liczba wszystkich klientów laserowych
                'sumujemy marze LASER+ObcyLASER
                If rbtTofiOddzial.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM (" & dbKlienci & ") AS KLT ) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                ElseIf rbtTofiBazowy.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, SUBSTRING(IDENTTOFI,1,5) AS IDENTTOFIBASE, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM (" & dbKlienci & ") AS KLT ) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                ElseIf rbtIdentKli.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED+LOMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM Klienci) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                Else
                End If




            Case 9      'liczba  klientów laserowych PRYZMAT
                'sumujemy marze LASER
                If rbtTofiOddzial.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM (" & dbKlienci & ") AS KLT ) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                ElseIf rbtTofiBazowy.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, SUBSTRING(IDENTTOFI,1,5) AS IDENTTOFIBASE, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM (" & dbKlienci & ") AS KLT ) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                ElseIf rbtIdentKli.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, ROKOWANIABR, " &
                               "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDPOPROKO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & popMiesiacDo & "')),0) + " &
                               " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDPOPROKM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & popMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & popMiesiacDo & "')),0)) AS SPRZEDAZPOPROK, " &
                                 "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
                                   "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                   " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZP, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 0) & "')),0)) AS SPRZEDAZ01, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 1) & "')),0)) AS SPRZEDAZ02, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 2) & "')),0)) AS SPRZEDAZ03, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 3) & "')),0)) AS SPRZEDAZ04, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 4) & "')),0)) AS SPRZEDAZ05, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 5) & "')),0)) AS SPRZEDAZ06, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 6) & "')),0)) AS SPRZEDAZ07, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 7) & "')),0)) AS SPRZEDAZ08, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 8) & "')),0)) AS SPRZEDAZ09, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 9) & "')),0)) AS SPRZEDAZ10, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 10) & "')),0)) AS SPRZEDAZ11, " &
                                      "(ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0) + " &
                                      " ISNULL((SELECT SUM(LMARSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & WyliczMiesiac(MiesiacOd, 11) & "')),0)) AS SPRZEDAZ12 " &
                                " FROM Klienci) AS KL "
                    dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>" & WartoscG.ToString("F0") & ") ORDER BY KL.SPRZEDAZ ASC"
                Else
                End If



        End Select






        DataSetKlienci = New DataSet
        DataViewKlienci = New DataView

        DataSetKlienci.EnforceConstraints = False

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

                'DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}

                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                DataViewKlienci.AllowDelete = False
                DataViewKlienci.AllowNew = False

                rSprzed = 0
                rSprzedP = 0
                rSprzed01 = 0
                rSprzed02 = 0
                rSprzed03 = 0
                rSprzed04 = 0
                rSprzed05 = 0
                rSprzed06 = 0
                rSprzed07 = 0
                rSprzed08 = 0
                rSprzed09 = 0
                rSprzed10 = 0
                rSprzed11 = 0
                rSprzed12 = 0

                rSprzedPopRok = 0
                rSprzedPopRokP = 0
                rSprzedPopRok01 = 0
                rSprzedPopRok02 = 0
                rSprzedPopRok03 = 0
                rSprzedPopRok04 = 0
                rSprzedPopRok05 = 0
                rSprzedPopRok06 = 0
                rSprzedPopRok07 = 0
                rSprzedPopRok08 = 0
                rSprzedPopRok09 = 0
                rSprzedPopRok10 = 0
                rSprzedPopRok11 = 0
                rSprzedPopRok12 = 0

                If DataViewKlienci.Count > 0 Then
                    For i = 0 To DataViewKlienci.Count - 1
                        drv = DataViewKlienci.Item(i)

                        Select Case CoAnalizujemy
                            Case 1, 2, 3, 4, 5      'Wartosciowe
                                rSprzedPopRok += drv("SPRZEDAZPOPROK")
                                'rSprzedPopRokP += drv("SPRZEDAZPOPROKP")
                                'rSprzedPopRok01 += drv("SPRZEDAZPOPROK01")
                                'rSprzedPopRok02 += drv("SPRZEDAZPOPROK02")
                                'rSprzedPopRok03 += drv("SPRZEDAZPOPROK03")
                                'rSprzedPopRok04 += drv("SPRZEDAZPOPROK04")
                                'rSprzedPopRok05 += drv("SPRZEDAZPOPROK05")
                                'rSprzedPopRok06 += drv("SPRZEDAZPOPROK06")
                                'rSprzedPopRok07 += drv("SPRZEDAZPOPROK07")
                                'rSprzedPopRok08 += drv("SPRZEDAZPOPROK08")
                                'rSprzedPopRok09 += drv("SPRZEDAZPOPROK09")
                                'rSprzedPopRok10 += drv("SPRZEDAZPOPROK10")
                                'rSprzedPopRok11 += drv("SPRZEDAZPOPROK11")
                                'rSprzedPopRok12 += drv("SPRZEDAZPOPROK12")

                                rSprzed += drv("SPRZEDAZ")
                                rSprzedP += drv("SPRZEDAZP")
                                rSprzed01 += drv("SPRZEDAZ01")
                                rSprzed02 += drv("SPRZEDAZ02")
                                rSprzed03 += drv("SPRZEDAZ03")
                                rSprzed04 += drv("SPRZEDAZ04")
                                rSprzed05 += drv("SPRZEDAZ05")
                                rSprzed06 += drv("SPRZEDAZ06")
                                rSprzed07 += drv("SPRZEDAZ07")
                                rSprzed08 += drv("SPRZEDAZ08")
                                rSprzed09 += drv("SPRZEDAZ09")
                                rSprzed10 += drv("SPRZEDAZ10")
                                rSprzed11 += drv("SPRZEDAZ11")
                                rSprzed12 += drv("SPRZEDAZ12")

                            Case 6, 7, 8, 9     'Ilosc klientów
                                rSprzed += IIf(drv("SPRZEDAZ") <> 0, 1, 0)
                                rSprzedP += IIf(drv("SPRZEDAZP") <> 0, 1, 0)
                                rSprzed01 += IIf(drv("SPRZEDAZ01") <> 0, 1, 0)
                                rSprzed02 += IIf(drv("SPRZEDAZ02") <> 0, 1, 0)
                                rSprzed03 += IIf(drv("SPRZEDAZ03") <> 0, 1, 0)
                                rSprzed04 += IIf(drv("SPRZEDAZ04") <> 0, 1, 0)
                                rSprzed05 += IIf(drv("SPRZEDAZ05") <> 0, 1, 0)
                                rSprzed06 += IIf(drv("SPRZEDAZ06") <> 0, 1, 0)
                                rSprzed07 += IIf(drv("SPRZEDAZ07") <> 0, 1, 0)
                                rSprzed08 += IIf(drv("SPRZEDAZ08") <> 0, 1, 0)
                                rSprzed09 += IIf(drv("SPRZEDAZ09") <> 0, 1, 0)
                                rSprzed10 += IIf(drv("SPRZEDAZ10") <> 0, 1, 0)
                                rSprzed11 += IIf(drv("SPRZEDAZ11") <> 0, 1, 0)
                                rSprzed12 += IIf(drv("SPRZEDAZ12") <> 0, 1, 0)

                                rSprzedPopRok += IIf(drv("SPRZEDAZPOPROK") <> 0, 1, 0)
                                'rSprzedPopRokP += IIf(drv("SPRZEDAZPOPROKP") <> 0, 1, 0)
                                'rSprzedPopRok01 += IIf(drv("SPRZEDAZPOPROK01") <> 0, 1, 0)
                                'rSprzedPopRok02 += IIf(drv("SPRZEDAZPOPROK02") <> 0, 1, 0)
                                'rSprzedPopRok03 += IIf(drv("SPRZEDAZPOPROK03") <> 0, 1, 0)
                                'rSprzedPopRok04 += IIf(drv("SPRZEDAZPOPROK04") <> 0, 1, 0)
                                'rSprzedPopRok05 += IIf(drv("SPRZEDAZPOPROK05") <> 0, 1, 0)
                                'rSprzedPopRok06 += IIf(drv("SPRZEDAZPOPROK06") <> 0, 1, 0)
                                'rSprzedPopRok07 += IIf(drv("SPRZEDAZPOPROK07") <> 0, 1, 0)
                                'rSprzedPopRok08 += IIf(drv("SPRZEDAZPOPROK08") <> 0, 1, 0)
                                'rSprzedPopRok09 += IIf(drv("SPRZEDAZPOPROK09") <> 0, 1, 0)
                                'rSprzedPopRok10 += IIf(drv("SPRZEDAZPOPROK10") <> 0, 1, 0)
                                'rSprzedPopRok11 += IIf(drv("SPRZEDAZPOPROK11") <> 0, 1, 0)
                                'rSprzedPopRok12 += IIf(drv("SPRZEDAZPOPROK12") <> 0, 1, 0)


                        End Select







                    Next
                End If
                RetValue = True

            Catch Ex As System.Exception
                MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)

            End Try
        End If

        Return RetValue
    End Function









    '**************************************************************
    '** wydruk z bazy danych
    '**************************************************************

    '-----------------------------------------------------------
    Dim NazwaStrukturyRaportu As String = "Analiza_Wskaznikowa.xlsx"
    Dim NazwaPlikuDoGenerowania As String = "AnalizaWskaznikowa" & IIf(Len(Trim(Par_IdentOddzialu)) = 0, "", "_" & Trim(Par_IdentOddzialu)) & ".xlsx"
    Dim StrukturaExport As String = ""
    Dim PlikExport As String = ""

    Dim AnalPrac As String = ""

    Private Sub cmdGeneruj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGeneruj.Click
        'sprawdz czy jest struktura robocza
        StrukturaExport = oKatProgramu & "\" & NazwaStrukturyRaportu
        If Not IO.File.Exists(StrukturaExport) Then
            MessageBox.Show("Brak struktury wzorcowej pliku" & vbCrLf &
                            "do eksportu danych " & vbCrLf &
                            StrukturaExport,
                            "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        Else
            If Not IO.Directory.Exists(txtKatalog.Text) Then
                IO.Directory.CreateDirectory(txtKatalog.Text)
            End If

            If Mid(txtKatalog.Text, Len(txtKatalog.Text), 1) = "\" Then
                PlikExport = txtKatalog.Text & NazwaPlikuDoGenerowania
            Else
                PlikExport = txtKatalog.Text & "\" & NazwaPlikuDoGenerowania
            End If
            If IO.File.Exists(PlikExport) Then IO.File.Delete(PlikExport)

            'przekopiuj strukture wzorcow¹ do pliku wyjsciowego...
            Dim FileNo As Integer = FreeFile()
            If Len(Dir(PlikExport)) = 0 Then
                'nie ma pliku - utworz aby XCOPY nie mial problemu
                FileOpen(FileNo, PlikExport, OpenMode.Output)   ' Open file for output
                PrintLine(FileNo, "")
                FileClose(FileNo)
            End If
            '===================FileCopy(SrcFile,DstFile )
            Call Shell("XCopy """ & StrukturaExport & """ """ & PlikExport & """ /Y", AppWinStyle.NormalNoFocus, True)
            '-------------------------------------------

            'If DataViewRoboczy.Count() = 0 Then
            '    MessageBox.Show("Brak danych do Raportu.", _
            '        "Uwaga :", _
            '        System.Windows.Forms.MessageBoxButtons.OK, _
            '        MessageBoxIcon.Information, _
            '        MessageBoxDefaultButton.Button1)
            'Else
            '    '-----------------------
            '    Exportuj(PlikExport, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text)
            '    '----------------------
            '    MessageBox.Show("Zapisa³em Raport do pliku" & vbCrLf & _
            '                    PlikExport, _
            '        "Ok, skoñczy³em..", _
            '        System.Windows.Forms.MessageBoxButtons.OK, _
            '        MessageBoxIcon.Information, _
            '        MessageBoxDefaultButton.Button1)
            '    System.Windows.Forms.Application.DoEvents()
            'End If

            '-----------------------
            Exportuj(PlikExport, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text)
            '----------------------

            MessageBox.Show("Raport znajduje siê w pliku" & vbCrLf &
                            PlikExport,
                            "OK, Skoñczy³em.",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
        End If
        lblRaport.Text = ""
        System.Windows.Forms.Application.DoEvents()
    End Sub









    Private Sub Exportuj(ByVal PlikExcel As String, ByVal pMiesiacOd As String, ByVal pMiesiacDo As String)
        Dim dr As DataRowView = Nothing
        Dim LicznikRec As Long = 0
        Dim LicznikDok As Long = 0
        Dim ii As Integer = 0
        Dim BuforSort As String = ""

        Dim SymbolZam As String = ""
        Dim SymbolAnal As String = ""
        Dim IloscZam As Double = 0

        'otwieramy pliki do eksportu
        '---------------------------------
        Dim app As Application = Nothing
        Dim workbks As Workbooks = Nothing      'Getting the workbooks collection
        Dim workbk As _Workbook = Nothing       'adding a new workbook
        Dim workshts As Sheets = Nothing            'Getting the worksheets collection
        Dim worksht As _Worksheet = Nothing     'adding a new worksheet
        Dim range As Microsoft.Office.Interop.Excel.Range = Nothing            'Setting the value for cell

        Dim popMiesiacOd As String = WyliczMiesiac(pMiesiacOd, -12)
        Dim popMiesiacDo As String = WyliczMiesiac(pMiesiacDo, -12)
        Dim popRok As String = WyliczRok(Mid(pMiesiacOd, 1, 4), -1)
        Dim bRok As String = Mid(pMiesiacOd, 1, 4)

        Me.Cursor = Cursors.WaitCursor
        lblRaport.Text = "0. Wyliczam wartoœæ graniczn¹ "
        System.Windows.Forms.Application.DoEvents()

        If rbtWGWylicz.Checked Then
            WartGraniczna = WyliczWartoscGraniczna(popRok, 80)
        ElseIf rbtWGPobierzP.Checked Then
            WartGraniczna = WyliczWartoscGraniczna(popRok, Par_WartGranProcent)
        ElseIf rbtWGPobierzW.Checked Then
            WartGraniczna = Par_WartGranWartosc
        Else
            WartGraniczna = WyliczWartoscGraniczna(popRok, 80)
        End If


        Try
            app = New Application()
            If app Is Nothing Then
                MessageBox.Show("Nie mogê uruchomiæ programu EXCEL", "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
            Else
                app.Visible = False      'EXCEL WIDOCZNY NA EKRANIE
                workbks = app.Workbooks
                workbk = workbks.Add(XlWBATemplate.xlWBATWorksheet)
                workbk = workbks.Open(PlikExcel)
                'workshts = workbk.Worksheets
                worksht = workshts.Item(1)
                'worksht = workshts.Item("Arkusz1")
                If (worksht Is Nothing) Then
                    MessageBox.Show("Nie mogê otworzyæ arkusza EXCEL", "Uwaga :",
                                     System.Windows.Forms.MessageBoxButtons.OK,
                                     MessageBoxIcon.Information,
                                     MessageBoxDefaultButton.Button1)
                Else
                    'worksht.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnRight
                    'worksht.Columns.ColumnWidth = 100

                    range = worksht.Range("A1", Missing.Value)
                    range.Value2 = "Punkt sprzeda¿y PRYZMAT : " & Par_IdentOddzialu

                    range = worksht.Range("B2", Missing.Value)
                    range.Value2 = "WSKANIKI SPRZEDA¯OWE-WARTOŒCIOWE"

                    range = worksht.Range("B3", Missing.Value)
                    range.Value2 = "WARTOŒC GRANICZNA " & popRok & " : " & Trim(WartGraniczna.ToString("F2"))

                    range = worksht.Range("B24", Missing.Value)
                    range.Value2 = "WSKANIKI SPRZEDA¯OWE-ILOŒCIOWE"



                    '----------------------------
                    range = worksht.Range("B4", Missing.Value)
                    range.Value2 = "  Obrót w grupie Klientów 80% (powy¿ej Wartoœci Granicznej z " & popRok & "r.)             %realizacja/plan"

                    range = worksht.Range("B5", Missing.Value)
                    range.Value2 = "         Zrealizowany  w " & bRok & " r."

                    range = worksht.Range("B6", Missing.Value)
                    range.Value2 = "         Planowany w " & bRok & " r."

                    range = worksht.Range("B7", Missing.Value)
                    range.Value2 = "        Zrealizowany w " & popRok & " r."



                    range = worksht.Range("B9", Missing.Value)
                    range.Value2 = "         Zrealizowany  w " & bRok & " r."

                    range = worksht.Range("B10", Missing.Value)
                    range.Value2 = "         Planowany w " & bRok & " r."

                    range = worksht.Range("B11", Missing.Value)
                    range.Value2 = "        Zrealizowany w " & popRok & " r."



                    range = worksht.Range("B13", Missing.Value)
                    range.Value2 = "         Zrealizowany  w " & bRok & " r."

                    range = worksht.Range("B14", Missing.Value)
                    range.Value2 = "         Planowany w " & bRok & " r."

                    range = worksht.Range("B15", Missing.Value)
                    range.Value2 = "        Zrealizowany w " & popRok & " r."



                    range = worksht.Range("B17", Missing.Value)
                    range.Value2 = "         Zrealizowany  w " & bRok & " r."

                    range = worksht.Range("B18", Missing.Value)
                    range.Value2 = "         Planowany w " & bRok & " r."

                    range = worksht.Range("B19", Missing.Value)
                    range.Value2 = "        Zrealizowany w " & popRok & " r."



                    range = worksht.Range("B21", Missing.Value)
                    range.Value2 = "         Zrealizowany  w " & bRok & " r."

                    range = worksht.Range("B22", Missing.Value)
                    range.Value2 = "         Planowany w " & bRok & " r."

                    range = worksht.Range("B23", Missing.Value)
                    range.Value2 = "        Zrealizowany w " & popRok & " r."




                    range = worksht.Range("B26", Missing.Value)
                    range.Value2 = "      Uzyskana liczba w " & bRok & " r."

                    range = worksht.Range("B27", Missing.Value)
                    range.Value2 = "     Planowana liczba  w " & bRok & " r."

                    range = worksht.Range("B28", Missing.Value)
                    range.Value2 = "     Uzyskana liczba w " & popRok & " r."



                    range = worksht.Range("B30", Missing.Value)
                    range.Value2 = "      Uzyskana liczba w " & bRok & " r."

                    range = worksht.Range("B31", Missing.Value)
                    range.Value2 = "     Planowana liczba  w " & bRok & " r."

                    range = worksht.Range("B32", Missing.Value)
                    range.Value2 = "     Uzyskana liczba w " & popRok & " r."



                    range = worksht.Range("B34", Missing.Value)
                    range.Value2 = "      Uzyskana liczba w " & bRok & " r."

                    range = worksht.Range("B35", Missing.Value)
                    range.Value2 = "     Planowana liczba  w " & bRok & " r."

                    range = worksht.Range("B36", Missing.Value)
                    range.Value2 = "     Uzyskana liczba w " & popRok & " r."



                    range = worksht.Range("B38", Missing.Value)
                    range.Value2 = "      Uzyskana liczba w " & bRok & " r."

                    range = worksht.Range("B39", Missing.Value)
                    range.Value2 = "     Planowana liczba  w " & bRok & " r."

                    range = worksht.Range("B40", Missing.Value)
                    range.Value2 = "     Uzyskana liczba w " & popRok & " r."
                    '----------------------------







                    'wiersz z okresami rozlicz
                    DoArkuszaTXT(worksht, "2",
                                 pMiesiacOd & " - " & WyliczMiesiac(pMiesiacOd, +9),
                                 pMiesiacOd & " - " & pMiesiacDo,
                                 WyliczMiesiac(pMiesiacOd, +0),
                                 WyliczMiesiac(pMiesiacOd, +1),
                                 WyliczMiesiac(pMiesiacOd, +2),
                                 WyliczMiesiac(pMiesiacOd, +3),
                                 WyliczMiesiac(pMiesiacOd, +4),
                                 WyliczMiesiac(pMiesiacOd, +5),
                                 WyliczMiesiac(pMiesiacOd, +6),
                                 WyliczMiesiac(pMiesiacOd, +7),
                                 WyliczMiesiac(pMiesiacOd, +8),
                                 WyliczMiesiac(pMiesiacOd, +9),
                                 WyliczMiesiac(pMiesiacOd, +10),
                                 WyliczMiesiac(pMiesiacOd, +11))









                    '------------------------------------
                    lblRaport.Text = "1. Analizujê obrót w grupie klientów 80% (powy¿ej WG)"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(1, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "5",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)

                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(1, popMiesiacOd, popMiesiacDo, WartGraniczna)
                    DoArkuszaNUM(worksht, "7",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)
                    '--------------------------------------
                    lblRaport.Text = "2.  Obrót PRYZMAT w grupie Klientów 80% (powy¿ej WG) "
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(2, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "9",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)

                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(2, popMiesiacOd, popMiesiacDo, WartGraniczna)
                    DoArkuszaNUM(worksht, "11",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)
                    '--------------------------------------
                    lblRaport.Text = "3.  Obrót na nowych Klientach R (Rokuj¹cych) "
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(3, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "13",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)

                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(3, popMiesiacOd, popMiesiacDo, WartGraniczna)
                    DoArkuszaNUM(worksht, "15",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)
                    '--------------------------------------
                    lblRaport.Text = "4.  Mar¿a w grupie Klientów 80% (powy¿ej WG)"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(4, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "17",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)

                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(4, popMiesiacOd, popMiesiacDo, WartGraniczna)
                    DoArkuszaNUM(worksht, "19",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)
                    '--------------------------------------
                    lblRaport.Text = "5.  Mar¿a PRYZMAT w grupie Klientów 80% (powy¿ej WG)"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(5, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "21",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)

                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(5, popMiesiacOd, popMiesiacDo, WartGraniczna)
                    DoArkuszaNUM(worksht, "23",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)



                    '--------------------------------------
                    lblRaport.Text = "6.  Liczba klientów (powy¿ej WG)"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(6, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "26",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)

                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(6, popMiesiacOd, popMiesiacDo, WartGraniczna)
                    DoArkuszaNUM(worksht, "28",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)
                    '--------------------------------------
                    lblRaport.Text = "7.  Liczba klientów laserowych Pryzmat w grupie klientów 80% (powy¿ej WG)"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(7, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "30",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)

                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(7, popMiesiacOd, popMiesiacDo, WartGraniczna)
                    DoArkuszaNUM(worksht, "32",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)
                    '--------------------------------------
                    lblRaport.Text = "8.  Liczba nowych Klientów R (Rokuj¹cych)"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(8, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "34",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)

                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(8, popMiesiacOd, popMiesiacDo, WartGraniczna)
                    DoArkuszaNUM(worksht, "36",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)
                    '--------------------------------------
                    '9. Liczba tonerów Pryzmat w grupie klientów 80% (powy¿ej WG)
                    '--------------------------------------
                    lblRaport.Text = "9.  Liczba tonerów PRYZMAT"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(9, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "38",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)

                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(9, popMiesiacOd, popMiesiacDo, WartGraniczna)
                    DoArkuszaNUM(worksht, "40",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)


                    '--------------------------------------
                    lblRaport.Text = "10.  £¹czna aktywnoœæ wizytowa"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane2(10, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "44",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)
                    '--------------------------------------
                    lblRaport.Text = "11.  £¹czna aktywnoœæ telefoniczna"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane2(11, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "47",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)
                    '--------------------------------------
                    lblRaport.Text = "12.  Liczba firm z prezentacj¹ RZU"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane2(12, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "50",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)
                    '--------------------------------------
                    lblRaport.Text = "13.  Liczba firm ze zbadanym potencja³em zakupowym"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane2(13, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "53",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)
                    '--------------------------------------
                    lblRaport.Text = "14.  Liczba wizyt z testem us³ugi do klientów R"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane2(14, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "56",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)
                    '--------------------------------------
                    lblRaport.Text = "15.  Liczba firm pow. WG z wizyt¹ (zaczynaj¹c od Kl.P6)"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane2(15, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "59",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)
                    '--------------------------------------
                    lblRaport.Text = "16.  Liczba firm pow. WG z  podjêtym kontaktem (zaczynaj¹c od Kl.P6)"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane2(16, cbxRokOd.Text & "-" & cbxMiesOd.Text, cbxRokDo.Text & "-" & cbxMiesDo.Text, WartGraniczna)
                    DoArkuszaNUM(worksht, "62",
                                 rSprzed,
                                 rSprzedP,
                                 rSprzed01,
                                 rSprzed02,
                                 rSprzed03,
                                 rSprzed04,
                                 rSprzed05,
                                 rSprzed06,
                                 rSprzed07,
                                 rSprzed08,
                                 rSprzed09,
                                 rSprzed10,
                                 rSprzed11,
                                 rSprzed12)







                End If
            End If


        Catch ex As Exception
            MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.Message, "Uwaga :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)

        Finally
            range = Nothing
            worksht = Nothing
            workshts = Nothing

            workbk.Close(SaveChanges:=True)
            workbk = Nothing
            workbks = Nothing

            app.UserControl = False
            app.Quit()
            app = Nothing

            Me.Cursor = Cursors.Default
        End Try
    End Sub




    Private Sub DoArkuszaTXT(ByRef worksht As _Worksheet, ByVal aRow As String,
                             ByVal TxtR As String,
                             ByVal TxtP As String,
                             ByVal Txt01 As String,
                             ByVal Txt02 As String,
                             ByVal Txt03 As String,
                             ByVal Txt04 As String,
                             ByVal Txt05 As String,
                             ByVal Txt06 As String,
                             ByVal Txt07 As String,
                             ByVal Txt08 As String,
                             ByVal Txt09 As String,
                             ByVal Txt10 As String,
                             ByVal Txt11 As String,
                             ByVal Txt12 As String)
        Dim range As Microsoft.Office.Interop.Excel.Range = Nothing            'Setting the value for cell

        range = worksht.Range("C" & aRow, Missing.Value)
        range.Value2 = TxtR

        range = worksht.Range("D" & aRow, Missing.Value)
        range.Value2 = TxtP

        range = worksht.Range("E" & aRow, Missing.Value)
        range.Value2 = Txt01

        range = worksht.Range("F" & aRow, Missing.Value)
        range.Value2 = Txt02

        range = worksht.Range("G" & aRow, Missing.Value)
        range.Value2 = Txt03

        range = worksht.Range("H" & aRow, Missing.Value)
        range.Value2 = Txt04

        range = worksht.Range("I" & aRow, Missing.Value)
        range.Value2 = Txt05

        range = worksht.Range("J" & aRow, Missing.Value)
        range.Value2 = Txt06

        range = worksht.Range("K" & aRow, Missing.Value)
        range.Value2 = Txt07

        range = worksht.Range("L" & aRow, Missing.Value)
        range.Value2 = Txt08

        range = worksht.Range("M" & aRow, Missing.Value)
        range.Value2 = Txt09

        range = worksht.Range("N" & aRow, Missing.Value)
        range.Value2 = Txt10

        range = worksht.Range("O" & aRow, Missing.Value)
        range.Value2 = Txt11

        range = worksht.Range("P" & aRow, Missing.Value)
        range.Value2 = Txt12

    End Sub


    Private Sub DoArkuszaNUM(ByRef worksht As _Worksheet, ByVal aRow As String,
                         ByVal NumR As Double,
                         ByVal NumP As Double,
                         ByVal Num01 As Double,
                         ByVal Num02 As Double,
                         ByVal Num03 As Double,
                         ByVal Num04 As Double,
                         ByVal Num05 As Double,
                         ByVal Num06 As Double,
                         ByVal Num07 As Double,
                         ByVal Num08 As Double,
                         ByVal Num09 As Double,
                         ByVal Num10 As Double,
                         ByVal Num11 As Double,
                         ByVal Num12 As Double)
        Dim range As Microsoft.Office.Interop.Excel.Range = Nothing            'Setting the value for cell

        range = worksht.Range("C" & aRow, Missing.Value)
        range.Value2 = Num10

        range = worksht.Range("D" & aRow, Missing.Value)
        range.Value2 = NumR

        range = worksht.Range("E" & aRow, Missing.Value)
        range.Value2 = Num01

        range = worksht.Range("F" & aRow, Missing.Value)
        range.Value2 = Num02

        range = worksht.Range("G" & aRow, Missing.Value)
        range.Value2 = Num03

        range = worksht.Range("H" & aRow, Missing.Value)
        range.Value2 = Num04

        range = worksht.Range("I" & aRow, Missing.Value)
        range.Value2 = Num05

        range = worksht.Range("J" & aRow, Missing.Value)
        range.Value2 = Num06

        range = worksht.Range("K" & aRow, Missing.Value)
        range.Value2 = Num07

        range = worksht.Range("L" & aRow, Missing.Value)
        range.Value2 = Num08

        range = worksht.Range("M" & aRow, Missing.Value)
        range.Value2 = Num09

        range = worksht.Range("N" & aRow, Missing.Value)
        range.Value2 = Num10

        range = worksht.Range("O" & aRow, Missing.Value)
        range.Value2 = Num11

        range = worksht.Range("P" & aRow, Missing.Value)
        range.Value2 = Num12

    End Sub


End Class
