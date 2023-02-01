Imports System.Reflection ' For Missing.Value and BindingFlags
Imports System.Runtime.InteropServices ' For COMException
Imports Microsoft.Office.Interop.Excel

Imports System.Math
Imports System.Xml
Imports System.IO

Imports System.Drawing.Printing

Public Class frmEksportDanychRankingWskaznikow
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEksportDanychRankingWskaznikow))
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
        Me.cmdWybierz.Location = New System.Drawing.Point(488, 223)
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
        Me.txtKatalog.Location = New System.Drawing.Point(91, 227)
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
        Me.Label3.Location = New System.Drawing.Point(12, 212)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(508, 16)
        Me.Label3.TabIndex = 375
        Me.Label3.Text = "Proszê wybraæ katalog dyskowy do którego wpisaæ plik Excel z Raportem"
        '
        'lblKlient
        '
        Me.lblKlient.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblKlient.ForeColor = System.Drawing.Color.Navy
        Me.lblKlient.Location = New System.Drawing.Point(12, 230)
        Me.lblKlient.Name = "lblKlient"
        Me.lblKlient.Size = New System.Drawing.Size(112, 16)
        Me.lblKlient.TabIndex = 374
        Me.lblKlient.Text = "Dysk / Katalog"
        '
        'cbxRokDo
        '
        Me.cbxRokDo.Enabled = False
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
        Me.cbxMiesDo.Enabled = False
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
        Me.cbxRokOd.Enabled = False
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
        Me.cbxMiesOd.Enabled = False
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
        Me.lblRaport.Location = New System.Drawing.Point(15, 261)
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
        'frmEksportDanychRankingWskaznikow
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(555, 332)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.cmdGeneruj)
        Me.Controls.Add(Me.cmdStop)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "frmEksportDanychRankingWskaznikow"
        Me.Text = "Ranking WskaŸników"
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


    Dim WartGraniczna As Double = 0

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

    Private Sub frmEksportDanychRankingWskaznikow_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
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










    Private Function PrzygotujDane(ByVal CoAnalizujemy As Integer,
                                   ByRef Wyniki() As Double,
                                   ByVal RokAnal As String,
                                   ByVal WartoscG As Double) As Boolean
        Dim RetValue As Boolean = False
        Dim i As Integer = 0
        Dim drv As DataRowView = Nothing
        Dim dr As DataRow = Nothing

        Dim MiesiacOd As String = RokAnal & "-01"
        Dim MiesiacDo As String = RokAnal & "-12"
        Dim PopMiesiacOd As String = WyliczRok(RokAnal, -1) & "-01"
        Dim PopMiesiacDo As String = WyliczRok(RokAnal, -1) & "-12"

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
            Case 1
                'sumujemy wartosc sprzeda¿y LASER+ATRAMENT+ObcyLASER+ObcyATRAMENT
                If rbtTofiOddzial.Checked Then
                    dbSelectSQLKlienci = "SELECT * FROM " &
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, " &
                                 "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & PopMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & PopMiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & PopMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & PopMiesiacDo & "')),0)) AS POPSPRZEDAZ, " &
                                 "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
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
                                     "(SELECT IDENTKLIENTA, IDENTTOFI, SUBSTRING(IDENTTOFI,1,5) AS IDENTTOFIBASE,  " &
                                 "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & PopMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & PopMiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & PopMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & PopMiesiacDo & "')),0)) AS POPSPRZEDAZ, " &
                                 "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=KLT.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
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
                                     "(SELECT IDENTKLIENTA, " &
                                 "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & PopMiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & PopMiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & PopMiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & PopMiesiacDo & "')),0)) AS POPSPRZEDAZ, " &
                                 "(ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDO FROM Obroty WHERE (Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(Obroty.DATA,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(Obroty.DATA,1,7)<='" & MiesiacDo & "')),0) + " &
                                 " ISNULL((SELECT SUM(LWARTSPRZED+AWARTSPRZED+LOWARTSPRZED+AOWARTSPRZED) AS SPRZEDM FROM ObrotyMies WHERE (ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA) AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)>='" & MiesiacOd & "') AND (SUBSTRING(ObrotyMies.MIESIAC,1,7)<='" & MiesiacDo & "')),0)) AS SPRZEDAZ, " &
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
                dbSelectSQLKlienci &= " WHERE (KL.SPRZEDAZ>=" & PrzecNaKrop(WartoscG.ToString("F2")) & ") "
                dbSelectSQLKlienci &= " ORDER BY KL.SPRZEDAZ ASC"

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
                DataSetKlienci.Tables(0).PrimaryKey = New DataColumn() {DataSetKlienci.Tables(0).Columns("NRKARTY")}

                DataViewKlienci = New DataView(DataSetKlienci.Tables(0))
                DataViewKlienci.AllowDelete = False
                DataViewKlienci.AllowNew = False

                Wyniki(0) = 0
                Wyniki(1) = 0
                Wyniki(2) = 0
                Wyniki(3) = 0
                Wyniki(4) = 0
                Wyniki(5) = 0
                Wyniki(6) = 0
                Wyniki(7) = 0
                Wyniki(8) = 0
                Wyniki(9) = 0
                Wyniki(10) = 0
                Wyniki(11) = 0

                If DataViewKlienci.Count > 0 Then
                    For i = 0 To DataViewKlienci.Count - 1
                        drv = DataViewKlienci.Item(i)
                        Select Case CoAnalizujemy
                            Case 1
                                Wyniki(0) += drv("SPRZEDAZ01")
                                Wyniki(1) += drv("SPRZEDAZ02")
                                Wyniki(2) += drv("SPRZEDAZ03")
                                Wyniki(3) += drv("SPRZEDAZ04")
                                Wyniki(4) += drv("SPRZEDAZ05")
                                Wyniki(5) += drv("SPRZEDAZ06")
                                Wyniki(6) += drv("SPRZEDAZ07")
                                Wyniki(7) += drv("SPRZEDAZ08")
                                Wyniki(8) += drv("SPRZEDAZ09")
                                Wyniki(9) += drv("SPRZEDAZ10")
                                Wyniki(10) += drv("SPRZEDAZ11")
                                Wyniki(11) += drv("SPRZEDAZ12")

                            Case Else
                                Wyniki(0) += IIf(drv("SPRZEDAZ01") <> 0, 1, 0)
                                Wyniki(1) += IIf(drv("SPRZEDAZ02") <> 0, 1, 0)
                                Wyniki(2) += IIf(drv("SPRZEDAZ03") <> 0, 1, 0)
                                Wyniki(3) += IIf(drv("SPRZEDAZ04") <> 0, 1, 0)
                                Wyniki(4) += IIf(drv("SPRZEDAZ05") <> 0, 1, 0)
                                Wyniki(5) += IIf(drv("SPRZEDAZ06") <> 0, 1, 0)
                                Wyniki(6) += IIf(drv("SPRZEDAZ07") <> 0, 1, 0)
                                Wyniki(7) += IIf(drv("SPRZEDAZ08") <> 0, 1, 0)
                                Wyniki(8) += IIf(drv("SPRZEDAZ09") <> 0, 1, 0)
                                Wyniki(9) += IIf(drv("SPRZEDAZ10") <> 0, 1, 0)
                                Wyniki(10) += IIf(drv("SPRZEDAZ11") <> 0, 1, 0)
                                Wyniki(11) += IIf(drv("SPRZEDAZ12") <> 0, 1, 0)

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
    Dim NazwaStrukturyRaportu As String = "Analiza_RankingWskaznikow.xlsx"
    Dim NazwaPlikuDoGenerowania As String = "RankingWskaznikow" & IIf(Len(Trim(Par_IdentOddzialu)) = 0, "", "_" & Trim(Par_IdentOddzialu)) & ".xlsx"
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
        Dim bRok As String = Mid(SysData, 1, 4)
        Dim popRok As String = WyliczRok(bRok, -1)

        Dim PR() As Double = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        Dim BR() As Double = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}

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
                workshts = workbk.Worksheets
                'worksht = workshts.Item("Arkusz1")
                worksht = workshts.Item(1)
                If (worksht Is Nothing) Then
                    MessageBox.Show("Nie mogê otworzyæ arkusza EXCEL", "Uwaga :",
                                     System.Windows.Forms.MessageBoxButtons.OK,
                                     MessageBoxIcon.Information,
                                     MessageBoxDefaultButton.Button1)
                Else
                    'worksht.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnRight
                    'worksht.Columns.ColumnWidth = 100

                    range = worksht.Range("H1", Missing.Value)
                    range.Value2 = Trim(WartGraniczna.ToString("F2"))


                    'wiersz z okresami rozlicz
                    DoArkuszaTXT(worksht, "1",
                                      "_01." & popRok,
                                      "_02." & popRok,
                                      "_03." & popRok,
                                      "_04." & popRok,
                                      "_05." & popRok,
                                      "_06." & popRok,
                                      "_07." & popRok,
                                      "_08." & popRok,
                                      "_09." & popRok,
                                      "_10." & popRok,
                                      "_11." & popRok,
                                      "_12." & popRok,
                                        "_01." & bRok,
                                        "_02." & bRok,
                                        "_03." & bRok,
                                        "_04." & bRok,
                                        "_05." & bRok,
                                        "_06." & bRok,
                                        "_07." & bRok,
                                        "_08." & bRok,
                                        "_09." & bRok,
                                        "_10." & bRok,
                                        "_11." & bRok,
                                        "_12." & bRok)


                    '------------------------------------
                    lblRaport.Text = "1. Obrót pow. WG - obrót narastaj¹co"
                    System.Windows.Forms.Application.DoEvents()
                    PrzygotujDane(1, PR, popRok, WartGraniczna)
                    PrzygotujDane(1, BR, bRok, WartGraniczna)
                    DoArkuszaNUM(worksht, "2",
                                 PR(0), PR(1), PR(2), PR(3), PR(4), PR(5), PR(6), PR(7), PR(8), PR(9), PR(10), PR(11),
                                 BR(0), BR(1), BR(2), BR(3), BR(4), BR(5), BR(6), BR(7), BR(8), BR(9), BR(10), BR(11))



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
                                ByVal TxtPR01 As String,
                                ByVal TxtPR02 As String,
                                ByVal TxtPR03 As String,
                                ByVal TxtPR04 As String,
                                ByVal TxtPR05 As String,
                                ByVal TxtPR06 As String,
                                ByVal TxtPR07 As String,
                                ByVal TxtPR08 As String,
                                ByVal TxtPR09 As String,
                                ByVal TxtPR10 As String,
                                ByVal TxtPR11 As String,
                                ByVal TxtPR12 As String,
                                   ByVal TxtR01 As String,
                                   ByVal TxtR02 As String,
                                   ByVal TxtR03 As String,
                                   ByVal TxtR04 As String,
                                   ByVal TxtR05 As String,
                                   ByVal TxtR06 As String,
                                   ByVal TxtR07 As String,
                                   ByVal TxtR08 As String,
                                   ByVal TxtR09 As String,
                                   ByVal TxtR10 As String,
                                   ByVal TxtR11 As String,
                                   ByVal TxtR12 As String)


        Dim range As Microsoft.Office.Interop.Excel.Range = Nothing            'Setting the value for cell

        range = worksht.Range("K" & aRow, Missing.Value)
        range.Value2 = TxtPR01

        range = worksht.Range("L" & aRow, Missing.Value)
        range.Value2 = TxtPR02

        range = worksht.Range("M" & aRow, Missing.Value)
        range.Value2 = TxtPR03

        range = worksht.Range("N" & aRow, Missing.Value)
        range.Value2 = TxtPR04

        range = worksht.Range("O" & aRow, Missing.Value)
        range.Value2 = TxtPR05

        range = worksht.Range("P" & aRow, Missing.Value)
        range.Value2 = TxtPR06

        range = worksht.Range("Q" & aRow, Missing.Value)
        range.Value2 = TxtPR07

        range = worksht.Range("R" & aRow, Missing.Value)
        range.Value2 = TxtPR08

        range = worksht.Range("S" & aRow, Missing.Value)
        range.Value2 = TxtPR09

        range = worksht.Range("T" & aRow, Missing.Value)
        range.Value2 = TxtPR10

        range = worksht.Range("U" & aRow, Missing.Value)
        range.Value2 = TxtPR11

        range = worksht.Range("V" & aRow, Missing.Value)
        range.Value2 = TxtPR12
        '---------------------------------------
        range = worksht.Range("W" & aRow, Missing.Value)
        range.Value2 = TxtR01

        range = worksht.Range("X" & aRow, Missing.Value)
        range.Value2 = TxtR02

        range = worksht.Range("Y" & aRow, Missing.Value)
        range.Value2 = TxtR03

        range = worksht.Range("Z" & aRow, Missing.Value)
        range.Value2 = TxtR04

        range = worksht.Range("AA" & aRow, Missing.Value)
        range.Value2 = TxtR05

        range = worksht.Range("AB" & aRow, Missing.Value)
        range.Value2 = TxtR06

        range = worksht.Range("AC" & aRow, Missing.Value)
        range.Value2 = TxtR07

        range = worksht.Range("AD" & aRow, Missing.Value)
        range.Value2 = TxtR08

        range = worksht.Range("AE" & aRow, Missing.Value)
        range.Value2 = TxtR09

        range = worksht.Range("AF" & aRow, Missing.Value)
        range.Value2 = TxtR10

        range = worksht.Range("AG" & aRow, Missing.Value)
        range.Value2 = TxtR11

        range = worksht.Range("AH" & aRow, Missing.Value)
        range.Value2 = TxtR12
        '---------------------------------------
    End Sub



    Private Sub DoArkuszaNUM(ByRef worksht As _Worksheet, ByVal aRow As String,
                                ByVal NumPR01 As Double,
                                ByVal NumPR02 As Double,
                                ByVal NumPR03 As Double,
                                ByVal NumPR04 As Double,
                                ByVal NumPR05 As Double,
                                ByVal NumPR06 As Double,
                                ByVal NumPR07 As Double,
                                ByVal NumPR08 As Double,
                                ByVal NumPR09 As Double,
                                ByVal NumPR10 As Double,
                                ByVal NumPR11 As Double,
                                ByVal NumPR12 As Double,
                                   ByVal NumR01 As Double,
                                   ByVal NumR02 As Double,
                                   ByVal NumR03 As Double,
                                   ByVal NumR04 As Double,
                                   ByVal NumR05 As Double,
                                   ByVal NumR06 As Double,
                                   ByVal NumR07 As Double,
                                   ByVal NumR08 As Double,
                                   ByVal NumR09 As Double,
                                   ByVal NumR10 As Double,
                                   ByVal NumR11 As Double,
                                   ByVal NumR12 As Double)


        Dim range As Microsoft.Office.Interop.Excel.Range = Nothing            'Setting the value for cell

        range = worksht.Range("P" & aRow, Missing.Value)
        range.Value2 = NumPR01

        range = worksht.Range("Q" & aRow, Missing.Value)
        range.Value2 = NumPR02

        range = worksht.Range("R" & aRow, Missing.Value)
        range.Value2 = NumPR03

        range = worksht.Range("S" & aRow, Missing.Value)
        range.Value2 = NumPR04

        range = worksht.Range("T" & aRow, Missing.Value)
        range.Value2 = NumPR05

        range = worksht.Range("U" & aRow, Missing.Value)
        range.Value2 = NumPR06

        range = worksht.Range("V" & aRow, Missing.Value)
        range.Value2 = NumPR07

        range = worksht.Range("W" & aRow, Missing.Value)
        range.Value2 = NumPR08

        range = worksht.Range("X" & aRow, Missing.Value)
        range.Value2 = NumPR09

        range = worksht.Range("Y" & aRow, Missing.Value)
        range.Value2 = NumPR10

        range = worksht.Range("Z" & aRow, Missing.Value)
        range.Value2 = NumPR11

        range = worksht.Range("AA" & aRow, Missing.Value)
        range.Value2 = NumPR12
        '---------------------------------------
        range = worksht.Range("AB" & aRow, Missing.Value)
        range.Value2 = NumR01

        range = worksht.Range("AC" & aRow, Missing.Value)
        range.Value2 = NumR02

        range = worksht.Range("AD" & aRow, Missing.Value)
        range.Value2 = NumR03

        range = worksht.Range("AE" & aRow, Missing.Value)
        range.Value2 = NumR04

        range = worksht.Range("AF" & aRow, Missing.Value)
        range.Value2 = NumR05

        range = worksht.Range("AG" & aRow, Missing.Value)
        range.Value2 = NumR06

        range = worksht.Range("AH" & aRow, Missing.Value)
        range.Value2 = NumR07

        range = worksht.Range("AI" & aRow, Missing.Value)
        range.Value2 = NumR08

        range = worksht.Range("AJ" & aRow, Missing.Value)
        range.Value2 = NumR09

        range = worksht.Range("AK" & aRow, Missing.Value)
        range.Value2 = NumR10

        range = worksht.Range("AL" & aRow, Missing.Value)
        range.Value2 = NumR11

        range = worksht.Range("AM" & aRow, Missing.Value)
        range.Value2 = NumR12
        '---------------------------------------
        range = worksht.Range("AN" & aRow, Missing.Value)
        'range.Value2 = (NumPR01 + NumPR02 + NumPR03 + NumPR04 + NumPR05 + NumPR06 + NumPR07 + NumPR08 + NumPR09 + NumPR10 + NumPR11 + NumPR12) / 12
        range.Value2 = (NumPR12) / 12

        range = worksht.Range("AO" & aRow, Missing.Value)
        'range.Value2 = (NumR01 + NumR02 + NumR03 + NumR04 + NumR05 + NumR06 + NumR07 + NumR08 + NumR09 + NumR10 + NumR11 + NumR12) / 12
        range.Value2 = (NumR12) / 12

    End Sub


End Class
