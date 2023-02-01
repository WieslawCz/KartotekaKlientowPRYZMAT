Imports System.Reflection

Public Class frmImportujKlientow
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Private _PZnakIdKlienta As String = ""
    Private _PZnakIdPracownika As String = ""

    Public Sub New(ByVal parPZnakIdKlienta As String, parPZnakIdPracownika As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _PZnakIdKlienta = parPZnakIdKlienta
        _PZnakIdPracownika = parPZnakIdPracownika
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
    Friend WithEvents cmdPobierz As System.Windows.Forms.Button
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtRaport As System.Windows.Forms.TextBox
    Friend WithEvents txtDataImportu As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdPobierz = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtDataImportu = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtRaport = New System.Windows.Forms.TextBox()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdPobierz
        '
        Me.cmdPobierz.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPobierz.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPobierz.Location = New System.Drawing.Point(497, 458)
        Me.cmdPobierz.Name = "cmdPobierz"
        Me.cmdPobierz.Size = New System.Drawing.Size(123, 23)
        Me.cmdPobierz.TabIndex = 32
        Me.cmdPobierz.Text = "Importuj Klientów"
        Me.cmdPobierz.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(626, 458)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 23)
        Me.cmdPowrot.TabIndex = 31
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.Location = New System.Drawing.Point(9, 457)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(104, 24)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 30
        Me.picLogo.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.txtDataImportu)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.txtRaport)
        Me.Panel1.Controls.Add(Me.ProgressBar)
        Me.Panel1.Location = New System.Drawing.Point(9, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(695, 442)
        Me.Panel1.TabIndex = 29
        '
        'txtDataImportu
        '
        Me.txtDataImportu.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.txtDataImportu.ForeColor = System.Drawing.Color.Purple
        Me.txtDataImportu.Location = New System.Drawing.Point(210, 7)
        Me.txtDataImportu.Name = "txtDataImportu"
        Me.txtDataImportu.ReadOnly = True
        Me.txtDataImportu.Size = New System.Drawing.Size(254, 20)
        Me.txtDataImportu.TabIndex = 312
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(5, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(229, 16)
        Me.Label1.TabIndex = 313
        Me.Label1.Text = "Data i godzina importu danych . . . . . . . . . . . . . . . . . . . . . . . . . ." &
    " . . "
        '
        'txtRaport
        '
        Me.txtRaport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRaport.Location = New System.Drawing.Point(8, 33)
        Me.txtRaport.Multiline = True
        Me.txtRaport.Name = "txtRaport"
        Me.txtRaport.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtRaport.Size = New System.Drawing.Size(671, 379)
        Me.txtRaport.TabIndex = 52
        '
        'ProgressBar
        '
        Me.ProgressBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar.Location = New System.Drawing.Point(8, 418)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(671, 8)
        Me.ProgressBar.TabIndex = 50
        '
        'frmImportujKlientow
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(713, 487)
        Me.Controls.Add(Me.cmdPobierz)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.Panel1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmImportujKlientow"
        Me.ShowInTaskbar = False
        Me.Text = "Importowanie danych z Bazy Danych Importowanych"
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region






    Private Sub frmObslugaPH_PobierzDaneSrodowiskoweZCentrali_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.picLogo.Image = My.Resources.logomini
        Me.cmdPowrot.Image = My.Resources._EXIT.ToBitmap
        'Me.cmdPobierz.Image = My.Resources.ok1.ToBitmap
        '-------------------------------
        txtDataImportu.Text = SysData() & "  " & SysGodz()
    End Sub

    Private Sub frmObslugaPH_PobierzDaneSrodowiskoweZCentrali_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub















    '*****************************************************
    '** Importujemy dane
    '*****************************************************

    '-----NIE HQConnectionUzytkownicy = oConnectionString
    'HQConnectionKlienci = oConnectionString                    IDENTKLIENTA
    'HQConnectionKontakty = oConnectionString                   UNIQID                  IDENTKLIENTA
    '-----NIE HQConnectionRodzajeKontaktow = oConnectionString
    'HQConnectionOsoby = oConnectionString                      IDENTKLIENTA+OSOBA      IDENTKLIENTA
    'HQConnectionObroty = oConnectionString                     IDENTKLIENTA+DATA       IDENTKLIENTA
    'HQConnectionObrotyMies = oConnectionString                 IDENTKLIENTA+MIESIAC    IDENTKLIENTA
    '-----NIE HQConnectionLogAktywnosci = oConnectionString         
    '-----NIE HQConnectionAkcjeSpec = oConnectionString         
    '-----NIE HQConnectionAnalizyZakupu = oConnectionString
    '-----NIE HQConnectionDaneDoRaportu = oConnectionString
    '-----NIE HQConnectionSlownikCoDalej = oConnectionString
    'HQConnectionCoDalej = oConnectionString                    UNIQID+IDENTKLIENTA     IDDENTKLIENTA
    '----NIE HQConnectionWersja = oConnectionString


    '----------------------
    'przenumerowanier klientów (nowy nr w NRTOFI INTEGER)
    '    UPDATE x
    '    Set x.NRTOFI = x.New_NRTOFI
    '    FROM (
    '      Select NRTOFI, ROW_NUMBER() OVER (ORDER BY [IDENTKLIENTA]) As New_NRTOFI
    '      FROM KLIENCI
    '      ) x
    '----------------------
    'zmiana klucza g³ównego z pola NRTOFI
    'ALTER TABLE Osoby DROP CONSTRAINT PK_Osoby

    'UPDATE [dbo].[Osoby]
    '   Set [IDENTKLIENTA] = RIGHT('000000' + LTRIM(CONVERT(VARCHAR(6),6891 + (SELECT NRTOFI From Klienci Where Osoby.IDENTKLIENTA=Klienci.IDENTKLIENTA))),6)

    'ALTER TABLE [dbo].[Osoby] With NOCHECK ADD
    '   CONSTRAINT [PK_Osoby] PRIMARY KEY  CLUSTERED 
    '   (   [IDENTKLIENTA],
    '       [OSOBA]
    '   )  ON [PRIMARY] 
    '----------------------
    'ALTER TABLE KontaktyHandlowe DROP CONSTRAINT PK_KontaktyHandlowe

    'UPDATE [dbo].[KontaktyHandlowe]
    '   Set [IDENTKLIENTA] = RIGHT('000000' + LTRIM(CONVERT(VARCHAR(6),6891 + (SELECT NRTOFI From Klienci Where KontaktyHandlowe.IDENTKLIENTA=Klienci.IDENTKLIENTA))),6)

    'ALTER TABLE [dbo].[KontaktyHandlowe] With NOCHECK ADD
    '   CONSTRAINT [PK_KontaktyHandlowe] PRIMARY KEY  CLUSTERED 
    '   (   [UNIQID]
    '   )  ON [PRIMARY] 
    '----------------------
    'ALTER TABLE Obroty DROP CONSTRAINT PK_Obroty

    'UPDATE [dbo].[Obroty]
    '   Set [IDENTKLIENTA] = RIGHT('000000' + LTRIM(CONVERT(VARCHAR(6),6891 + (SELECT NRTOFI From Klienci Where Obroty.IDENTKLIENTA=Klienci.IDENTKLIENTA))),6)

    'ALTER TABLE [dbo].[Obroty] With NOCHECK ADD
    '   CONSTRAINT [PK_Obroty] PRIMARY KEY  CLUSTERED 
    '   (   [IDENTKLIENTA],[DATA]
    '   )  ON [PRIMARY] 
    '----------------------
    ' ALTER TABLE ObrotyMies DROP CONSTRAINT PK_ObrotyMies

    'UPDATE [dbo].[ObrotyMies]
    '   Set [IDENTKLIENTA] = RIGHT('000000' + LTRIM(CONVERT(VARCHAR(6),6891 + (SELECT NRTOFI From Klienci Where ObrotyMies.IDENTKLIENTA=Klienci.IDENTKLIENTA))),6)

    'ALTER TABLE [dbo].[ObrotyMies] With NOCHECK ADD
    '   CONSTRAINT [PK_ObrotyMies] PRIMARY KEY  CLUSTERED 
    '   (   [IDENTKLIENTA],[MIESIAC]
    '   )  ON [PRIMARY] 
    '----------------------
    'ALTER TABLE CoDalejPlan DROP CONSTRAINT PK_CoDalej

    'UPDATE [dbo].[CoDalejPlan]
    '   Set [IDENTKLIENTA] = RIGHT('000000' + LTRIM(CONVERT(VARCHAR(6),6891 + (SELECT NRTOFI From Klienci Where CoDalejPlan.IDENTKLIENTA=Klienci.IDENTKLIENTA))),6)

    'ALTER TABLE [dbo].[CoDalejPlan] With NOCHECK ADD
    '   CONSTRAINT [PK_CoDalej] PRIMARY KEY  CLUSTERED 
    '   (   [UNIQID],[IDENTKLIENTA]
    '   )  ON [PRIMARY] 
    '----------------------
    'ALTER TABLE Klienci DROP CONSTRAINT PK_Klienci

    'UPDATE [dbo].[Klienci]
    '   Set [IDENTKLIENTA] = RIGHT('000000' + LTRIM(CONVERT(VARCHAR(6),6891 + NRTOFI)),6)

    'ALTER TABLE [dbo].[Klienci] With NOCHECK ADD
    '   CONSTRAINT [PK_Klienci] PRIMARY KEY  CLUSTERED 
    '   (   [IDENTKLIENTA]
    '   )  ON [PRIMARY] 


    Private Sub cmdImportuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPobierz.Click
        Dim WynikOK As Boolean = True
        Dim kwerenda As String = ""
        Dim ListaIdent As String = ""

        If MessageBox.Show("Czy na pewno mam po³¹czyæ Bazy Danych" & vbCrLf &
                           "i dopisac informacje z bazy importowanej " & ParL_HQDataBase & vbCrLf &
                           "do aktualnej Bazy Danych " & ParL_DataBase & " ?",
                    "Proszê o potwierdzenie :",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
            'OK - importujemy
            Try

                If Len(_PZnakIdKlienta) > 0 Then
                    PiszRaport("ZMIENIAM ID KLIENTA W BAZIE IMPORTOWANEJ")
                    If WynikOK Then
                        PiszRaport("Zmieniam Id Klienta dodaj¹c " & _PZnakIdKlienta & " na pocz¹tek identyfikatora klienta")

                        If WynikOK Then
                            kwerenda = "ALTER TABLE Klienci DROP CONSTRAINT PK_Klienci"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "W tabeli KLIENCI")
                        End If
                        If WynikOK Then
                            kwerenda = "UPDATE Klienci SET IDENTKLIENTA = '" & _PZnakIdKlienta & "'+SUBSTRING(IDENTKLIENTA,2,5)"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                        If WynikOK Then
                            kwerenda = "ALTER TABLE [dbo].[Klienci] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_Klienci] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENTKLIENTA]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                        '--------------------------------------------------------------------
                        If WynikOK Then
                            kwerenda = "ALTER TABLE AkcjeSpec DROP CONSTRAINT PK_AkcjeSpec"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "W tabeli AKCJESPEC")
                        End If
                        If WynikOK Then
                            kwerenda = "UPDATE AkcjeSpec SET IDENTKLI = '" & _PZnakIdKlienta & "'+SUBSTRING(IDENTKLI,2,5)"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                        If WynikOK Then
                            kwerenda = "ALTER TABLE [dbo].[AkcjeSpec] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_AkcjeSpec] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENTAKCJI], " & vbCrLf &
                               "   [IDENTKLI]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                        '--------------------------------------------------------------------
                        If WynikOK Then
                            kwerenda = "ALTER TABLE AnalizyZakupu DROP CONSTRAINT PK_AnalizyZakupu"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "W tabeli AnalizyZakupu")
                        End If
                        If WynikOK Then
                            kwerenda = "UPDATE AnalizyZakupu SET IDENTKLIENTA = '" & _PZnakIdKlienta & "'+SUBSTRING(IDENTKLIENTA,2,5)"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                        If WynikOK Then
                            kwerenda = "ALTER TABLE [dbo].[AnalizyZakupu] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_AnalizyZakupu] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENTKLIENTA] " & vbCrLf &
                               "   )  ON [PRIMARY] "
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                        '--------------------------------------------------------------------
                        If WynikOK Then
                            kwerenda = "ALTER TABLE CoDalejPlan DROP CONSTRAINT PK_CoDalej"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "W tabeli CoDalejPlan")
                        End If
                        If WynikOK Then
                            kwerenda = "UPDATE CoDalejPlan SET IDENTKLIENTA = '" & _PZnakIdKlienta & "'+SUBSTRING(IDENTKLIENTA,2,5)"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                        If WynikOK Then
                            kwerenda = "ALTER TABLE [dbo].[CoDalejPlan] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_CoDalej] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [UNIQID], " & vbCrLf &
                               "   [IDENTKLIENTA] " & vbCrLf &
                               "   )  ON [PRIMARY] "
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                        '--------------------------------------------------------------------
                        If WynikOK Then
                            kwerenda = "UPDATE KontaktyHandlowe SET IDENTKLIENTA = '" & _PZnakIdKlienta & "'+SUBSTRING(IDENTKLIENTA,2,5)"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "W tabeli KontaktyHandlowe")
                        End If
                        '--------------------------------------------------------------------
                        If WynikOK Then
                            kwerenda = "ALTER TABLE Obroty DROP CONSTRAINT PK_Obroty"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "W tabeli Obroty")
                        End If
                        If WynikOK Then
                            kwerenda = "UPDATE Obroty SET IDENTKLIENTA = '" & _PZnakIdKlienta & "'+SUBSTRING(IDENTKLIENTA,2,5)"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                        If WynikOK Then
                            kwerenda = "ALTER TABLE [dbo].[Obroty] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_Obroty] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENTKLIENTA], " & vbCrLf &
                               "   [DATA]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                        '--------------------------------------------------------------------
                        If WynikOK Then
                            kwerenda = "ALTER TABLE ObrotyMies DROP CONSTRAINT PK_ObrotyMies"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "W tabeli ObrotyMies")
                        End If
                        If WynikOK Then
                            kwerenda = "UPDATE ObrotyMies SET IDENTKLIENTA = '" & _PZnakIdKlienta & "'+SUBSTRING(IDENTKLIENTA,2,5)"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                        If WynikOK Then
                            kwerenda = "ALTER TABLE [dbo].[ObrotyMies] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_ObrotyMies] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENTKLIENTA], " & vbCrLf &
                               "   [MIESIAC]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                        '--------------------------------------------------------------------
                        If WynikOK Then
                            kwerenda = "ALTER TABLE Osoby DROP CONSTRAINT PK_Osoby"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "W tabeli Osoby")
                        End If
                        If WynikOK Then
                            kwerenda = "UPDATE Osoby SET IDENTKLIENTA = '" & _PZnakIdKlienta & "'+SUBSTRING(IDENTKLIENTA,2,5)"
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                        If WynikOK Then
                            kwerenda = "ALTER TABLE [dbo].[Osoby] WITH NOCHECK ADD " & vbCrLf &
                               "   CONSTRAINT [PK_Osoby] PRIMARY KEY  CLUSTERED " & vbCrLf &
                               "   (" & vbCrLf &
                               "   [IDENTKLIENTA], " & vbCrLf &
                               "   [OSOBA]" & vbCrLf &
                               "   )  ON [PRIMARY] "
                            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                        End If
                    End If
                End If

                'If Len(_PZnakIdKlienta) > 0 Then
                '    If WynikOK Then
                '        PiszRaport("")
                '        PiszRaport("ZMIENIAM ID PRACOWNIKA W BAZIE IMPORTOWANEJ")
                '        PiszRaport("Zmieniam Id Pracownika dodaj¹c " & _PZnakIdPracownika & " na pocz¹tek identyfikatora pracownika")

                '        If WynikOK Then
                '            kwerenda = "ALTER TABLE Uzytkownicy DROP CONSTRAINT PK_Uzytkownicy"
                '            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "W tabeli Uzytkownicy")
                '        End If
                '        If WynikOK Then
                '            kwerenda = "UPDATE Uzytkownicy SET IDENT = SUBSTRING('" & _PZnakIdPracownika & "'+LTRIM(IDENT),1,10)"
                '            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                '        End If
                '        If WynikOK Then
                '            kwerenda = "ALTER TABLE [dbo].[Uzytkownicy] WITH NOCHECK ADD " & vbCrLf &
                '               "   CONSTRAINT [PK_Uzytkownicy] PRIMARY KEY  CLUSTERED " & vbCrLf &
                '               "   (" & vbCrLf &
                '               "   [IDENT]" & vbCrLf &
                '               "   )  ON [PRIMARY] "
                '            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                '        End If
                '        '--------------------------------------------------------------------
                '        If WynikOK Then
                '            kwerenda = "UPDATE LogAktywnosci SET PRACOWNIK = SUBSTRING('" & _PZnakIdPracownika & "'+LTRIM(PRACOWNIK),1,10)"
                '            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "W tabeli LogAktywnosci")
                '        End If
                '        '--------------------------------------------------------------------
                '        If WynikOK Then
                '            kwerenda = "UPDATE KontaktyHandlowe SET PRACOWNIK = SUBSTRING('" & _PZnakIdPracownika & "'+LTRIM(PRACOWNIK),1,10)"
                '            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "W tabeli KontaktyHandlowe")
                '        End If
                '        '--------------------------------------------------------------------
                '        If WynikOK Then
                '            kwerenda = "UPDATE Klienci SET KTOWPISAL = SUBSTRING('" & _PZnakIdPracownika & "'+LTRIM(KTOWPISAL),1,10) WHERE LEN(LTRIM(KTOWPISAL))>0 "
                '            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "W tabeli Klienci")
                '        End If
                '        If WynikOK Then
                '            kwerenda = "UPDATE Klienci SET SPRZEDOPIEKUN = SUBSTRING('" & _PZnakIdPracownika & "'+LTRIM(SPRZEDOPIEKUN),1,10) WHERE LEN(LTRIM(SPRZEDOPIEKUN))>0 "
                '            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                '        End If
                '        '--------------------------------------------------------------------
                '        If WynikOK Then
                '            kwerenda = "ALTER TABLE DaneDoRaportu DROP CONSTRAINT PK_DaneDoRaportu"
                '            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "W tabeli DaneDoRaportu")
                '        End If
                '        If WynikOK Then
                '            kwerenda = "UPDATE DaneDoRaportu SET PRACOWNIK = SUBSTRING('" & _PZnakIdPracownika & "'+LTRIM(PRACOWNIK),1,10)"
                '            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                '        End If
                '        If WynikOK Then
                '            kwerenda = "ALTER TABLE [dbo].[DaneDoRaportu] WITH NOCHECK ADD " & vbCrLf &
                '               "   CONSTRAINT [PK_DaneDoRaportu] PRIMARY KEY  CLUSTERED " & vbCrLf &
                '               "   (" & vbCrLf &
                '               "   [PRACOWNIK], " & vbCrLf &
                '               "   [DATARAPORTU]" & vbCrLf &
                '               "   )  ON [PRIMARY] "
                '            WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")
                '        End If
                '        '--------------------------------------------------------------------



                '    End If
                'End If



                If WynikOK Then
                    If WynikOK Then
                        'kwerenda = "SELECT ',' + RTRIM(cast(IDENT as varchar(10))) as [text()] " &
                        '           "  FROM " & ParL_HQDataBase & ".[dbo].[Uzytkownicy] " &
                        '           "  WHERE (" &
                        '           "SELECT COUNT(*) from " & ParL_DataBase & ".[dbo].[Uzytkownicy] where (" & ParL_DataBase & ".[dbo].[Uzytkownicy].IDENT=" & ParL_HQDataBase & ".[dbo].[Uzytkownicy].IDENT) " &
                        '           ") > 0 " &
                        '           "  For xml path('')"

                        kwerenda = "SELECT ',' + RTRIM(cast(IDENT as varchar(10))) as [text()] " &
                                   "  FROM " & ParL_HQDataBase & ".[dbo].[Uzytkownicy] " &
                                   "  WHERE (" &
                                   "SELECT COUNT(*) from " & ParL_HQDataBase & ".[dbo].[Uzytkownicy] AS A INNER JOIN " & ParL_DataBase & ".[dbo].[Uzytkownicy] AS B ON A.IDENT=B.IDENT " &
                                   ") > 0 " &
                                   "  For xml path('')"
                        ListaIdent = SQLExecuteScalarTxt(kwerenda, ConnectionStrings(), "")


                        If ListaIdent Is Nothing Then
                        Else
                            If Len(ListaIdent) > 0 Then
                                PiszRaport("")
                                PiszRaport("USUWAM PRACOWNIKÓW O NIEJEDNOZNACZNYCH SYMBOLACH Z KATALOGU PRACOWNIKÓW W BAZIE IMPORTOWANEJ")
                                PiszRaport(ListaIdent)

                                'kwerenda = "DELETE FROM " & ParL_HQDataBase & ".[dbo].[Uzytkownicy] " &
                                '                "WHERE (" &
                                '                "SELECT COUNT(*) from " & ParL_DataBase & ".[dbo].[Uzytkownicy] where (" & ParL_HQDataBase & ".[dbo].[Uzytkownicy].IDENT=" & ParL_DataBase & ".[dbo].[Uzytkownicy].IDENT)" &
                                '                ") >0 "

                                kwerenda = "DELETE FROM " & ParL_HQDataBase & ".[dbo].[Uzytkownicy] " &
                                                "WHERE (" &
                                                "SELECT COUNT(*) from " & ParL_HQDataBase & ".[dbo].[Uzytkownicy] AS A INNER JOIN " & ParL_DataBase & ".[dbo].[Uzytkownicy] AS B ON A.IDENT=B.IDENT " &
                                                ") >0 "


                                WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "Tabela Pracownicy")
                            End If
                        End If
                    End If


                    If WynikOK Then
                        'kwerenda = "SELECT ',' + RTRIM(cast(IDENT as varchar(10))) as [text()] " &
                        '           "  FROM " & ParL_HQDataBase & ".[dbo].[AkcjeOpis] " &
                        '           "  WHERE (" &
                        '           "SELECT COUNT(*) from " & ParL_DataBase & ".[dbo].[AkcjeOpis] where (" & ParL_DataBase & ".[dbo].[AkcjeOpis].IDENT=" & ParL_HQDataBase & ".[dbo].[AkcjeOpis].IDENT) " &
                        '           "        ) > 0 " &
                        '           "  For xml path('')"


                        kwerenda = "SELECT ',' + RTRIM(cast(IDENT as varchar(10))) as [text()] " &
                                   "  FROM " & ParL_HQDataBase & ".[dbo].[AkcjeOpis] " &
                                   "  WHERE (" &
                                   "SELECT COUNT(*) from " & ParL_HQDataBase & ".[dbo].[AkcjeOpis] AS A INNER JOIN " & ParL_DataBase & ".[dbo].[AkcjeOpis] AS B ON A.IDENT=B.IDENT " &
                                   "        ) > 0 " &
                                   "  For xml path('')"

                        ListaIdent = SQLExecuteScalarTxt(kwerenda, ConnectionStrings(), "")


                        If ListaIdent Is Nothing Then
                        Else
                            If Len(ListaIdent) > 0 Then
                                PiszRaport("")
                                PiszRaport("USUWAM AKCJE MARKETINGOWE O NIEJEDNOZNACZNYCH IDENTYFIKATORACH Z KATALOGU AKCJEOPIS I AKCJESPEC W BAZIE IMPORTOWANEJ")
                                PiszRaport(ListaIdent)

                                'kwerenda = "DELETE FROM " & ParL_HQDataBase & ".[dbo].[AkcjeOpis] " &
                                '                "WHERE (" &
                                '                "SELECT COUNT(*) from " & ParL_DataBase & ".[dbo].[AkcjeOpis] where (" & ParL_HQDataBase & ".[dbo].[AkcjeOpis].IDENT=" & ParL_DataBase & ".[dbo].[AkcjeOpis].IDENT) " &
                                '                ") >0 "

                                kwerenda = "DELETE FROM " & ParL_HQDataBase & ".[dbo].[AkcjeOpis] " &
                                                "WHERE (" &
                                                "SELECT COUNT(*) from " & ParL_HQDataBase & ".[dbo].[AkcjeOpis] AS A INNER JOIN " & ParL_DataBase & ".[dbo].[AkcjeOpis] AS B ON A.IDENT=B.IDENT " &
                                                ") >0 "
                                WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "Tabela AkcjeOpis")

                                'kwerenda = "DELETE FROM " & ParL_HQDataBase & ".[dbo].[AkcjeSpec] " &
                                '                "WHERE (" &
                                '                "SELECT COUNT(*) from " & ParL_DataBase & ".[dbo].[AkcjeSpec] where (" & ParL_HQDataBase & ".[dbo].[AkcjeSpec].IDENTAKCJI=" & ParL_DataBase & ".[dbo].[AkcjeSpec].IDENTAKCJI) " &
                                '                ") >0 "
                                kwerenda = "DELETE FROM " & ParL_HQDataBase & ".[dbo].[AkcjeSpec] " &
                                                "WHERE (" &
                                                "SELECT COUNT(*) from " & ParL_HQDataBase & ".[dbo].[AkcjeSpec] AS A INNER JOIN " & ParL_DataBase & ".[dbo].[AkcjeSpec] AS B ON A.IDENTAKCJI=B.IDENTAKCJI " &
                                                ") >0 "
                                WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "Tabela AkcjeSpec")
                            End If
                        End If
                    End If


                End If





                'INSERT INTO Pryzmat_WarszawaKasprowicza.dbo.Uzytkownicy SELECT * FROM Pryzmat_WarszawaJagiellonska.dbo.Uzytkownicy
                If WynikOK Then
                    PiszRaport("")
                    PiszRaport("DOPISUJE TABELE Z BAZY DANYCH IMPORTOWANYCH DO AKTUALNEJ BAZY DANYCH")

                    '--------------------------------------------------------------------
                    If WynikOK Then
                        kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.Uzytkownicy SELECT * FROM " & ParL_HQDataBase & ".dbo.Uzytkownicy"
                        WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabele Uzytkownicy")
                    End If
                    '--------------------------------------------------------------------
                    If WynikOK Then
                        kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.AkcjeOpis SELECT * FROM " & ParL_HQDataBase & ".dbo.AkcjeOpis"
                        WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabela AkcjeOpis")
                    End If
                    '--------------------------------------------------------------------
                    If WynikOK Then
                        kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.AkcjeSpec " &
                                   "SELECT " & ParL_HQDataBase & ".dbo.AkcjeSpec.IDENTAKCJI, " &
                                          "" & ParL_HQDataBase & ".dbo.AkcjeSpec.IDENTKLI, " &
                                          "" & ParL_HQDataBase & ".dbo.AkcjeSpec.WERSJA " &
                                  "FROM " & ParL_HQDataBase & ".dbo.AkcjeSpec INNER JOIN " & ParL_HQDataBase & ".dbo.AkcjeOpis " &
                                  "ON " & ParL_HQDataBase & ".dbo.AkcjeSpec.IDENTAKCJI=" & ParL_HQDataBase & ".dbo.AkcjeOpis.IDENT"
                        WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabela AkcjeSpec")
                    End If











                    '--------------------------------------------------------------------
                    'If WynikOK Then
                    '    kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.AnalizyZakupu SELECT * FROM " & ParL_HQDataBase & ".dbo.AnalizyZakupu"
                    '    WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabela AnalizyZakupu")
                    'End If
                    '--------------------------------------------------------------------
                    'UNIQID + IDENTKLIENTA
                    If WynikOK Then
                        kwerenda = "DELETE FROM " & ParL_HQDataBase & ".[dbo].[CoDalejPlan] WHERE " &
                                                "(SELECT COUNT(*) from " & ParL_DataBase & ".[dbo].[CoDalejPlan] where " &
                                                " (" & ParL_HQDataBase & ".[dbo].[CoDalejPlan].UNIQID=" & ParL_DataBase & ".[dbo].[CoDalejPlan].UNIQID) AND " &
                                                " (" & ParL_HQDataBase & ".[dbo].[CoDalejPlan].IDENTKLIENTA=" & ParL_DataBase & ".[dbo].[CoDalejPlan].IDENTKLIENTA) ) >0 "
                        WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")

                        kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.CoDalejPlan SELECT * FROM " & ParL_HQDataBase & ".dbo.CoDalejPlan"
                        WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabela CoDalejPlan")
                    End If
                    '--------------------------------------------------------------------
                    'If WynikOK Then
                    '    kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.DaneDoRaportu SELECT * FROM " & ParL_HQDataBase & ".dbo.DaneDoRaportu"
                    '    WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabela DaneDoRaportu")
                    'End If
                    '--------------------------------------------------------------------
                    If WynikOK Then
                        kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.Klienci SELECT * FROM " & ParL_HQDataBase & ".dbo.Klienci"
                        WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabela Klienci")
                    End If
                    '--------------------------------------------------------------------
                    If WynikOK Then
                        kwerenda = "DELETE FROM " & ParL_HQDataBase & ".[dbo].[KontaktyHandlowe] WHERE " &
                                                "(SELECT COUNT(*) from " & ParL_DataBase & ".[dbo].[KontaktyHandlowe] where " &
                                                " (" & ParL_HQDataBase & ".[dbo].[KontaktyHandlowe].UNIQID=" & ParL_DataBase & ".[dbo].[KontaktyHandlowe].UNIQID) ) >0 "
                        WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")

                        kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.KontaktyHandlowe SELECT * FROM " & ParL_HQDataBase & ".dbo.KontaktyHandlowe"
                        WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabela KontaktyHandlowe")
                    End If
                    '--------------------------------------------------------------------
                    If WynikOK Then
                        kwerenda = "DELETE FROM " & ParL_HQDataBase & ".[dbo].[LogAktywnosci] WHERE " &
                                                "(SELECT COUNT(*) from " & ParL_DataBase & ".[dbo].[LogAktywnosci] where " &
                                                " (" & ParL_HQDataBase & ".[dbo].[LogAktywnosci].UNIQID=" & ParL_DataBase & ".[dbo].[LogAktywnosci].UNIQID) ) >0 "
                        WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")

                        kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.LogAktywnosci SELECT * FROM " & ParL_HQDataBase & ".dbo.LogAktywnosci"
                        WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabele LogAktywnosci")
                    End If
                    '--------------------------------------------------------------------
                    If WynikOK Then
                        kwerenda = "DELETE FROM " & ParL_HQDataBase & ".[dbo].[Obroty] WHERE " &
                                                "(SELECT COUNT(*) from " & ParL_DataBase & ".[dbo].[Obroty] where " &
                                                " (" & ParL_HQDataBase & ".[dbo].[Obroty].IDENTKLIENTA=" & ParL_DataBase & ".[dbo].[Obroty].IDENTKLIENTA) AND " &
                                                " (" & ParL_HQDataBase & ".[dbo].[Obroty].DATA=" & ParL_DataBase & ".[dbo].[Obroty].DATA) ) >0 "
                        WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")

                        kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.Obroty SELECT * FROM " & ParL_HQDataBase & ".dbo.Obroty"
                        WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabele Obroty")
                    End If
                    '--------------------------------------------------------------------
                    If WynikOK Then
                        kwerenda = "DELETE FROM " & ParL_HQDataBase & ".[dbo].[ObrotyMies] WHERE " &
                                                "(SELECT COUNT(*) from " & ParL_DataBase & ".[dbo].[ObrotyMies] where " &
                                                " (" & ParL_HQDataBase & ".[dbo].[ObrotyMies].IDENTKLIENTA=" & ParL_DataBase & ".[dbo].[ObrotyMies].IDENTKLIENTA) AND " &
                                                " (" & ParL_HQDataBase & ".[dbo].[ObrotyMies].MIESIAC=" & ParL_DataBase & ".[dbo].[ObrotyMies].MIESIAC) ) >0 "
                        WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")

                        kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.ObrotyMies SELECT * FROM " & ParL_HQDataBase & ".dbo.ObrotyMies"
                        WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabele ObrotyMies")
                    End If
                    '--------------------------------------------------------------------
                    If WynikOK Then
                        kwerenda = "DELETE FROM " & ParL_HQDataBase & ".[dbo].[Osoby] WHERE " &
                                                "(SELECT COUNT(*) from " & ParL_DataBase & ".[dbo].[Osoby] where " &
                                                " (" & ParL_HQDataBase & ".[dbo].[Osoby].IDENTKLIENTA=" & ParL_DataBase & ".[dbo].[Osoby].IDENTKLIENTA) AND " &
                                                " (" & ParL_HQDataBase & ".[dbo].[Osoby].OSOBA=" & ParL_DataBase & ".[dbo].[Osoby].OSOBA) ) >0 "
                        WynikOK = SQLCommand(kwerenda, HQConnectionStrings(), "")

                        kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.Osoby SELECT * FROM " & ParL_HQDataBase & ".dbo.Osoby"
                        WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabele Osoby")
                    End If
                    '--------------------------------------------------------------------
                    'If WynikOK Then
                    '    kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.Parametry SELECT * FROM " & ParL_HQDataBase & ".dbo.Parametry"
                    '    WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabela Parametry")
                    'End If
                    '--------------------------------------------------------------------



                    'If WynikOK Then
                    '    kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.RodzajeKontaktow SELECT * FROM " & ParL_HQDataBase & ".dbo.RodzajeKontaktow"
                    '    WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabele RodzajeKontaktow")
                    'End If
                    '--------------------------------------------------------------------
                    'If WynikOK Then
                    '    kwerenda = "INSERT INTO " & ParL_DataBase & ".dbo.SlownikCoDalej SELECT * FROM " & ParL_HQDataBase & ".dbo.SlownikCoDalej"
                    '    WynikOK = SQLCommand(kwerenda, ConnectionStrings(), "Tabele SlownikCoDalej")
                    'End If
                    '--------------------------------------------------------------------
                End If

                '---------------------------------
                PiszRaport("")
                PiszRaport("OK, Skoñczy³em.")




                If WynikOK Then
                    MessageBox.Show("Po³¹czy³em Bazy Danych" & vbCrLf &
                           "Dopisa³em informacje z bazy importowanej " & ParL_HQDataBase & vbCrLf &
                           "do aktualnej Bazy Danych " & ParL_DataBase,
                        "OK, skoñczy³em",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                Else
                    MessageBox.Show("NIESTETY nie po³¹czy³em Baz Danych" & vbCrLf &
                           "W trakcie ³¹czenia wyst¹pi³ B£AD." & vbCrLf &
                           "Konieczne jest odtworzenie obu baz danych z ARCHIWUM...",
                        "OK, skoñczy³em",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1)
                End If



            Catch Ex As System.Exception
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try


        Else
            'NIE IMPORTUJEMY
        End If




        ''---------------------------------
        'Pobierz_Tabele("Klienci", "SELECT * FROM Klienci")
        'Pobierz_Tabele("Kontakty", "SELECT * FROM Kontakty")
        'Pobierz_Tabele("Osoby", "SELECT * FROM Osoby")
        'Pobierz_Tabele("Obroty", "SELECT * FROM Obroty")
        'Pobierz_Tabele("ObrotyMies", "SELECT * FROM ObrotyMies")
        'Pobierz_Tabele("CoDalejPlan", "SELECT * FROM CoDalejPlan")


    End Sub




    Private Sub PiszRaport(ByVal komunikat As String)
        txtRaport.Text &= SysTime() & "  " & komunikat & vbCrLf
        txtRaport.SelectionStart = Len(txtRaport.Text)
        txtRaport.SelectionLength = 0
        txtRaport.ScrollToCaret()
        System.Windows.Forms.Application.DoEvents()
    End Sub



    '*******************************************
    ' obs³uga bazy SQL
    '*******************************************

    Private Function SQLCommand(ByVal parKwe As String,
                                ByVal parConnectionStr As String,
                                ByVal parKomunikat As String) As Boolean
        Dim Wynik As Integer = 0
        Dim sqlConnection As SqlClient.SqlConnection
        Dim sqlCmdd As SqlClient.SqlCommand
        '---------------------------------
        If Len(parKomunikat) > 0 Then PiszRaport(parKomunikat)

        sqlConnection = New SqlClient.SqlConnection(parConnectionStr)
        sqlCmdd = New SqlClient.SqlCommand
        Try
            sqlCmdd.Connection = sqlConnection
            sqlCmdd.CommandType = CommandType.Text
            sqlCmdd.CommandText = parKwe
            sqlConnection.Open()
            Wynik = sqlCmdd.ExecuteNonQuery()
            If Wynik = -1 Then
            End If
            'txtRaport.Text &= "OK, aktualizacja wykonana poprawnie" & vbCrLf

        Catch ex As SqlClient.SqlException
            'MessageBox.Show(ex.Message)
            PiszRaport("B³¹d wykonania aktualizacji tabeli : " & vbCrLf & ex.Message)

            Return (False)
        Finally
            sqlConnection.Close()
        End Try
        sqlCmdd.Dispose()
        sqlConnection.Dispose()
        Return (True)
    End Function


    Public Function SQLExecuteScalarTxt(ByVal parKwe As String,
                                        ByVal parConnectionStr As String,
                                        ByVal parKomunikat As String) As String
        Dim sqlConnection As SqlClient.SqlConnection
        Dim sqlCommand As SqlClient.SqlCommand
        Dim Wynik As String = ""
        '---------------------------------
        If Len(parKomunikat) > 0 Then PiszRaport(parKomunikat)

        sqlConnection = New SqlClient.SqlConnection(parConnectionStr)
        sqlCommand = New SqlClient.SqlCommand
        Try
            sqlCommand.Connection = sqlConnection
            sqlCommand.CommandType = CommandType.Text
            sqlCommand.CommandText = parKwe
            sqlConnection.Open()
            Wynik = sqlCommand.ExecuteScalar()

        Catch ex As SqlClient.SqlException
            'MessageBox.Show(ex.Message)
            PiszRaport("B³¹d wykonania : " & vbCrLf & ex.Message)
            Return ""
        Finally
            sqlConnection.Close()
        End Try
        sqlCommand.Dispose()
        sqlConnection.Dispose()
        Return Wynik
    End Function



















    Private Sub Pobierz_Tabele(ByVal NazwaTabeli As String,
                               ByVal SelectTabeli As String)
        Dim RText As String = ""
        Dim SRCConnString As String = HQConnectionStrings()
        Dim DSTConnString As String = ConnectionStrings()

        txtRaport.Text &= "Pobieram dane z tabeli " & NazwaTabeli & "...  "
        txtRaport.SelectionStart = Len(txtRaport.Text)
        txtRaport.SelectionLength = 0
        txtRaport.ScrollToCaret()
        System.Windows.Forms.Application.DoEvents()

        If Kopiuj_Tabele(NazwaTabeli,
                      SRCConnString,
                      SelectTabeli,
                      DSTConnString,
                      RText) Then
            txtRaport.Text &= "Pobrano " & RText & " rekordów." & vbCrLf
        Else
            txtRaport.Text &= "B³¹d podczas pobierania." & vbCrLf & "-----> " & RText & vbCrLf
        End If
        txtRaport.SelectionStart = Len(txtRaport.Text)
        txtRaport.SelectionLength = 0
        txtRaport.ScrollToCaret()
        System.Windows.Forms.Application.DoEvents()
    End Sub




    Private Function Kopiuj_Tabele(ByVal pTabela As String,
                          ByVal pSrcConnectionString As String,
                          ByVal pSrcSelect As String,
                          ByVal pDstConnectionString As String,
                          ByRef RetTxt As String) As Boolean
        Dim Ret As Boolean = True

        ' Open a connection to the AdventureWorks database. 
        Using sourceConnection As SqlClient.SqlConnection = New SqlClient.SqlConnection(pSrcConnectionString)
            sourceConnection.Open()

            ' Get data from the source table as a SqlDataReader. 
            Dim commandSourceData As SqlClient.SqlCommand = New SqlClient.SqlCommand(pSrcSelect, sourceConnection)
            Dim reader As SqlClient.SqlDataReader = commandSourceData.ExecuteReader

            Using destinationConnection As SqlClient.SqlConnection = New SqlClient.SqlConnection(pDstConnectionString)

                ''usun wszystkie rekordy z tabeli
                'Dim cmd As SqlClient.SqlCommand = New SqlClient.SqlCommand("DELETE FROM " & pTabela, destinationConnection)
                'destinationConnection.Open()
                'cmd.ExecuteNonQuery()

                destinationConnection.Open()

                ' Set up the bulk copy object.  
                ' The column positions in the source data reader  
                ' match the column positions in the destination table,  
                ' so there is no need to map columns. 
                Using bulkCopy As SqlClient.SqlBulkCopy = New SqlClient.SqlBulkCopy(destinationConnection)
                    bulkCopy.DestinationTableName = pTabela

                    Try

                        ' Write from the source to the destination.
                        bulkCopy.WriteToServer(reader)

                        Dim RowCount As New SqlClient.SqlCommand("SELECT COUNT(*) FROM " & pTabela, destinationConnection)
                        RetTxt = System.Convert.ToInt32(RowCount.ExecuteScalar()).ToString

                    Catch ex As Exception
                        'Console.WriteLine(ex.Message)
                        Ret = False
                        RetTxt = ex.Message


                    Finally
                        ' Close the SqlDataReader. The SqlBulkCopy 
                        ' object is automatically closed at the end 
                        ' of the Using block.
                        reader.Close()
                    End Try
                End Using
            End Using
        End Using
        Return Ret
    End Function




End Class
