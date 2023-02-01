'Imports System.Data.Odbc
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




Public Class GeneratorWydrukow

    Private Sub GeneratorWydrukow_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        DefListePol()
        txtNazwa.Text = "Zestawienie klientów."
        txtSchemat.Text = ""
        lblPath.Text = ""
        lblSzerokoscArkusza.Text = 290.ToString("N0")
        lblSzerokoscWydruku.Text = 0.ToString("N0")
        oGeneratorWydruku = ""
    End Sub

    Private Sub GeneratorWydrukow_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub

    Private Sub DefListePol()
        'ListaPol.Clear()
        ListaPol.Visible = True
        ListaPol.Enabled = True
        ListaPol.ForeColor = System.Drawing.Color.Purple
        'zdefiniuj poszczegolne pola w bazie KLIENCI
        '-------------------------------------------------
        If _DataSet.Tables(0).Columns.Count() > 0 Then
            Dim ic As Integer
            For ic = 0 To _DataSet.Tables(0).Columns.Count() - 1
                DefItem(_DataSet.Tables(0).Columns(ic).ColumnName, _
                        _DataSet.Tables(0).Columns(ic).DataType.ToString)
            Next
        End If
        DefItem("KONTAKTYHANDLOWE", "System.String")
        DefItem("CODALEJ", "System.String")
    End Sub


    Private Sub DefItem(ByVal NazwaPola As String, _
                        ByVal TypPola As String)
        Dim Item00 As New ListViewItem(NazwaPola)
        Item00.SubItems.Add(TypPola)
        Item00.SubItems.Add("")
        'Item00.BackColor = bColor
        'Item00.ForeColor = System.Drawing.Color.Purple
        ListaPol.Items.Add(Item00)
    End Sub

    '****************************************
    '** definiuj klawisze
    '****************************************

    Private Sub cmdDodaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDodaj.Click
        If ListaPol.Items.Count > 0 Then
            If ListaPol.SelectedIndices.Count > 0 Then
                'dopisz na koniec listy
                Dim Index As Integer = ListaPol.SelectedIndices.Item(0)
                Dim Item00 As New ListViewItem(ListaPol.Items(Index).SubItems(0).Text)
                Item00.SubItems.Add(ListaPol.Items(Index).SubItems(1).Text)
                Item00.SubItems.Add("20")
                Item00.SubItems.Add(ListaPol.Items(Index).SubItems(0).Text)
                Item00.SubItems.Add("Do lewej")
                Item00.SubItems.Add("1")
                ListaWydruku.Items.Add(Item00)

                lblSzerokoscWydruku.Text = WyliczSzerokoscWydruku().ToString("N0")
            End If
        End If
    End Sub

    Private Sub cmdUsun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsun.Click
        If ListaWydruku.Items.Count > 0 Then
            If ListaWydruku.SelectedIndices.Count > 0 Then
                'usun wskazany zapis
                Dim Index As Integer = ListaWydruku.SelectedIndices.Item(0)
                If Index >= 0 Then
                    ListaWydruku.Items(Index).Remove()
                    lblSzerokoscWydruku.Text = WyliczSzerokoscWydruku().ToString("N0")
                End If
            End If
        End If
    End Sub

    Private Sub cmdGora_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGora.Click
        If ListaWydruku.Items.Count > 0 Then
            If ListaWydruku.SelectedIndices.Count > 0 Then
                'usun wskazany zapis
                Dim Index As Integer = ListaWydruku.SelectedIndices.Item(0)
                If Index > 0 Then
                    'zamien m iejscami 
                    Dim Nazwa As String = ListaWydruku.Items(Index).SubItems(0).Text
                    Dim Typ As String = ListaWydruku.Items(Index).SubItems(1).Text
                    Dim Rozmiar As String = ListaWydruku.Items(Index).SubItems(2).Text
                    Dim Naglowek As String = ListaWydruku.Items(Index).SubItems(3).Text
                    Dim Wyrownanie As String = ListaWydruku.Items(Index).SubItems(4).Text
                    Dim MaxIlWierszy As String = ListaWydruku.Items(Index).SubItems(5).Text

                    ListaWydruku.Items(Index).SubItems(0).Text = ListaWydruku.Items(Index - 1).SubItems(0).Text
                    ListaWydruku.Items(Index).SubItems(1).Text = ListaWydruku.Items(Index - 1).SubItems(1).Text
                    ListaWydruku.Items(Index).SubItems(2).Text = ListaWydruku.Items(Index - 1).SubItems(2).Text
                    ListaWydruku.Items(Index).SubItems(3).Text = ListaWydruku.Items(Index - 1).SubItems(3).Text
                    ListaWydruku.Items(Index).SubItems(4).Text = ListaWydruku.Items(Index - 1).SubItems(4).Text
                    ListaWydruku.Items(Index).SubItems(5).Text = ListaWydruku.Items(Index - 1).SubItems(5).Text

                    ListaWydruku.Items(Index - 1).SubItems(0).Text = Nazwa
                    ListaWydruku.Items(Index - 1).SubItems(1).Text = Typ
                    ListaWydruku.Items(Index - 1).SubItems(2).Text = Rozmiar
                    ListaWydruku.Items(Index - 1).SubItems(3).Text = Naglowek
                    ListaWydruku.Items(Index - 1).SubItems(4).Text = Wyrownanie
                    ListaWydruku.Items(Index - 1).SubItems(5).Text = MaxIlWierszy
                End If
            End If
        End If
    End Sub

    Private Sub cmdDol_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDol.Click
        If ListaWydruku.Items.Count > 0 Then
            If ListaWydruku.SelectedIndices.Count > 0 Then
                'usun wskazany zapis
                Dim Index As Integer = ListaWydruku.SelectedIndices.Item(0)
                If Index < ListaWydruku.Items.Count - 1 Then
                    'zamien m iejscami 
                    Dim Nazwa As String = ListaWydruku.Items(Index).SubItems(0).Text
                    Dim Typ As String = ListaWydruku.Items(Index).SubItems(1).Text
                    Dim Rozmiar As String = ListaWydruku.Items(Index).SubItems(2).Text
                    Dim Naglowek As String = ListaWydruku.Items(Index).SubItems(3).Text
                    Dim Wyrownanie As String = ListaWydruku.Items(Index).SubItems(4).Text
                    Dim MaxIlWierszy As String = ListaWydruku.Items(Index).SubItems(5).Text

                    ListaWydruku.Items(Index).SubItems(0).Text = ListaWydruku.Items(Index + 1).SubItems(0).Text
                    ListaWydruku.Items(Index).SubItems(1).Text = ListaWydruku.Items(Index + 1).SubItems(1).Text
                    ListaWydruku.Items(Index).SubItems(2).Text = ListaWydruku.Items(Index + 1).SubItems(2).Text
                    ListaWydruku.Items(Index).SubItems(3).Text = ListaWydruku.Items(Index + 1).SubItems(3).Text
                    ListaWydruku.Items(Index).SubItems(4).Text = ListaWydruku.Items(Index + 1).SubItems(4).Text
                    ListaWydruku.Items(Index).SubItems(5).Text = ListaWydruku.Items(Index + 1).SubItems(5).Text

                    ListaWydruku.Items(Index + 1).SubItems(0).Text = Nazwa
                    ListaWydruku.Items(Index + 1).SubItems(1).Text = Typ
                    ListaWydruku.Items(Index + 1).SubItems(2).Text = Rozmiar
                    ListaWydruku.Items(Index + 1).SubItems(3).Text = Naglowek
                    ListaWydruku.Items(Index + 1).SubItems(4).Text = Wyrownanie
                    ListaWydruku.Items(Index + 1).SubItems(5).Text = MaxIlWierszy
                End If
            End If
        End If
    End Sub

    Private Sub ListaPol_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListaPol.DoubleClick
        cmdDodaj_Click(sender, e)
    End Sub

    '*****************************************
    '** Edycja pola na wydruku
    '*****************************************

    Private Sub ListaWydruku_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListaWydruku.DoubleClick
        If ListaWydruku.Items.Count > 0 Then
            If ListaWydruku.SelectedIndices.Count > 0 Then
                Dim Index As Integer = ListaWydruku.SelectedIndices.Item(0)
                GenNazwa = ListaWydruku.Items(Index).SubItems(0).Text
                GenTyp = ListaWydruku.Items(Index).SubItems(1).Text
                GenRozmiar = ListaWydruku.Items(Index).SubItems(2).Text
                GenNaglowek = ListaWydruku.Items(Index).SubItems(3).Text
                GenWyrownanie = ListaWydruku.Items(Index).SubItems(4).Text
                GenMaxIlWierszy = ListaWydruku.Items(Index).SubItems(5).Text

                Dim Form As New EdytujGeneratorWydrukow
                Form.ShowDialog()

                ListaWydruku.Items(Index).SubItems(0).Text = GenNazwa
                ListaWydruku.Items(Index).SubItems(1).Text = GenTyp
                ListaWydruku.Items(Index).SubItems(2).Text = GenRozmiar
                ListaWydruku.Items(Index).SubItems(3).Text = GenNaglowek
                ListaWydruku.Items(Index).SubItems(4).Text = GenWyrownanie
                ListaWydruku.Items(Index).SubItems(5).Text = GenMaxIlWierszy

                lblSzerokoscWydruku.Text = WyliczSzerokoscWydruku().ToString("N0")
            End If
        End If
    End Sub

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        oGeneratorWydruku = ZdefiniujSzablonWydruku()
        '----------------------------
        Me.Close()
    End Sub

    Private Sub cmdZrezygnuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdZrezygnuj.Click
        oGeneratorWydruku = ""
        Me.Close()
    End Sub

    Private Function WyliczSzerokoscWydruku() As Integer
        Dim ij As Integer = 0
        Dim Szer As Integer = 0
        For ij = 0 To ListaWydruku.Items.Count - 1
            Szer += CInt(ListaWydruku.Items(ij).SubItems(2).Text)
        Next
        Return Szer
    End Function

    Private Function ZdefiniujSzablonWydruku() As String
        Dim ij As Integer = 0
        Dim szablon As String = Trim(txtNazwa.Text) & vbCrLf
        For ij = 0 To ListaWydruku.Items.Count - 1
            'Opis wydruku=NAZWAPOLA|TYPPOLA|SZEROKOSC|NAGLOWEK|WYROWNANIE|MAXILWIERSZY<crlf>
            szablon += ListaWydruku.Items(ij).SubItems(0).Text & "|" & _
                       ListaWydruku.Items(ij).SubItems(1).Text & "|" & _
                       ListaWydruku.Items(ij).SubItems(2).Text & "|" & _
                       ListaWydruku.Items(ij).SubItems(3).Text & "|" & _
                       ListaWydruku.Items(ij).SubItems(4).Text & "|" & _
                       ListaWydruku.Items(ij).SubItems(5).Text & vbCrLf
        Next
        Return szablon
    End Function

    Private Sub txtSchemat_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSchemat.GotFocus
        StartEdycjiTxt(txtSchemat)
    End Sub
    Private Sub txtSchemat_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSchemat.LostFocus
        KoniecEdycjiTxt(txtSchemat)
    End Sub
    Private Sub txtSchemat_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSchemat.TextChanged
        TestLen(txtSchemat, "NAZWA SZABLONU", 10)
    End Sub

    Private Sub cmdCzysc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCzysc.Click
        If ListaWydruku.Items.Count > 0 Then
            Dim i As Integer
            For i = ListaWydruku.Items.Count - 1 To 0 Step -1
                ListaWydruku.Items(i).Remove()
            Next
            lblSzerokoscWydruku.Text = 0.ToString("N0")
        End If
        txtNazwa.Text = ""
    End Sub

    Private Sub cmdZapisz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdZapisz.Click
        If Len(Trim(txtSchemat.Text)) = 0 Then
            MessageBox.Show("Proszê okreœliæ nazwê szablonu...", _
                "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            oGeneratorWydruku = ZdefiniujSzablonWydruku()
            ZapiszSzablon(Trim(txtSchemat.Text) & ".sw", oGeneratorWydruku)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        With OpenFileDialog1
            .Title = "Proszê wybraæ plik ze schematem wydruku"
            .InitialDirectory = oKatParametry
            .DefaultExt = "sw"
            .Filter = "Schemat wydruku (*.sw)|*.sw|Wszystkie pliki|*.*"

            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                lblPath.Text = .FileName

                'oddziel plik od katalogow
                Dim Plik As String = .FileName
                Dim ChPos As String
                ChPos = InStr(Plik, "\")
                Do While ChPos > 0
                    Plik = Mid(Plik, ChPos + 1)
                    ChPos = InStr(Plik, "\")
                Loop

                'oddziel rozszerzenie pliku...
                ChPos = InStr(Plik, ".")
                txtSchemat.Text = Mid(Plik, 1, ChPos - 1)

                'pobierz schemat
                'PlikZSzablonem = Plik
                'PobierzSzablon()
                'AktListePol()
            End If
        End With

    End Sub

    Private Sub cmdPobierz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPobierz.Click
        If Len(Trim(txtSchemat.Text)) = 0 Then
            MessageBox.Show("Proszê okreœliæ nazwê szablonu...", _
                "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            PobierzSzablon(Trim(txtSchemat.Text) & ".sw")
        End If
    End Sub


    Private Sub ZapiszSzablon(ByVal PlikZSzablonem As String, ByVal Szablon As String)
        Dim FileNum As Integer
        Dim pos As Integer
        Dim txt As String

        FileNum = FreeFile()    'kolejny nr otwartego zbioru
        FileOpen(FileNum, oKatParametry & "\" & PlikZSzablonem, OpenMode.Output)

        PrintLine(FileNum, "SOFTART : Szablon wydruku z katalogu klientów")
        pos = InStr(Szablon, vbCrLf)
        Do While pos > 0
            txt = Mid(Szablon, 1, pos - 1)
            PrintLine(FileNum, txt)
            Szablon = Mid(Szablon, pos + 2)
            pos = InStr(Szablon, vbCrLf)
        Loop
        PrintLine(FileNum, "Koniec szablonu")
        FileClose(FileNum)
        '-------------------------------
        MessageBox.Show("Zapisa³em szablon wydruku do pliku dyskowego" & vbCrLf & _
                oKatParametry & "\" & PlikZSzablonem, _
                "Ok, skoñczy³em.", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
    End Sub

    Private Sub PobierzSzablon(ByVal PlikZSzablonem As String)
        Dim FileNum As Integer
        Dim NextLine As String
        Dim pos As Integer = 0
        Dim kol As String = ""

        Dim N As String = ""
        Dim T As String = ""
        Dim R As String = ""
        Dim H As String = ""
        Dim W As String = ""
        Dim I As String = ""

        'wyczyœæ szablon
        If ListaWydruku.Items.Count > 0 Then
            Dim ii As Integer
            For ii = ListaWydruku.Items.Count - 1 To 0 Step -1
                ListaWydruku.Items(ii).Remove()
            Next
        End If

        If Len(Dir(oKatParametry & "\" & PlikZSzablonem)) > 0 Then
            FileNum = FreeFile()    'kolejny nr otwartego zbioru
            FileOpen(FileNum, oKatParametry & "\" & PlikZSzablonem, OpenMode.Input)
            If Not EOF(FileNum) Then
                NextLine = LineInput(FileNum)

                'sprawdz czy to plik z parametrami programu
                If InStr(NextLine, "SOFTART : Szablon wydruku z katalogu klientów") = 1 Then
                    'na poczatku jest naglowek
                    NextLine = LineInput(FileNum)
                    txtNazwa.Text = NextLine
                    'w kolejnych wierszach sa opisane kolejne kolumny wydruku
                    Do Until EOF(FileNum)
                        NextLine = LineInput(FileNum)
                        If NextLine = "Koniec szablonu" Then
                        Else
                            kol = NextLine

                            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
                            N = Mid(kol, 1, pos - 1)
                            kol = Mid(kol, pos + 1)

                            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
                            T = Mid(kol, 1, pos - 1)
                            kol = Mid(kol, pos + 1)

                            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
                            R = Mid(kol, 1, pos - 1)
                            kol = Mid(kol, pos + 1)

                            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
                            H = Mid(kol, 1, pos - 1)
                            kol = Mid(kol, pos + 1)

                            pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
                            W = Mid(kol, 1, pos - 1)
                            I = Mid(kol, pos + 1)

                            'dopisz na koniec listy
                            Dim Item00 As New ListViewItem(N)
                            Item00.SubItems.Add(T)
                            Item00.SubItems.Add(R)
                            Item00.SubItems.Add(H)
                            Item00.SubItems.Add(W)
                            Item00.SubItems.Add(I)
                            ListaWydruku.Items.Add(Item00)
                        End If
                    Loop
                End If
                FileClose(FileNum)
            End If
        End If
        lblSzerokoscWydruku.Text = WyliczSzerokoscWydruku().ToString("N0")
    End Sub

    '**********************************
    '*** przepisujemy wydruk do Excella
    '**********************************

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
    '"provider=Microsoft.Jet.OLEDB.4.0; " & _
    '"data source=" & oPlikNaDysku & "; " & _
    '"Extended Properties='Excel 8.0; HDR=YES; IMEX=1';"


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

    Private Sub DOExcela()
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
        Dim EXCELCon As New OleDb.OleDbConnection(EXCELConnectionOLeDB)
        Dim EXCELCommandSelect As New OleDbCommand("SELECT * FROM [Arkusz1$]", EXCELCon)
        Dim EXCELCommandInsert As New OleDbCommand("SELECT * FROM [Arkusz1$]", EXCELCon)
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


            'definiuj DataView
            EXCELdv = New DataView(EXCELds.Tables(0))


        End If
    End Sub





    '**********************************************
    '** przepisz do EZCELa
    '**********************************************

    Private Sub cmdDoExcela_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        oGeneratorWydruku = ZdefiniujSzablonWydruku()
        '---------------------------------
        Dim Text As String = ""
        Dim ik As Integer = 0
        Dim ileKolumn As Integer = GWIleKolumn(oGeneratorWydruku)
        Dim Kolumny As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


        Dim N As String = ""
        Dim T As String = ""
        Dim R As Integer = 0
        Dim H As String = ""
        Dim W As String = ""
        Dim I As Integer = 0

        '---------------------------------
        Dim app As Application
        app = New Application()
        If app Is Nothing Then
            MessageBox.Show("Nie mogê uruchomiæ program EXCEL", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            app.Visible = True      'EXCEL WIDOCZNY NA EKRANIE

            Dim workbooks As Workbooks      'Getting the workbooks collection
            workbooks = app.Workbooks

            Dim workbook As _Workbook       'adding a new workbook
            workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)

            Dim sheets As Sheets            'Getting the worksheets collection
            sheets = workbook.Worksheets

            Dim worksheet As _Worksheet     'adding a new worksheet
            worksheet = sheets.Item(1)
            If worksheet Is Nothing Then
                MessageBox.Show("Nie mogê dodaæ nowego arkusza EXCEL", "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            End If
            worksheet.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnRight
            worksheet.Columns.ColumnWidth = 100

            '--------------------------

            ' Naglowek strony - drukuj tytul wydruku
            Text = GWNaglowek(oGeneratorWydruku)
            WpiszDoArkusza(worksheet, "C1", Text, 100, "Centralnie")

            For ik = 1 To IleKolumn
                N = GWKolNazwa(oGeneratorWydruku, ik)
                T = GWKolTyp(oGeneratorWydruku, ik)
                R = GWKolRozmiar(oGeneratorWydruku, ik)
                H = GWKolNaglowek(oGeneratorWydruku, ik)
                W = GWKolWyrownanie(oGeneratorWydruku, ik)
                I = GWKolIleWierszy(oGeneratorWydruku, ik)

                'WpiszDoArkusza(worksheet, Mid(Kolumny, ik, 1) & Trim(Str(ik)), H, R, W)
                WpiszDoArkusza(worksheet, Mid(Kolumny, ik, 1) & "3", H, R, W)
            Next

            If _DataView.Count > 0 Then
                Dim ir As Integer = 0
                Dim dr As DataRowView
                For ir = 0 To _DataView.Count - 1
                    dr = _DataView.Item(ir)

                    For ik = 1 To ileKolumn
                        N = GWKolNazwa(oGeneratorWydruku, ik)
                        T = GWKolTyp(oGeneratorWydruku, ik)
                        R = GWKolRozmiar(oGeneratorWydruku, ik)
                        H = GWKolNaglowek(oGeneratorWydruku, ik)
                        W = GWKolWyrownanie(oGeneratorWydruku, ik)
                        I = GWKolIleWierszy(oGeneratorWydruku, ik)

                        Text = dr.Item(N)
                        WpiszDoArkusza(worksheet, Mid(Kolumny, ik, 1) & Trim(Str(4 + ir)), Text, R, W)
                    Next
                Next
            End If

            'wybierz góre arkusza...
            Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
            range = worksheet.Range("A1", Missing.Value)
            range.Select()

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

    Private Sub WpiszDoArkusza(ByVal Arkusz As _Worksheet, _
                               ByVal Adres As String, _
                               ByVal Zawartosc As String, _
                               ByVal Szerokosc As Integer, _
                               ByVal Wyrownanie As String)
        Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
        Dim ColRange As Microsoft.Office.Interop.Excel.Range
        'Dim RowRange As Microsoft.Office.Interop.Excel.Range

        range = Arkusz.Range(Adres, Missing.Value)
        If range Is Nothing Then
            MessageBox.Show("Nie mogê zdefiniowaæ zakresu arkusza EXCEL", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            range.Value2 = Zawartosc
            range.ColumnWidth = Int(Szerokosc / 2.125)
            range.WrapText = True

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

    '*****************************************
    '** edycja graficzna tabeli
    '*****************************************

    Private Sub cmdEdytuj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdytuj.Click
        oGeneratorWydruku = ZdefiniujSzablonWydruku()
        '-------------------------------
        Dim Form As New frmEdytujTabele(oGeneratorWydruku)
        Form.ShowDialog()
        '-------------------------------
        Dim pos As Integer = 0
        Dim kol As String = ""

        Dim N As String = ""
        Dim T As String = ""
        Dim R As String = ""
        Dim H As String = ""
        Dim W As String = ""
        Dim I As String = ""

        'wyczyœæ szablon
        If ListaWydruku.Items.Count > 0 Then
            Dim ii As Integer
            For ii = ListaWydruku.Items.Count - 1 To 0 Step -1
                ListaWydruku.Items(ii).Remove()
            Next
        End If

        pos = InStr(oGeneratorWydruku, vbCrLf)
        If pos > 0 Then
            txtNazwa.Text = Mid(oGeneratorWydruku, 1, pos - 1)
            oGeneratorWydruku = Mid(oGeneratorWydruku, pos + 2)
            pos = InStr(oGeneratorWydruku, vbCrLf)
            Do While pos > 0
                kol = Mid(oGeneratorWydruku, 1, pos - 1)
                oGeneratorWydruku = Mid(oGeneratorWydruku, pos + 2)

                pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
                N = Mid(kol, 1, pos - 1)
                kol = Mid(kol, pos + 1)

                pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK
                T = Mid(kol, 1, pos - 1)
                kol = Mid(kol, pos + 1)

                pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
                R = Mid(kol, 1, pos - 1)
                kol = Mid(kol, pos + 1)

                pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
                H = Mid(kol, 1, pos - 1)
                kol = Mid(kol, pos + 1)

                pos = InStr(kol, "|")   'NAZWA|TYP|ROZMIAR|NAGLOWEK|WYROWNANIE|ILEWIERSZY
                W = Mid(kol, 1, pos - 1)
                I = Mid(kol, pos + 1)

                'dopisz na koniec listy
                Dim Item00 As New ListViewItem(N)
                Item00.SubItems.Add(T)
                Item00.SubItems.Add(R)
                Item00.SubItems.Add(H)
                Item00.SubItems.Add(W)
                Item00.SubItems.Add(I)
                ListaWydruku.Items.Add(Item00)

                pos = InStr(oGeneratorWydruku, vbCrLf)
            Loop
        End If

        lblSzerokoscWydruku.Text = WyliczSzerokoscWydruku().ToString("N0")
    End Sub

End Class