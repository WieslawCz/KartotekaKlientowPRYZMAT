Module modParametry
    Public KatParametryMDB As String = oPlikKartoteki       'Parametry programu
    Public KatParametryBD As String = oBazaDanych
    Public KatParametryTabela As String = "Parametry"





    'parametry lokalne programu
    Public ParL_OpisInstalacji As String = "" 'opis biezacej instalacji
    Public ParL_NazwaArchiwum As String = "" 'opis biezacej instalacji
    Public ParL_KatalogArchiwum As String = "" 'opis biezacej instalacji
    Public ParL_KatalogArchiwumZIP As String = "" 'opis biezacej instalacji
    Public ParL_KatalogRaporty As String = "" 'opis biezacej instalacji
    Public ParL_DataBaseType As String = "MS ACCESS"   'typ bazy danych ACCESS/MSSQL

    Public ParL_eMailService As String = ""
    Public ParL_eMailThunderbirdPath As String = ""
    Public ParL_eMailAdres As String = ""
    Public ParL_SMTP As String = ""
    Public ParL_SMTPPort As String = ""
    Public ParL_POP3 As String = ""
    Public ParL_POP3Port As String = ""
    Public ParL_eMailUser As String = ""
    Public ParL_eMailPass As String = ""
    Public ParL_SSLProtocol As String = ""
    'parametry wspolpracy z baz¹ danych ACCESS
    Public ParL_KatZbiorow As String = ""     'katalog z plikami danych
    'parametry wspolpracy z baza danych SQL
    Public ParL_Autentykacja As String = "" 'sposób autentykacji
    Public ParL_Serwer As String = "" 'opis biezacej instalacji
    Public ParL_Login As String = "" 'opis biezacej instalacji
    Public ParL_Password As String = "" 'opis biezacej instalacji
    Public ParL_DataBase As String = "" 'opis biezacej instalacji

    '--------------------- Parametry dla Bazy do Importu danych
    Public ParL_HQ As String = "NIE"
    Public ParL_HQDataBaseType As String = "MS SQL"   'typ bazy danych ACCESS/MSSQL
    'parametry wspolpracy z baz¹ danych ACCESS
    Public ParL_HQPlikMDB As String = ""     'kataLOG I plik danych ACCES
    'parametry wspolpracy z baza danych SQL
    Public ParL_HQAutentykacja As String = "" 'sposób autentykacji
    Public ParL_HQSerwer As String = "" 'opis biezacej instalacji
    Public ParL_HQLogin As String = "" 'opis biezacej instalacji
    Public ParL_HQPassword As String = "" 'opis biezacej instalacji
    Public ParL_HQDataBase As String = "" 'opis biezacej instalacji



    'parametry programu
    Public Par_IdentUzytkownika As String = ""
    Public Par_NazwaUzytkownika As String = ""
    Public Par_AdresUzytkownika As String = ""
    Public Par_KontoUzytkownika As String = ""
    Public Par_BankUzytkownika As String = ""
    Public Par_MiejscowoscUzytkownika As String = ""
    Public Par_NIPUzytkownika As String = ""
    Public Par_IdentOddzialu As String = ""

    Public Par_DataAktAnalizy As String = ""
    Public Par_OstOkresAnalizy As String = ""
    Public Par_DaneDoAnalizy As String = ""
    Public Par_MiesAnalizy1 As String = ""
    Public Par_MiesAnalizy2 As String = ""
    Public Par_DaneDoPrognozy As String = ""
    Public Par_MiesPrognozy As String = ""

    Public Par_KatalogOferty As String = ""
    Public Par_wwwPAYBACK As String = ""

    Public Par_RAKolumna01 As String = ""
    Public Par_RAKolumna02 As String = ""
    Public Par_RAKolumna03 As String = ""
    Public Par_RAKolumna04 As String = ""
    Public Par_RAKolumna05 As String = ""
    Public Par_RAKolumna06 As String = ""
    Public Par_RAKolumna07 As String = ""
    Public Par_RAKolumna08 As String = ""
    Public Par_RAKolumna09 As String = ""
    Public Par_RAKolumna10 As String = ""

    Public Par_WartGranWartosc As Double = 0
    Public Par_WartGranProcent As Double = 0

    Public Par_WartGranMatWartosc As Double = 0
    Public Par_WartGranMatProcent As Double = 0

    Public Par_Wersja As Integer = 1











    'zmienne globalne wykorzystywane w programie
    'zmienne globalne wykorzystywane w programie
    Public oKatProgramu As String = ""       'katalog w którym zainstalowano program
    Public oKatRoboczy As String = ""        'katalog roboczy
    Public oKatParametry As String = ""      'katalog w którym powinny znajdowac sie parametry
    Public oPlikParametry As String = ""     'plik z parametrami
    Public oNazwaProgramu As String = ""


    Public oPlikSzerokosciKolumnKartoteka As String = ""     'plik z parametrami
    Public oPlikSzerokosciKolumnAnaliza As String = ""     'plik z parametrami
    Public oPlikSzerokosciKolumnKontakty As String = ""     'plik z parametrami
    Public oPlikSzerokosciKolumnKlienciICoDalej As String = ""     'plik z parametrami

    Public oLocalHostName As String = ""     'nazwa komputera
    Public oLocalHostIP As String = ""       'adres IP 

    Public Sub DefiniujZmienneGlobalne()


        Dim ExecutingApp As System.Reflection.Assembly
        ExecutingApp = System.Reflection.Assembly.GetExecutingAssembly()

        Dim Name As System.Reflection.AssemblyName
        Name = ExecutingApp.GetName()
        oNazwaProgramu = Trim(Name.Name)

        Dim AppFolder As String = Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData)
        oKatRoboczy = AppFolder & "\SoftArt\" & Trim(Name.Name)
        If Not IO.Directory.Exists(oKatRoboczy) Then IO.Directory.CreateDirectory(oKatRoboczy)

        oKatProgramu = Trim(System.Windows.Forms.Application.StartupPath)   'katalog pliku EXE
        oKatParametry = oKatProgramu        'katalog z parametrami
        'If Len(Program_Instancja) > 0 Then
        '    'jesli zdefiniowano instancje - zakladamy podkatalog
        '    oKatParametry = oKatRoboczy & "\" & Trim(Program_Instancja)       'katalog z parametrami
        '    If Not IO.Directory.Exists(oKatParametry) Then IO.Directory.CreateDirectory(oKatParametry)
        'Else
        '    oKatParametry = oKatRoboczy        'katalog z parametrami
        'End If
        oPlikParametry = Trim(Name.Name) & ".CFG"      'plik z parametrami

        oPlikSzerokosciKolumnKartoteka = Trim(Name.Name) & ".KCW"      'plik z parametrami
        oPlikSzerokosciKolumnAnaliza = Trim(Name.Name) & ".ACW"      'plik z parametrami
        oPlikSzerokosciKolumnKontakty = Trim(Name.Name) & ".HCW"      'plik z parametrami
        oPlikSzerokosciKolumnKlienciICoDalej = Trim(Name.Name) & ".CCW"      'plik z parametrami

        'Get IP Address
        oLocalHostName = System.Net.Dns.GetHostName()
        oLocalHostIP = System.Net.Dns.GetHostEntry(oLocalHostName).AddressList(0).ToString()
    End Sub





























    '**************************************
    '** procedury obs³ugi parametrów lokalnych
    '**************************************

    Public Sub InicjujParametryLokalne()
        ParL_OpisInstalacji = ""    'opis biezacej instalacji
        ParL_NazwaArchiwum = "KartotekaKlientowPryzmat"    'opis biezacej instalacji
        ParL_KatalogArchiwum = "c:\"    'opis biezacej instalacji
        ParL_KatalogArchiwumZIP = "c:\"    'opis biezacej instalacji
        ParL_KatalogRaporty = "c:\"    'opis biezacej instalacji
        ParL_DataBaseType = "MS ACCESS"
        '-------------------------------------
        ParL_eMailService = "MS Outlook"
        ParL_eMailAdres = ""
        ParL_SMTP = ""
        ParL_SMTPPort = "587"          '"25"
        ParL_POP3 = ""
        ParL_POP3Port = "110"
        ParL_eMailUser = ""
        ParL_eMailPass = ""
        ParL_SSLProtocol = "NIE"
        '-------------------------------------
        ParL_KatZbiorow = Trim(System.Windows.Forms.Application.StartupPath)        'katalog z plikami danych
        '-------------------------------------
        ParL_Autentykacja = "Windows"
        ParL_Serwer = ""
        ParL_Login = ""
        ParL_Password = ""
        ParL_DataBase = oBazaDanych
        '-------------------------------------
        ParL_HQ = "NIE"
        ParL_HQDataBaseType = "MS SQL"
        ParL_HQPlikMDB = Trim(System.Windows.Forms.Application.StartupPath) & "\" & oNazwaProgramu & ".mdb"
        ParL_HQAutentykacja = "Windows"
        ParL_HQSerwer = ""
        ParL_HQLogin = ""
        ParL_HQPassword = ""
        ParL_HQDataBase = oBazaDanych
    End Sub

    Public Sub PobierzParametryLokalne()
        Dim FileNum As Integer
        Dim ZnakRownaSie As Integer
        Dim NextLine As String
        Dim Klucz As String
        Dim Parametr As String

        InicjujParametryLokalne()
        If Len(Dir(oKatParametry & "\" & oPlikParametry)) > 0 Then
            FileNum = FreeFile()    'kolejny nr otwartego zbioru
            FileOpen(FileNum, oKatParametry & "\" & oPlikParametry, OpenMode.Input)
            If Not EOF(FileNum) Then
                NextLine = LineInput(FileNum)
                'sprawdz czy to plik z parametrami programu
                If InStr(NextLine, "SOFTART : Parametry lokalne programu Magazyn") = 1 Then
                    Do Until EOF(FileNum)
                        NextLine = LineInput(FileNum)
                        ZnakRownaSie = InStr(NextLine, "=")
                        If ZnakRownaSie > 0 Then
                            Klucz = Mid(NextLine, 1, ZnakRownaSie - 1)
                            Parametr = Mid(NextLine, ZnakRownaSie + 1)
                            Select Case Klucz
                                Case "Opis instalacji"
                                    ParL_OpisInstalacji = Parametr
                                Case "Nazwa archiwum"
                                    ParL_NazwaArchiwum = Parametr
                                Case "Katalog archiwum"
                                    ParL_KatalogArchiwum = Parametr
                                Case "Katalog archiwum ZIP"
                                    ParL_KatalogArchiwumZIP = Parametr
                                Case "Katalog na raporty"
                                    ParL_KatalogRaporty = Parametr
                                Case "Jaka Baza Danych"
                                    ParL_DataBaseType = Parametr
                                    '-------------------------------------
                                Case "eMail serwis"
                                    ParL_eMailService = Parametr
                                Case "Adres eMail"
                                    ParL_eMailAdres = Parametr
                                Case "Serwer SMTP"
                                    ParL_SMTP = Parametr
                                Case "Port SMTP"
                                    ParL_SMTPPort = Parametr
                                Case "Serwer POP3"
                                    ParL_POP3 = Parametr
                                Case "Port POP3"
                                    ParL_POP3Port = Parametr
                                Case "Nazwa u¿ytkownika"
                                    ParL_eMailUser = Parametr
                                Case "Has³o u¿ytkownika"
                                    ParL_eMailPass = Parametr
                                Case "Protokol SSL"
                                    ParL_SSLProtocol = Parametr
                                    '-------------------------------------
                                Case "Katalog dyskowy z Baz¹ danych ACCESS", "Katalog dyskowy z Baz¹ danych"
                                    ParL_KatZbiorow = Parametr
                                    '-------------------------------------
                                Case "Autentykacja"
                                    ParL_Autentykacja = Parametr
                                Case "Serwer SQL"
                                    ParL_Serwer = Parametr
                                Case "Login"
                                    ParL_Login = Parametr
                                Case "Password"
                                    ParL_Password = Parametr
                                Case "DataBase"
                                    ParL_DataBase = Parametr
                                    '-------------------------------------
                                Case "IMPORT Czy wspolpracuje"
                                    ParL_HQ = Parametr
                                Case "IMPORT Jaka Baza Danych"
                                    'ParL_HQDataBaseType = Parametr
                                    ParL_HQDataBaseType = "MS SQL"
                                Case "IMPORT Baza danych ACCESS"
                                    ParL_HQPlikMDB = Parametr
                                Case "IMPORT Autentykacja"
                                    ParL_HQAutentykacja = Parametr
                                Case "IMPORT Serwer SQL"
                                    ParL_HQSerwer = Parametr
                                Case "IMPORT Login"
                                    ParL_HQLogin = Parametr
                                Case "IMPORT Password"
                                    ParL_HQPassword = Parametr
                                Case "IMPORT DataBase"
                                    ParL_HQDataBase = Parametr
                                    '-------------------------------------



                            End Select
                        End If
                    Loop
                End If
                FileClose(FileNum)
            End If
        Else
            InicjujParametryLokalne()
        End If
    End Sub

    Public Sub ZapiszParametryLokalne()
        Dim FileNum As Integer

        FileNum = FreeFile()    'kolejny nr otwartego zbioru
        FileOpen(FileNum, oKatParametry & "\" & oPlikParametry, OpenMode.Output)
        PrintLine(FileNum, "SOFTART : Parametry lokalne programu Magazyn")
        PrintLine(FileNum, "Opis instalacji=" & Trim(ParL_OpisInstalacji))
        PrintLine(FileNum, "Nazwa archiwum=" & Trim(ParL_NazwaArchiwum))
        PrintLine(FileNum, "Katalog archiwum=" & Trim(ParL_KatalogArchiwum))
        PrintLine(FileNum, "Katalog archiwum ZIP=" & Trim(ParL_KatalogArchiwumZIP))
        PrintLine(FileNum, "Katalog na raporty=" & Trim(ParL_KatalogRaporty))
        PrintLine(FileNum, "Jaka Baza Danych=" & Trim(ParL_DataBaseType))
        '-------------------------------------
        PrintLine(FileNum, "eMail serwis=" & Trim(ParL_eMailService))
        PrintLine(FileNum, "Adres eMail=" & Trim(ParL_eMailAdres))
        PrintLine(FileNum, "Serwer SMTP=" & Trim(ParL_SMTP))
        PrintLine(FileNum, "Port SMTP=" & Trim(ParL_SMTPPort))
        PrintLine(FileNum, "Serwer POP3=" & Trim(ParL_POP3))
        PrintLine(FileNum, "Port POP3=" & Trim(ParL_POP3Port))
        PrintLine(FileNum, "Nazwa u¿ytkownika=" & Trim(ParL_eMailUser))
        PrintLine(FileNum, "Has³o u¿ytkownika=" & Trim(ParL_eMailPass))
        PrintLine(FileNum, "Protokol SSL=" & Trim(ParL_SSLProtocol))
        '-------------------------------------
        PrintLine(FileNum, "Katalog dyskowy z Baz¹ danych ACCESS=" & Trim(ParL_KatZbiorow))
        '-------------------------------------
        PrintLine(FileNum, "Autentykacja=" & Trim(ParL_Autentykacja))
        PrintLine(FileNum, "Serwer SQL=" & Trim(ParL_Serwer))
        PrintLine(FileNum, "Login=" & Trim(ParL_Login))
        PrintLine(FileNum, "Password=" & Trim(ParL_Password))
        PrintLine(FileNum, "DataBase=" & Trim(ParL_DataBase))
        '-------------------------------------
        PrintLine(FileNum, "IMPORT Czy wspolpracuje=" & Trim(ParL_HQ))
        PrintLine(FileNum, "IMPORT Jaka Baza Danych=" & Trim(ParL_HQDataBaseType))
        PrintLine(FileNum, "IMPORT Baza danych ACCESS=" & Trim(ParL_HQPlikMDB))
        PrintLine(FileNum, "IMPORT Autentykacja=" & Trim(ParL_HQAutentykacja))
        PrintLine(FileNum, "IMPORT Serwer SQL=" & Trim(ParL_HQSerwer))
        PrintLine(FileNum, "IMPORT Login=" & Trim(ParL_HQLogin))
        PrintLine(FileNum, "IMPORT Password=" & Trim(ParL_HQPassword))
        PrintLine(FileNum, "IMPORT DataBase=" & Trim(ParL_HQDataBase))
        '-------------------------------------
        PrintLine(FileNum, "Koniec pliku parametrów lokalnych")
        FileClose(FileNum)
    End Sub

    '**************************************
    '** procedury narzedziowe
    '**************************************

    Public Function ZmienWersje(ByVal PoprzedniaWersja As Integer) As Integer
        Return IIf(PoprzedniaWersja < 32000, PoprzedniaWersja + 1, 1)
    End Function


    Public Function TestujParametryeMail() As Boolean
        Dim Ret As Boolean = True
        Dim Txt As String = "Brak definicji parametrów poczty eMail:" & vbCrLf

        '    Public ParL_eMailAdres As String = ""
        '    Public ParL_SMTP As String = ""
        '    Public ParL_POP3 As String = ""
        '    Public ParL_eMailUser As String = ""
        '    Public ParL_eMailPass As String = ""

        If Len(ParL_eMailAdres) = 0 Then
            Txt &= "adres eMail " & vbCrLf
            Ret = False
        End If
        If Len(ParL_eMailUser) = 0 Then
            Txt &= "nazwa uzytkownika " & vbCrLf
            Ret = False
        End If
        If Len(ParL_eMailPass) = 0 Then
            Txt &= "has³o uzytkownika " & vbCrLf
            Ret = False
        End If
        If Len(ParL_SMTP) = 0 Then
            Txt &= "nazwa serwera wysy³ania SMTP " & vbCrLf
            Ret = False
        End If
        If Len(ParL_POP3) = 0 Then
            Txt &= "nazwa serwera odbierania poczty POP3 " & vbCrLf
            Ret = False
        End If
        '------------------------
        If Not Ret Then
            MessageBox.Show(Txt &
                            "Proszê zdefiniowaæ parametry i ponownie uruchomiæ funkcje wysy³ania.",
                            "Uwaga :",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation)
        End If
        Return Ret
    End Function














    '**************************************
    '** procedury obs³ugi parametrów
    '**************************************

    Dim sConnectionParametry As String = ConnectionStrings()

    Dim sSelectSQLParametry As String = "SELECT " &
                                            "IDENT, " &
                                            "IDENTUZYTKOWNIKA, " &
                                            "NAZWAUZYTKOWNIKA, " &
                                            "ADRESUZYTKOWNIKA, " &
                                            "BANKUZYTKOWNIKA, " &
                                            "KONTOUZYTKOWNIKA, " &
                                            "MIEJSCOWOSCUZYTKOWNIKA, " &
                                            "NIPUZYTKOWNIKA, " &
                                            "IDENTODDZIALU, " &
                                            "DATAAKTANALIZY, " &
                                            "OSTOKRESANALIZY, " &
                                            "DANEDOANALIZY, " &
                                            "MIESANALIZY1, " &
                                            "MIESANALIZY2, " &
                                            "DANEDOPROGNOZY, " &
                                            "MIESPROGNOZY, " &
                                            "KATALOGOFERTY, " &
                                            "WWWPAYBACK, " &
                                            "RAKOLUMNA01, " &
                                            "RAKOLUMNA02, " &
                                            "RAKOLUMNA03, " &
                                            "RAKOLUMNA04, " &
                                            "RAKOLUMNA05, " &
                                            "RAKOLUMNA06, " &
                                            "RAKOLUMNA07, " &
                                            "RAKOLUMNA08, " &
                                            "RAKOLUMNA09, " &
                                            "RAKOLUMNA10, " &
                                            "WARTGRANWARTOSC, " &
                                            "WARTGRANPROCENT, " &
                                            "WARTGRANMATWARTOSC, " &
                                            "WARTGRANMATPROCENT, " &
                                            "WERSJA " &
                                        "FROM Parametry"


    Dim sUpdateSQLParametry As String = "UPDATE Parametry SET " &
                                            "IdentUzytkownika=@oIdentUzytkownika, " &
                                            "NazwaUzytkownika=@oNazwaUzytkownika, " &
                                            "AdresUzytkownika=@oAdresUzytkownika, " &
                                            "BankUzytkownika=@oBankUzytkownika, " &
                                            "KontoUzytkownika=@oKontoUzytkownika, " &
                                            "MiejscowoscUzytkownika=@oMiejscowoscUzytkownika, " &
                                            "NIPUzytkownika=@oNIPUzytkownika, " &
                                            "IDENTODDZIALU=@oIDENTODDZIALU, " &
                                            "DATAAKTANALIZY=@oDATAAKTANALIZY, " &
                                            "OSTOKRESANALIZY=@oOSTOKRESANALIZY, " &
                                            "DANEDOANALIZY=@oDANEDOANALIZY, " &
                                            "MIESANALIZY1=@oMIESANALIZY1, " &
                                            "MIESANALIZY2=@oMIESANALIZY2, " &
                                            "DANEDOPROGNOZY=@oDANEDOPROGNOZY, " &
                                            "MIESPROGNOZY=@oMIESPROGNOZY, " &
                                            "KATALOGOFERTY=@oKATALOGOFERTY, " &
                                            "WWWPAYBACK=@owwwPAYBACK, " &
                                            "RAKOLUMNA01=@oRAKOLUMNA01, " &
                                            "RAKOLUMNA02=@oRAKOLUMNA02, " &
                                            "RAKOLUMNA03=@oRAKOLUMNA03, " &
                                            "RAKOLUMNA04=@oRAKOLUMNA04, " &
                                            "RAKOLUMNA05=@oRAKOLUMNA05, " &
                                            "RAKOLUMNA06=@oRAKOLUMNA06, " &
                                            "RAKOLUMNA07=@oRAKOLUMNA07, " &
                                            "RAKOLUMNA08=@oRAKOLUMNA08, " &
                                            "RAKOLUMNA09=@oRAKOLUMNA09, " &
                                            "RAKOLUMNA10=@oRAKOLUMNA10, " &
                                            "WARTGRANWARTOSC=@oWARTGRANWARTOSC, " &
                                            "WARTGRANPROCENT=@oWARTGRANPROCENT, " &
                                            "WARTGRANMATWARTOSC=@oWARTGRANMATWARTOSC, " &
                                            "WARTGRANMATPROCENT=@oWARTGRANMATPROCENT, " &
                                            "WERSJA=@oWersja " &
                                        "WHERE (IDENT=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Dim sInsertSQLParametry As String = "INSERT INTO Parametry " &
                                         "(IDENT, " &
                                            "IDENTUZYTKOWNIKA, " &
                                            "NAZWAUZYTKOWNIKA, " &
                                            "ADRESUZYTKOWNIKA, " &
                                            "BANKUZYTKOWNIKA, " &
                                            "KONTOUZYTKOWNIKA, " &
                                            "MIEJSCOWOSCUZYTKOWNIKA, " &
                                            "NIPUZYTKOWNIKA, " &
                                            "IDENTODDZIALU, " &
                                            "DATAAKTANALIZY, " &
                                            "OSTOKRESANALIZY, " &
                                            "DANEDOANALIZY, " &
                                            "MIESANALIZY1, " &
                                            "MIESANALIZY2, " &
                                            "DANEDOPROGNOZY, " &
                                            "MIESPROGNOZY, " &
                                            "KATALOGOFERTY, " &
                                            "WWWPAYBACK, " &
                                            "RAKOLUMNA01, " &
                                            "RAKOLUMNA02, " &
                                            "RAKOLUMNA03, " &
                                            "RAKOLUMNA04, " &
                                            "RAKOLUMNA05, " &
                                            "RAKOLUMNA06, " &
                                            "RAKOLUMNA07, " &
                                            "RAKOLUMNA08, " &
                                            "RAKOLUMNA09, " &
                                            "RAKOLUMNA10, " &
                                            "WARTGRANWARTOSC, " &
                                            "WARTGRANPROCENT, " &
                                            "WERSJA " &
                                         ") " &
                                "VALUES  (@oIDENT, " &
                                            "@oIdentUzytkownika, " &
                                            "@oNazwaUzytkownika, " &
                                            "@oAdresUzytkownika, " &
                                            "@oBankUzytkownika, " &
                                            "@oKontoUzytkownika, " &
                                            "@oMiejscowoscUzytkownika, " &
                                            "@oNIPUzytkownika, " &
                                            "@oIDENTODDZIALU, " &
                                            "@oDATAAKTANALIZY, " &
                                            "@oOSTOKRESANALIZY, " &
                                            "@oDANEDOANALIZY, " &
                                            "@oMIESANALIZY1, " &
                                            "@oMIESANALIZY2, " &
                                            "@oDANEDOPROGNOZY, " &
                                            "@oMIESPROGNOZY, " &
                                            "@oKATALOGOFERTY, " &
                                            "@oWWWPAYBACK, " &
                                            "@oRAKOLUMNA01, " &
                                            "@oRAKOLUMNA02, " &
                                            "@oRAKOLUMNA03, " &
                                            "@oRAKOLUMNA04, " &
                                            "@oRAKOLUMNA05, " &
                                            "@oRAKOLUMNA06, " &
                                            "@oRAKOLUMNA07, " &
                                            "@oRAKOLUMNA08, " &
                                            "@oRAKOLUMNA09, " &
                                            "@oRAKOLUMNA10, " &
                                            "@oWARTGRANWARTOSC, " &
                                            "@oWARTGRANPROCENT, " &
                                            "@oWARTGRANMATWARTOSC, " &
                                            "@oWARTGRANMATPROCENT, " &
                                            "@oWERSJA " &
                                         ")"


    Dim dbConnectionParametry As OleDb.OleDbConnection = Nothing
    Dim dbCommandSelectParametry As OleDb.OleDbCommand = Nothing
    Dim dbCommandUpdateParametry As OleDb.OleDbCommand = Nothing
    Dim dbCommandInsertParametry As OleDb.OleDbCommand = Nothing
    Dim dbDataAdapterParametry As OleDb.OleDbDataAdapter = Nothing

    Dim sqlConnectionParametry As SqlClient.SqlConnection = Nothing
    Dim sqlCommandSelectParametry As SqlClient.SqlCommand = Nothing
    Dim sqlCommandUpdateParametry As SqlClient.SqlCommand = Nothing
    Dim sqlCommandInsertParametry As SqlClient.SqlCommand = Nothing
    Dim sqlDataAdapterParametry As SqlClient.SqlDataAdapter = Nothing

    Dim DataSetParametry As DataSet
    Dim DataViewParametry As DataView
    Dim ConnectionState As System.Data.ConnectionState


    Public Sub InicjujParametry()
        Par_IdentUzytkownika = ""
        Par_NazwaUzytkownika = ""
        Par_AdresUzytkownika = ""
        Par_KontoUzytkownika = ""
        Par_BankUzytkownika = ""
        Par_MiejscowoscUzytkownika = ""
        Par_NIPUzytkownika = ""
        Par_IdentOddzialu = ""
        Par_DataAktAnalizy = ""
        Par_OstOkresAnalizy = ""
        Par_DaneDoAnalizy = ""
        Par_MiesAnalizy1 = "TTTTTTTTTTTT"
        Par_MiesAnalizy2 = "TTTTTTTTTTTT"
        Par_DaneDoPrognozy = ""
        Par_MiesPrognozy = "TTTTTTTTTTTT"
        Par_KatalogOferty = ""
        Par_wwwPAYBACK = "https://www.payback.pl/pb/id/36624/index.html/?utm_source=pryzmatwww&utm_medium=homepage&utm_campaign=ekupony"

        Par_RAKolumna01 = ""
        Par_RAKolumna02 = ""
        Par_RAKolumna03 = ""
        Par_RAKolumna04 = ""
        Par_RAKolumna05 = ""
        Par_RAKolumna06 = ""
        Par_RAKolumna07 = ""
        Par_RAKolumna08 = ""
        Par_RAKolumna09 = ""
        Par_RAKolumna10 = ""

        Par_WartGranWartosc = 0
        Par_WartGranProcent = 80
        Par_WartGranMatWartosc = 0
        Par_WartGranMatProcent = 80
        Par_Wersja = 0
    End Sub

    Public Sub PobierzParametry()
        sConnectionParametry = ConnectionStrings()
        DataSetParametry = New DataSet
        DataViewParametry = New DataView
        DataSetParametry.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionParametry = New OleDb.OleDbConnection(sConnectionParametry)
            dbCommandSelectParametry = New OleDb.OleDbCommand(sSelectSQLParametry, dbConnectionParametry)
            dbDataAdapterParametry = New OleDb.OleDbDataAdapter(dbCommandSelectParametry)

            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = dbDataAdapterParametry.TableMappings.Add("Table", "TABELA_Parametry")
            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionParametry.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionState = dbConnectionParametry.State
            End Try
        Else
            sqlConnectionParametry = New SqlClient.SqlConnection(sConnectionParametry)
            sqlCommandSelectParametry = New SqlClient.SqlCommand(sSelectSQLParametry, sqlConnectionParametry)
            sqlDataAdapterParametry = New SqlClient.SqlDataAdapter(sqlCommandSelectParametry)

            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterParametry.TableMappings.Add("Table", "TABELA_Parametry")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionParametry.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionState = sqlConnectionParametry.State
            End Try
        End If

        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapterParametry.FillSchema(DataSetParametry, SchemaType.Mapped, "TABELA_Parametry")
                    dbDataAdapterParametry.Fill(DataSetParametry)
                    dbConnectionParametry.Close()
                Else
                    sqlDataAdapterParametry.FillSchema(DataSetParametry, SchemaType.Mapped, "TABELA_Parametry")
                    sqlDataAdapterParametry.Fill(DataSetParametry)
                    sqlConnectionParametry.Close()
                End If

                'definiuj DataView
                DataViewParametry = New DataView(DataSetParametry.Tables(0))
                If DataViewParametry.Count > 0 Then
                    Par_IdentUzytkownika = Trim(GetTxtDRVField(DataViewParametry.Item(0), "IdentUzytkownika"))
                    Par_NazwaUzytkownika = GetTxtDRVField(DataViewParametry.Item(0), "NazwaUzytkownika")
                    Par_AdresUzytkownika = GetTxtDRVField(DataViewParametry.Item(0), "AdresUzytkownika")
                    Par_BankUzytkownika = GetTxtDRVField(DataViewParametry.Item(0), "BankUzytkownika")
                    Par_KontoUzytkownika = GetTxtDRVField(DataViewParametry.Item(0), "KontoUzytkownika")
                    Par_MiejscowoscUzytkownika = GetTxtDRVField(DataViewParametry.Item(0), "MiejscowoscUzytkownika")
                    Par_NIPUzytkownika = GetTxtDRVField(DataViewParametry.Item(0), "NIPUzytkownika")
                    Par_IdentOddzialu = GetTxtDRVField(DataViewParametry.Item(0), "IdentOddzialu")
                    Par_DataAktAnalizy = GetTxtDRVField(DataViewParametry.Item(0), "DATAAKTANALIZY")
                    Par_OstOkresAnalizy = GetTxtDRVField(DataViewParametry.Item(0), "OSTOKRESANALIZY")
                    Par_DaneDoAnalizy = GetTxtDRVField(DataViewParametry.Item(0), "DANEDOANALIZY")
                    Par_MiesAnalizy1 = GetTxtDRVField(DataViewParametry.Item(0), "MIESANALIZY1")
                    Par_MiesAnalizy2 = GetTxtDRVField(DataViewParametry.Item(0), "MIESANALIZY2")
                    Par_DaneDoPrognozy = GetTxtDRVField(DataViewParametry.Item(0), "DANEDOPROGNOZY")
                    Par_MiesPrognozy = GetTxtDRVField(DataViewParametry.Item(0), "MIESPROGNOZY")
                    Par_KatalogOferty = GetTxtDRVField(DataViewParametry.Item(0), "KATALOGOFERTY")
                    Par_wwwPAYBACK = GetTxtDRVField(DataViewParametry.Item(0), "WWWPAYBACK")

                    Par_RAKolumna01 = GetTxtDRVField(DataViewParametry.Item(0), "RAKOLUMNA01")
                    Par_RAKolumna02 = GetTxtDRVField(DataViewParametry.Item(0), "RAKOLUMNA02")
                    Par_RAKolumna03 = GetTxtDRVField(DataViewParametry.Item(0), "RAKOLUMNA03")
                    Par_RAKolumna04 = GetTxtDRVField(DataViewParametry.Item(0), "RAKOLUMNA04")
                    Par_RAKolumna05 = GetTxtDRVField(DataViewParametry.Item(0), "RAKOLUMNA05")
                    Par_RAKolumna06 = GetTxtDRVField(DataViewParametry.Item(0), "RAKOLUMNA06")
                    Par_RAKolumna07 = GetTxtDRVField(DataViewParametry.Item(0), "RAKOLUMNA07")
                    Par_RAKolumna08 = GetTxtDRVField(DataViewParametry.Item(0), "RAKOLUMNA08")
                    Par_RAKolumna09 = GetTxtDRVField(DataViewParametry.Item(0), "RAKOLUMNA09")
                    Par_RAKolumna10 = GetTxtDRVField(DataViewParametry.Item(0), "RAKOLUMNA10")

                    Par_WartGranWartosc = GetTxtDRVField(DataViewParametry.Item(0), "WARTGRANWARTOSC")
                    Par_WartGranProcent = GetTxtDRVField(DataViewParametry.Item(0), "WARTGRANPROCENT")
                    Par_WartGranMatWartosc = GetTxtDRVField(DataViewParametry.Item(0), "WARTGRANMATWARTOSC")
                    Par_WartGranMatProcent = GetTxtDRVField(DataViewParametry.Item(0), "WARTGRANMATPROCENT")

                    Par_Wersja = DataViewParametry.Item(0).Item("WERSJA")
                Else
                    InicjujParametry()
                End If
            Catch Ex As System.Exception
                'ViewInLog(ex, MyForm.Name(), Nothing)
            End Try

        End If
    End Sub

    Public Sub ZapiszParametry(ByVal MyForm As Form)
        sConnectionParametry = ConnectionStrings()
        DataSetParametry = New DataSet
        DataSetParametry.Locale = New System.Globalization.CultureInfo("pl-PL")

        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionParametry = New OleDb.OleDbConnection(sConnectionParametry)
            dbCommandSelectParametry = New OleDb.OleDbCommand(sSelectSQLParametry, dbConnectionParametry)
            dbCommandUpdateParametry = New OleDb.OleDbCommand(sUpdateSQLParametry, dbConnectionParametry)
            dbCommandInsertParametry = New OleDb.OleDbCommand(sInsertSQLParametry, dbConnectionParametry)
            dbDataAdapterParametry = New OleDb.OleDbDataAdapter(dbCommandSelectParametry)
            DataSetParametry = New DataSet
            DataViewParametry = New DataView

            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = dbDataAdapterParametry.TableMappings.Add("Table", "TABELA_Parametry")
            Dim objParam As OleDb.OleDbParameter
            '------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            'objParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 10, "IDENT")
            'objParam.SourceVersion = DataRowVersion.Original
            'objParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            'objParam.SourceVersion = DataRowVersion.Original
            'dbDataAdapter.DeleteCommand = dbCommandDelete

            '------------------------------------------
            'komenda UPDATE
            'parametry aktualizacji
            'dbCommandUpdate.Parameters.Add("@oIdent", OleDb.OleDbType.Char, 10, "IDENT")
            dbCommandUpdateParametry.Parameters.Add("@oIdentUzytkownika", OleDb.OleDbType.VarChar, 50, "IDENTUZYTKOWNIKA")
            dbCommandUpdateParametry.Parameters.Add("@oNazwaUzytkownika", OleDb.OleDbType.Char, 50, "NAZWAUZYTKOWNIKA")
            dbCommandUpdateParametry.Parameters.Add("@oAdresUzytkownika", OleDb.OleDbType.Char, 50, "ADRESUZYTKOWNIKA")
            dbCommandUpdateParametry.Parameters.Add("@oBankUzytkownika", OleDb.OleDbType.Char, 50, "BANKUZYTKOWNIKA")
            dbCommandUpdateParametry.Parameters.Add("@oKontoUzytkownika", OleDb.OleDbType.Char, 50, "KONTOUZYTKOWNIKA")
            dbCommandUpdateParametry.Parameters.Add("@oMiejscowoscUzytkownika", OleDb.OleDbType.Char, 50, "MIEJSCOWOSCUZYTKOWNIKA")
            dbCommandUpdateParametry.Parameters.Add("@oNIPUzytkownika", OleDb.OleDbType.Char, 15, "NIPUZYTKOWNIKA")
            dbCommandUpdateParametry.Parameters.Add("@oIdentOddzialu", OleDb.OleDbType.Char, 50, "IDENTODDZIALU")
            dbCommandUpdateParametry.Parameters.Add("@oDATAAKTANALIZY", OleDb.OleDbType.Char, 10, "DATAAKTANALIZY")
            dbCommandUpdateParametry.Parameters.Add("@oOSTOKRESANALIZY", OleDb.OleDbType.Char, 14, "OSTOKRESANALIZY")
            dbCommandUpdateParametry.Parameters.Add("@oDANEDOANALIZY", OleDb.OleDbType.Char, 20, "DANEDOANALIZY")
            dbCommandUpdateParametry.Parameters.Add("@oMIESANALIZY1", OleDb.OleDbType.Char, 12, "MIESANALIZY1")
            dbCommandUpdateParametry.Parameters.Add("@oMIESANALIZY2", OleDb.OleDbType.Char, 12, "MIESANALIZY2")
            dbCommandUpdateParametry.Parameters.Add("@oDANEDOPROGNOZY", OleDb.OleDbType.Char, 10, "DANEDOPROGNOZY")
            dbCommandUpdateParametry.Parameters.Add("@oMIESPROGNOZY", OleDb.OleDbType.Char, 12, "MIESPROGNOZY")
            dbCommandUpdateParametry.Parameters.Add("@oKATALOGOFERTY", OleDb.OleDbType.VarChar, 100, "KATALOGOFERTY")
            dbCommandUpdateParametry.Parameters.Add("@oWWWPAYBACK", OleDb.OleDbType.VarChar, 200, "WWWPAYBACK")

            dbCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA01", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA01")
            dbCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA02", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA02")
            dbCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA03", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA03")
            dbCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA04", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA04")
            dbCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA05", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA05")
            dbCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA06", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA06")
            dbCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA07", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA07")
            dbCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA06", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA08")
            dbCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA09", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA09")
            dbCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA10", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA10")

            dbCommandUpdateParametry.Parameters.Add("@oWARTGRANWARTOSC", OleDb.OleDbType.Double, Nothing, "WARTGRANWARTOSC")
            dbCommandUpdateParametry.Parameters.Add("@oWARTGRANPROCENT", OleDb.OleDbType.Double, Nothing, "WARTGRANPROCENT")
            dbCommandUpdateParametry.Parameters.Add("@oWARTGRANMATWARTOSC", OleDb.OleDbType.Double, Nothing, "WARTGRANMATWARTOSC")
            dbCommandUpdateParametry.Parameters.Add("@oWARTGRANMATPROCENT", OleDb.OleDbType.Double, Nothing, "WARTGRANMATPROCENT")

            dbCommandUpdateParametry.Parameters.Add("@oWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")

            'parametry filtrowania
            objParam = dbCommandUpdateParametry.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 10, "IDENT")
            objParam.SourceVersion = DataRowVersion.Original
            objParam = dbCommandUpdateParametry.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Original

            dbDataAdapterParametry.UpdateCommand = dbCommandUpdateParametry

            '------------------------------------------
            'komenda INSERT
            objParam = dbCommandInsertParametry.Parameters.Add("@oIdent", OleDb.OleDbType.Char, 10, "IDENT")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oIdentUzytkownika", OleDb.OleDbType.VarChar, 50, "IDENTUZYTKOWNIKA")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oNazwaUzytkownika", OleDb.OleDbType.Char, 50, "NAZWAUZYTKOWNIKA")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oAdresUzytkownika", OleDb.OleDbType.Char, 50, "ADRESUZYTKOWNIKA")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oBankUzytkownika", OleDb.OleDbType.Char, 50, "BANKUZYTKOWNIKA")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oKontoUzytkownika", OleDb.OleDbType.Char, 50, "KONTOUZYTKOWNIKA")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oMiejscowoscUzytkownika", OleDb.OleDbType.Char, 50, "MIEJSCOWOSCUZYTKOWNIKA")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oNIPUzytkownika", OleDb.OleDbType.Char, 15, "NIPUZYTKOWNIKA")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oIdentOddzialu", OleDb.OleDbType.Char, 50, "IDENTODDZIALU")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oDATAAKTANALIZY", OleDb.OleDbType.Char, 10, "DATAAKTANALIZY")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oOSTOKRESANALIZY", OleDb.OleDbType.Char, 14, "OSTOKRESANALIZY")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oDANEDOANALIZY", OleDb.OleDbType.Char, 20, "DANEDOANALIZY")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oMIESANALIZY1", OleDb.OleDbType.Char, 12, "MIESANALIZY1")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oMIESANALIZY2", OleDb.OleDbType.Char, 12, "MIESANALIZY2")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oDANEDOPROGNOZY", OleDb.OleDbType.Char, 10, "DANEDOPROGNOZY")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oMIESPROGNOZY", OleDb.OleDbType.Char, 12, "MIESPROGNOZY")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oKATALOGOFERTY", OleDb.OleDbType.VarChar, 100, "KATALOGOFERTY")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oWWWPAYBACK", OleDb.OleDbType.VarChar, 200, "WWWPAYBACK")
            objParam.SourceVersion = DataRowVersion.Current

            objParam = dbCommandInsertParametry.Parameters.Add("@oRAKOLUMNA01", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA01")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oRAKOLUMNA02", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA02")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oRAKOLUMNA03", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA03")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oRAKOLUMNA04", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA04")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oRAKOLUMNA05", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA05")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oRAKOLUMNA06", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA06")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oRAKOLUMNA07", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA07")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oRAKOLUMNA08", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA08")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oRAKOLUMNA09", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA09")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oRAKOLUMNA10", OleDb.OleDbType.VarChar, 100, "RAKOLUMNA10")
            objParam.SourceVersion = DataRowVersion.Current

            objParam = dbCommandInsertParametry.Parameters.Add("@oWARTGRANWARTOSC", OleDb.OleDbType.Double, Nothing, "WARTGRANWARTOSC")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oWARTGRANPROCENT", OleDb.OleDbType.Double, Nothing, "WARTGRANPROCENT")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oWARTGRANMATWARTOSC", OleDb.OleDbType.Double, Nothing, "WARTGRANMATWARTOSC")
            objParam.SourceVersion = DataRowVersion.Current
            objParam = dbCommandInsertParametry.Parameters.Add("@oWARTGRANMATPROCENT", OleDb.OleDbType.Double, Nothing, "WARTGRANMATPROCENT")
            objParam.SourceVersion = DataRowVersion.Current


            objParam = dbCommandInsertParametry.Parameters.Add("@oWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
            objParam.SourceVersion = DataRowVersion.Current

            dbDataAdapterParametry.InsertCommand = dbCommandInsertParametry

            ' Add the RowUpdated event handler.
            AddHandler dbDataAdapterParametry.RowUpdated, New OleDb.OleDbRowUpdatedEventHandler(AddressOf OnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionParametry.Open()
            Catch Ex As System.Exception
                'ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)\
            Finally
                ConnectionState = dbConnectionParametry.State
            End Try
        Else
            sqlConnectionParametry = New SqlClient.SqlConnection(sConnectionParametry)
            sqlCommandSelectParametry = New SqlClient.SqlCommand(sSelectSQLParametry, sqlConnectionParametry)
            sqlCommandUpdateParametry = New SqlClient.SqlCommand(sUpdateSQLParametry, sqlConnectionParametry)
            sqlCommandInsertParametry = New SqlClient.SqlCommand(sInsertSQLParametry, sqlConnectionParametry)
            sqlDataAdapterParametry = New SqlClient.SqlDataAdapter(sqlCommandSelectParametry)
            DataSetParametry = New DataSet
            DataViewParametry = New DataView

            Dim MapowanieTabeli As System.Data.Common.DataTableMapping
            MapowanieTabeli = sqlDataAdapterParametry.TableMappings.Add("Table", "TABELA_Parametry")
            Dim sqlObjParam As SqlClient.SqlParameter
            '------------------------------------------'------------------------------------------
            'komenda DELETE
            'parametry filtrowania
            'sqlObjParam = sqlcommandDelete.Parameters.Add("@orygSymbol", sqldbtype.Char, 10, "IDENT")
            'sqlObjParam.SourceVersion = DataRowVersion.Original
            'sqlObjParam = sqlcommandDelete.Parameters.Add("@orygWersja", sqldbtype.Integer, Nothing, "WERSJA")
            'sqlObjParam.SourceVersion = DataRowVersion.Original
            'sqldataadapter.DeleteCommand = sqlcommandDelete

            '------------------------------------------
            'komenda UPDATE
            'parametry aktualizacji
            'sqlcommandUpdate.Parameters.Add("@oIdent", sqldbtype.Char, 10, "IDENT")
            sqlCommandUpdateParametry.Parameters.Add("@oIdentUzytkownika", SqlDbType.VarChar, 50, "IDENTUZYTKOWNIKA")
            sqlCommandUpdateParametry.Parameters.Add("@oNazwaUzytkownika", SqlDbType.Char, 50, "NAZWAUZYTKOWNIKA")
            sqlCommandUpdateParametry.Parameters.Add("@oAdresUzytkownika", SqlDbType.Char, 50, "ADRESUZYTKOWNIKA")
            sqlCommandUpdateParametry.Parameters.Add("@oBankUzytkownika", SqlDbType.Char, 50, "BANKUZYTKOWNIKA")
            sqlCommandUpdateParametry.Parameters.Add("@oKontoUzytkownika", SqlDbType.Char, 50, "KONTOUZYTKOWNIKA")
            sqlCommandUpdateParametry.Parameters.Add("@oMiejscowoscUzytkownika", SqlDbType.Char, 50, "MIEJSCOWOSCUZYTKOWNIKA")
            sqlCommandUpdateParametry.Parameters.Add("@oNIPUzytkownika", SqlDbType.Char, 15, "NIPUZYTKOWNIKA")
            sqlCommandUpdateParametry.Parameters.Add("@oIdentOddzialu", SqlDbType.Char, 50, "IDENTODDZIALU")
            sqlCommandUpdateParametry.Parameters.Add("@oDATAAKTANALIZY", SqlDbType.Char, 10, "DATAAKTANALIZY")
            sqlCommandUpdateParametry.Parameters.Add("@oOSTOKRESANALIZY", SqlDbType.Char, 14, "OSTOKRESANALIZY")
            sqlCommandUpdateParametry.Parameters.Add("@oDANEDOANALIZY", SqlDbType.Char, 20, "DANEDOANALIZY")
            sqlCommandUpdateParametry.Parameters.Add("@oMIESANALIZY1", SqlDbType.Char, 12, "MIESANALIZY1")
            sqlCommandUpdateParametry.Parameters.Add("@oMIESANALIZY2", SqlDbType.Char, 12, "MIESANALIZY2")
            sqlCommandUpdateParametry.Parameters.Add("@oDANEDOPROGNOZY", SqlDbType.Char, 10, "DANEDOPROGNOZY")
            sqlCommandUpdateParametry.Parameters.Add("@oMIESPROGNOZY", SqlDbType.Char, 12, "MIESPROGNOZY")
            sqlCommandUpdateParametry.Parameters.Add("@oKATALOGOFERTY", SqlDbType.VarChar, 100, "KATALOGOFERTY")
            sqlCommandUpdateParametry.Parameters.Add("@oWWWPAYBACK", SqlDbType.VarChar, 200, "WWWPAYBACK")

            sqlCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA01", SqlDbType.VarChar, 100, "RAKOLUMNA01")
            sqlCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA02", SqlDbType.VarChar, 100, "RAKOLUMNA02")
            sqlCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA03", SqlDbType.VarChar, 100, "RAKOLUMNA03")
            sqlCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA04", SqlDbType.VarChar, 100, "RAKOLUMNA04")
            sqlCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA05", SqlDbType.VarChar, 100, "RAKOLUMNA05")
            sqlCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA06", SqlDbType.VarChar, 100, "RAKOLUMNA06")
            sqlCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA07", SqlDbType.VarChar, 100, "RAKOLUMNA07")
            sqlCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA08", SqlDbType.VarChar, 100, "RAKOLUMNA08")
            sqlCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA09", SqlDbType.VarChar, 100, "RAKOLUMNA09")
            sqlCommandUpdateParametry.Parameters.Add("@oRAKOLUMNA10", SqlDbType.VarChar, 100, "RAKOLUMNA10")

            sqlCommandUpdateParametry.Parameters.Add("@oWARTGRANWARTOSC", SqlDbType.Float, Nothing, "WARTGRANWARTOSC")
            sqlCommandUpdateParametry.Parameters.Add("@oWARTGRANPROCENT", SqlDbType.Float, Nothing, "WARTGRANPROCENT")
            sqlCommandUpdateParametry.Parameters.Add("@oWARTGRANMATWARTOSC", SqlDbType.Float, Nothing, "WARTGRANMATWARTOSC")
            sqlCommandUpdateParametry.Parameters.Add("@oWARTGRANMATPROCENT", SqlDbType.Float, Nothing, "WARTGRANMATPROCENT")

            sqlCommandUpdateParametry.Parameters.Add("@oWersja", SqlDbType.Int, Nothing, "WERSJA")

            'parametry filtrowania
            sqlObjParam = sqlCommandUpdateParametry.Parameters.Add("@orygSymbol", SqlDbType.Char, 10, "IDENT")
            sqlObjParam.SourceVersion = DataRowVersion.Original
            sqlObjParam = sqlCommandUpdateParametry.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
            sqlObjParam.SourceVersion = DataRowVersion.Original

            sqlDataAdapterParametry.UpdateCommand = sqlCommandUpdateParametry

            '------------------------------------------
            'komenda INSERT
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oIdent", SqlDbType.Char, 10, "IDENT")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oIdentUzytkownika", SqlDbType.VarChar, 50, "IDENTUZYTKOWNIKA")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oNazwaUzytkownika", SqlDbType.Char, 50, "NAZWAUZYTKOWNIKA")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oAdresUzytkownika", SqlDbType.Char, 50, "ADRESUZYTKOWNIKA")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oBankUzytkownika", SqlDbType.Char, 50, "BANKUZYTKOWNIKA")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oKontoUzytkownika", SqlDbType.Char, 50, "KONTOUZYTKOWNIKA")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oMiejscowoscUzytkownika", SqlDbType.Char, 50, "MIEJSCOWOSCUZYTKOWNIKA")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oNIPUzytkownika", SqlDbType.Char, 15, "NIPUZYTKOWNIKA")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oIdentOddzialu", SqlDbType.Char, 50, "IDENTODDZIALU")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oDATAAKTANALIZY", SqlDbType.Char, 10, "DATAAKTANALIZY")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oOSTOKRESANALIZY", SqlDbType.Char, 14, "OSTOKRESANALIZY")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oDANEDOANALIZY", SqlDbType.Char, 20, "DANEDOANALIZY")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oMIESANALIZY1", SqlDbType.Char, 12, "MIESANALIZY1")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oMIESANALIZY2", SqlDbType.Char, 12, "MIESANALIZY2")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oDANEDOPROGNOZY", SqlDbType.Char, 10, "DANEDOPROGNOZY")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oMIESPROGNOZY", SqlDbType.Char, 12, "MIESPROGNOZY")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oKATALOGOFERTY", SqlDbType.VarChar, 100, "KATALOGOFERTY")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oWWWPAYBACK", SqlDbType.VarChar, 200, "WWWPAYBACK")
            sqlObjParam.SourceVersion = DataRowVersion.Current

            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oRAKOLUMNA01", SqlDbType.VarChar, 100, "RAKOLUMNA01")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oRAKOLUMNA02", SqlDbType.VarChar, 100, "RAKOLUMNA02")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oRAKOLUMNA03", SqlDbType.VarChar, 100, "RAKOLUMNA03")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oRAKOLUMNA04", SqlDbType.VarChar, 100, "RAKOLUMNA04")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oRAKOLUMNA05", SqlDbType.VarChar, 100, "RAKOLUMNA05")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oRAKOLUMNA06", SqlDbType.VarChar, 100, "RAKOLUMNA06")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oRAKOLUMNA07", SqlDbType.VarChar, 100, "RAKOLUMNA07")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oRAKOLUMNA08", SqlDbType.VarChar, 100, "RAKOLUMNA08")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oRAKOLUMNA09", SqlDbType.VarChar, 100, "RAKOLUMNA09")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oRAKOLUMNA10", SqlDbType.VarChar, 100, "RAKOLUMNA10")
            sqlObjParam.SourceVersion = DataRowVersion.Current

            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oWARTGRANWARTOSC", SqlDbType.Float, Nothing, "WARTGRANWARTOSC")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oWARTGRANPROCENT", SqlDbType.Float, Nothing, "WARTGRANPROCENT")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oWARTGRANMATWARTOSC", SqlDbType.Float, Nothing, "WARTGRANMATWARTOSC")
            sqlObjParam.SourceVersion = DataRowVersion.Current
            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oWARTGRANMATPROCENT", SqlDbType.Float, Nothing, "WARTGRANMATPROCENT")
            sqlObjParam.SourceVersion = DataRowVersion.Current

            sqlObjParam = sqlCommandInsertParametry.Parameters.Add("@oWersja", SqlDbType.Int, Nothing, "WERSJA")
            sqlObjParam.SourceVersion = DataRowVersion.Current

            sqlDataAdapterParametry.InsertCommand = sqlCommandInsertParametry

            ' Add the RowUpdated event handler.
            AddHandler sqlDataAdapterParametry.RowUpdated, New SqlClient.SqlRowUpdatedEventHandler(AddressOf sqlOnRowUpdated)

            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionParametry.Open()
            Catch Ex As System.Exception
                'ViewInLog(Ex, Me.Name(), Nothing)
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)\
            Finally
                ConnectionState = sqlConnectionParametry.State
            End Try
        End If


        If ConnectionState = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    dbDataAdapterParametry.FillSchema(DataSetParametry, SchemaType.Mapped, "TABELA_Parametry")
                    dbDataAdapterParametry.Fill(DataSetParametry)
                    dbConnectionParametry.Close()
                Else
                    sqlDataAdapterParametry.FillSchema(DataSetParametry, SchemaType.Mapped, "TABELA_Parametry")
                    sqlDataAdapterParametry.Fill(DataSetParametry)
                    sqlConnectionParametry.Close()
                End If

                Dim foundRow As DataRow
                'definiuj DataView
                DataViewParametry = New DataView(DataSetParametry.Tables(0))
                If DataViewParametry.Count > 0 Then
                    ' Create an array for the key values to find.
                    Dim findTheseVals(0) As Object
                    findTheseVals(0) = "SOFTART"
                    foundRow = DataSetParametry.Tables(0).Rows.Find(findTheseVals)
                    ' sprawdz czy jest...
                    If Not (foundRow Is Nothing) Then
                        'foundRow("IDENT") = "SOFTART"
                        foundRow("IdentUzytkownika") = Par_IdentUzytkownika
                        foundRow("NazwaUzytkownika") = Par_NazwaUzytkownika
                        foundRow("AdresUzytkownika") = Par_AdresUzytkownika
                        foundRow("BankUzytkownika") = Par_BankUzytkownika
                        foundRow("KontoUzytkownika") = Par_KontoUzytkownika
                        foundRow("MiejscowoscUzytkownika") = Par_MiejscowoscUzytkownika
                        foundRow("NIPUzytkownika") = Par_NIPUzytkownika
                        foundRow("IDENTOddzialu") = Par_IdentOddzialu

                        foundRow("DATAAKTANALIZY") = Par_DataAktAnalizy
                        foundRow("OSTOKRESANALIZY") = Par_OstOkresAnalizy
                        foundRow("DANEDOANALIZY") = Par_DaneDoAnalizy
                        foundRow("MIESANALIZY1") = Par_MiesAnalizy1
                        foundRow("MIESANALIZY2") = Par_MiesAnalizy2
                        foundRow("DANEDOPROGNOZY") = Par_DaneDoPrognozy
                        foundRow("MIESPROGNOZY") = Par_MiesPrognozy
                        foundRow("KATALOGOFERTY") = Par_KatalogOferty
                        foundRow("WWWPAYBACK") = Par_wwwPAYBACK

                        foundRow("RAKOLUMNA01") = Par_RAKolumna01
                        foundRow("RAKOLUMNA02") = Par_RAKolumna02
                        foundRow("RAKOLUMNA03") = Par_RAKolumna03
                        foundRow("RAKOLUMNA04") = Par_RAKolumna04
                        foundRow("RAKOLUMNA05") = Par_RAKolumna05
                        foundRow("RAKOLUMNA06") = Par_RAKolumna06
                        foundRow("RAKOLUMNA07") = Par_RAKolumna07
                        foundRow("RAKOLUMNA08") = Par_RAKolumna08
                        foundRow("RAKOLUMNA09") = Par_RAKolumna09
                        foundRow("RAKOLUMNA10") = Par_RAKolumna10

                        foundRow("WARTGRANWARTOSC") = Par_WartGranWartosc
                        foundRow("WARTGRANPROCENT") = Par_WartGranProcent
                        foundRow("WARTGRANMATWARTOSC") = Par_WartGranMatWartosc
                        foundRow("WARTGRANMATPROCENT") = Par_WartGranMatProcent

                        foundRow("WERSJA") = ZmienWersje(Par_Wersja)    'Wersja         Integer, 2
                    Else
                        'stworz nowy rekord
                        foundRow = DataSetParametry.Tables(0).NewRow()

                        foundRow("IDENT") = "SOFTART"
                        foundRow("IdentUzytkownika") = Par_IdentUzytkownika
                        foundRow("NazwaUzytkownika") = Par_NazwaUzytkownika
                        foundRow("AdresUzytkownika") = Par_AdresUzytkownika
                        foundRow("BankUzytkownika") = Par_BankUzytkownika
                        foundRow("KontoUzytkownika") = Par_KontoUzytkownika
                        foundRow("MiejscowoscUzytkownika") = Par_MiejscowoscUzytkownika
                        foundRow("NIPUzytkownika") = Par_NIPUzytkownika
                        foundRow("IDENTOddzialu") = Par_IdentOddzialu

                        foundRow("DATAAKTANALIZY") = Par_DataAktAnalizy
                        foundRow("OSTOKRESANALIZY") = Par_OstOkresAnalizy
                        foundRow("DANEDOANALIZY") = Par_DaneDoAnalizy
                        foundRow("MIESANALIZY1") = Par_MiesAnalizy1
                        foundRow("MIESANALIZY2") = Par_MiesAnalizy2
                        foundRow("DANEDOPROGNOZY") = Par_DaneDoPrognozy
                        foundRow("MIESPROGNOZY") = Par_MiesPrognozy
                        foundRow("KATALOGOFERTY") = Par_KatalogOferty
                        foundRow("WWWPAYBACK") = Par_wwwPAYBACK

                        foundRow("RAKOLUMNA01") = Par_RAKolumna01
                        foundRow("RAKOLUMNA02") = Par_RAKolumna02
                        foundRow("RAKOLUMNA03") = Par_RAKolumna03
                        foundRow("RAKOLUMNA04") = Par_RAKolumna04
                        foundRow("RAKOLUMNA05") = Par_RAKolumna05
                        foundRow("RAKOLUMNA06") = Par_RAKolumna06
                        foundRow("RAKOLUMNA07") = Par_RAKolumna07
                        foundRow("RAKOLUMNA08") = Par_RAKolumna08
                        foundRow("RAKOLUMNA09") = Par_RAKolumna09
                        foundRow("RAKOLUMNA10") = Par_RAKolumna10

                        foundRow("WARTGRANWARTOSC") = Par_WartGranWartosc
                        foundRow("WARTGRANPROCENT") = Par_WartGranProcent
                        foundRow("WARTGRANMATWARTOSC") = Par_WartGranMatWartosc
                        foundRow("WARTGRANMATPROCENT") = Par_WartGranMatProcent

                        foundRow("WERSJA") = 1    'Wersja         Integer, 2

                        'dodaj rekord do DataSet
                        DataSetParametry.Tables(0).Rows.Add(foundRow)
                    End If
                Else
                    'dopisuj - stworz nowy rekord
                    foundRow = DataSetParametry.Tables(0).NewRow()

                    foundRow("IDENT") = "SOFTART"
                    foundRow("IdentUzytkownika") = Par_IdentUzytkownika
                    foundRow("NazwaUzytkownika") = Par_NazwaUzytkownika
                    foundRow("AdresUzytkownika") = Par_AdresUzytkownika
                    foundRow("BankUzytkownika") = Par_BankUzytkownika
                    foundRow("KontoUzytkownika") = Par_KontoUzytkownika
                    foundRow("MiejscowoscUzytkownika") = Par_MiejscowoscUzytkownika
                    foundRow("NIPUzytkownika") = Par_NIPUzytkownika
                    foundRow("IDENTOddzialu") = Par_IdentOddzialu

                    foundRow("DATAAKTANALIZY") = Par_DataAktAnalizy
                    foundRow("OSTOKRESANALIZY") = Par_OstOkresAnalizy
                    foundRow("DANEDOANALIZY") = Par_DaneDoAnalizy
                    foundRow("MIESANALIZY1") = Par_MiesAnalizy1
                    foundRow("MIESANALIZY2") = Par_MiesAnalizy2
                    foundRow("DANEDOPROGNOZY") = Par_DaneDoPrognozy
                    foundRow("MIESPROGNOZY") = Par_MiesPrognozy
                    foundRow("KATALOGOFERTY") = Par_KatalogOferty
                    foundRow("WWWPAYBACK") = Par_wwwPAYBACK

                    foundRow("RAKOLUMNA01") = Par_RAKolumna01
                    foundRow("RAKOLUMNA02") = Par_RAKolumna02
                    foundRow("RAKOLUMNA03") = Par_RAKolumna03
                    foundRow("RAKOLUMNA04") = Par_RAKolumna04
                    foundRow("RAKOLUMNA05") = Par_RAKolumna05
                    foundRow("RAKOLUMNA06") = Par_RAKolumna06
                    foundRow("RAKOLUMNA07") = Par_RAKolumna07
                    foundRow("RAKOLUMNA08") = Par_RAKolumna08
                    foundRow("RAKOLUMNA09") = Par_RAKolumna09
                    foundRow("RAKOLUMNA10") = Par_RAKolumna10

                    foundRow("WARTGRANWARTOSC") = Par_WartGranWartosc
                    foundRow("WARTGRANPROCENT") = Par_WartGranProcent
                    foundRow("WARTGRANMATWARTOSC") = Par_WartGranMatWartosc
                    foundRow("WARTGRANMATPROCENT") = Par_WartGranMatProcent

                    foundRow("WERSJA") = 1    'Wersja         Integer, 2

                    'dodaj rekord do DataSet
                    DataSetParametry.Tables(0).Rows.Add(foundRow)
                End If

                '-----------------------------------------------------
                'Aktualizuj
                If ParL_DataBaseType = "MS ACCESS" Then
                    Try
                        dbConnectionParametry.Open()
                    Catch Ex As System.Exception
                        ViewInLog(Ex, MyForm.Name(), Nothing)
                    End Try

                    If dbConnectionParametry.State = ConnectionState.Open Then
                        oWystapilConcurrencyException = False
                        '------------------------------------------------------------
                        Try
                            dbDataAdapterParametry.Update(DataSetParametry, "TABELA_Parametry")
                        Catch Ex As System.Exception
                            ViewInLog(Ex, MyForm.Name(), Nothing)
                        End Try
                        '------------------------------------------------------------
                        If oWystapilConcurrencyException = True Then
                            dbDataAdapterParametry.Fill(DataSetParametry)
                        End If
                        dbConnectionParametry.Close()
                    End If
                Else
                    Try
                        sqlConnectionParametry.Open()
                    Catch Ex As System.Exception
                        ViewInLog(Ex, MyForm.Name(), Nothing)
                    End Try

                    If sqlConnectionParametry.State = ConnectionState.Open Then
                        oWystapilConcurrencyException = False
                        '------------------------------------------------------------
                        Try
                            sqlDataAdapterParametry.Update(DataSetParametry, "TABELA_Parametry")
                        Catch Ex As System.Exception
                            ViewInLog(Ex, MyForm.Name(), Nothing)
                        End Try
                        '------------------------------------------------------------
                        If oWystapilConcurrencyException = True Then
                            sqlDataAdapterParametry.Fill(DataSetParametry)
                        End If
                        sqlConnectionParametry.Close()
                    End If
                End If
                '-----------------------------------------------------
            Catch Ex As System.Exception
                ViewInLog(Ex, MyForm.Name(), Nothing)
            End Try
        End If
    End Sub


End Module
