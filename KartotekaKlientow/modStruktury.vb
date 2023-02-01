Module modStruktury
    Public oAktualnaWersjaBazyDanych As Integer = 520
    Public oConnectionString As String = ""

    '-------------------------------------------------
    'pliki .MDB
    Public oPlikKartoteki As String = "KartotekaKlientow.mdb" 'kartoteki, slowniki
    Public oPlik As String = ""
    Public oBazaDanych As String = "KartotekaKlientowPryzmat"               'kartoteki, slowniki

    'tabele z danymi - Kartoteki
    Public KatKlienciMDB As String = oPlikKartoteki         'slownik JM
    Public KatKlienciBD As String = oBazaDanych
    Public KatKlienciTabela As String = "Klienci"








    Public KatKontaktyMDB As String = oPlikKartoteki        'kontakty handlowe
    Public KatKontaktyBD As String = oBazaDanych
    Public KatKontaktyTabela As String = "Kontakty"

    Public KatRodzajeKontaktowMDB As String = oPlikKartoteki        'kontakty handlowe
    Public KatRodzajeKontaktowBD As String = oBazaDanych
    Public KatRodzajeKontaktowTabela As String = "RodzajeKontaktow"

    Public KatOsobyKontaktoweMDB As String = oPlikKartoteki 'osoby kontaktowe u klientw
    Public KatOsobyKontaktoweBD As String = oBazaDanych
    Public KatOsobyKontaktoweTabela As String = "Osoby"

    Public KatObrotyMDB As String = oPlikKartoteki 'obroty 
    Public KatObrotyBD As String = oBazaDanych
    Public KatObrotyTabela As String = "Obroty"

    Public KatObrotyMiesMDB As String = oPlikKartoteki 'obroty 
    Public KatObrotyMiesBD As String = oBazaDanych
    Public KatObrotyMiesTabela As String = "ObrotyMies"

    Public KatAkcjeOpisMDB As String = oPlikKartoteki 'obroty 
    Public KatAkcjeOpisBD As String = oBazaDanych
    Public KatAkcjeOpisTabela As String = "AkcjeOpis"

    Public KatAkcjeSpecMDB As String = oPlikKartoteki 'obroty 
    Public KatAkcjeSpecBD As String = oBazaDanych
    Public KatAkcjeSpecTabela As String = "AkcjeSpec"

    Public KatAnalizyZakupuMDB As String = oPlikKartoteki 'obroty 
    Public KatAnalizyZakupuBD As String = oBazaDanych
    Public KatAnalizyZakupuTabela As String = "AnalizyZakupu"

    Public KatDaneDoRaportuMDB As String = oPlikKartoteki 'obroty 
    Public KatDaneDoRaportuBD As String = oBazaDanych
    Public KatDaneDoRaportuTabela As String = "DaneDoRaportu"

    Public KatWersjaMDB As String = oPlikKartoteki 'Wersja katalogu
    Public KatWersjaBD As String = oBazaDanych
    Public KatWersjaTabela As String = "Wersja"




    ' ### ###  #####  ######  #####  #####  ##### ####### ### ### ###
    ' ### ##  ### ### ### ###  ###  ### ###  ###  ###     ### ##  ###
    ' #####   ####### ######   ###  ### ###  ###  #####   #####   ###
    ' ### ##  ### ### ### ###  ###  ### ###  ###  ###     ### ##  ###
    ' ### ### ### ### ### ###  ###   #####   ###  ####### ### ### ###


    ' Plik KartotekaKlientow.MDB
    '   Tabela          Indeks          Klucz
    ' --------------------------------------------------
    '   Parametry       WgSymboli       PARAMETR
    '   Uzytkownicy     WgSymboli       IDENT
    '   Klienci         WgSymboli       IDENTKLIENTA
    '   Kontakty        WgSymboli       IDENTKLIENTA
    '   RodzajeKontaktow  WgSymboli       IDENTKLIENTA
    '   Osoby           WgSymboli       IDENTKLIENTA
    '   Obroty          WgSymboli       IDENTKLIENTA+Data
    '   ObrotyMies      WgSymboli       IDENTKLIENTA+MIESIAC
    '   Wersja          WgSymboli       IDENT

    '---------------------------------------------------------------------
    'Katalog Parametrow
    Public oParametr As String          'PARAMETR       Text    20
    Public oOpisPat As String           'OPIS           Text    50
    Public oWartoscPar As String        'WARTOSC        Text    50
    Public oWersjaPar As Integer        'WERSJA         Integer














































    '---------------------------------------------------------------------
    'DaneDoRaportu
    '---------------------------------------------------------------------
    'Id Pracownika			String
    'Data Raportu			Data
    'IloscKlBezDrukari		Integer		Iloœc obs³. klientów bez drukarki
    'IloscOfert			Integer		Iloœæ z³o¿onych w tym dniu ofert
    'IloœæTestyWstawione		Integer		??
    'IloœæCzyszczenieKlienci`		Integer		??
    'IloœæDostawy			Integer		??
    'IloœæSkup			Integer		??


    Public ddrPracownik As String        'PRACOWNIK   Text, 10
    Public ddrDataRaportu As String      'DATARAPORTU Text, 7   RRRR-MM
    Public ddrKlBezDrukarki As Integer   'KLBEZDRUKARKI  Integer
    Public ddrOferty As Integer          'OFERTY         Integer
    Public ddrTestyWstawionw As Integer  'TESTYWSTAWIONE Integer
    Public ddrCzyszczenie As Integer     'CZYSZCZENIE    Integer
    Public ddrDostawy As Integer         'DOSTAWY        Integer
    Public ddrSkup As Integer            'SKUP           Integer

    Public ddrKolExcel01 As String            'KolExCEL01           String
    Public ddrKolExcel02 As String            'KolExCEL02           String
    Public ddrKolExcel03 As String            'KolExCEL03           String
    Public ddrKolExcel04 As String            'KolExCEL04           String
    Public ddrKolExcel05 As String            'KolExCEL05           String
    Public ddrKolExcel06 As String            'KolExCEL06           String
    Public ddrKolExcel07 As String            'KolExCEL07           String
    Public ddrKolExcel08 As String            'KolExCEL08           String
    Public ddrKolExcel09 As String            'KolExCEL09           String
    Public ddrKolExcel10 As String            'KolExCEL10           String

    Public ddrExcel01 As Integer            'EXCEL01           Integer
    Public ddrExcel02 As Integer            'EXCEL02           Integer
    Public ddrExcel03 As Integer            'EXCEL03           Integer
    Public ddrExcel04 As Integer            'EXCEL04           Integer
    Public ddrExcel05 As Integer            'EXCEL05           Integer
    Public ddrExcel06 As Integer            'EXCEL06           Integer
    Public ddrExcel07 As Integer            'EXCEL07           Integer
    Public ddrExcel08 As Integer            'EXCEL08           Integer
    Public ddrExcel09 As Integer            'EXCEL09           Integer
    Public ddrExcel10 As Integer            'EXCEL10           Integer

    Public ddrWersja As Integer          'WERSJA         Integer

    Public sConnectionDaneDoRaportu As String = ConnectionStrings()
    Public HQConnectionDaneDoRaportu As String = HQConnectionStrings()

    Public sSelectSQLDaneDoRaportu As String = "SELECT " &
                                                     "PRACOWNIK," &
                                                     "DATARAPORTU," &
                                                     "KLBEZDRUKARKI," &
                                                     "OFERTY," &
                                                     "TESTYWSTAWIONE," &
                                                     "CZYSZCZENIE," &
                                                     "DOSTAWY," &
                                                     "SKUP," &
                                                     "KOLEXCEL01," &
                                                     "KOLEXCEL02," &
                                                     "KOLEXCEL03," &
                                                     "KOLEXCEL04," &
                                                     "KOLEXCEL05," &
                                                     "KOLEXCEL06," &
                                                     "KOLEXCEL07," &
                                                     "KOLEXCEL08," &
                                                     "KOLEXCEL09," &
                                                     "KOLEXCEL10," &
                                                     "EXCEL01," &
                                                     "EXCEL02," &
                                                     "EXCEL03," &
                                                     "EXCEL04," &
                                                     "EXCEL05," &
                                                     "EXCEL06," &
                                                     "EXCEL07," &
                                                     "EXCEL08," &
                                                     "EXCEL09," &
                                                     "EXCEL10," &
                                                     "WERSJA " &
                                               "FROM DaneDoRaportu"

    Public sDeleteSQLDaneDoRaportu As String = "DELETE FROM DaneDoRaportu " &
                                        "WHERE (PRACOWNIK=@orygSymbol) AND " &
                                              "(DATARAPORTU=@orygData) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLDaneDoRaportu As String = "UPDATE DaneDoRaportu SET " &
                                                     "KLBEZDRUKARKI=@dKLBEZDRUKARKI," &
                                                     "OFERTY=@dOFERTY," &
                                                     "TESTYWSTAWIONE=@dTESTYWSTAWIONE," &
                                                     "CZYSZCZENIE=@dCZYSZCZENIE," &
                                                     "DOSTAWY=@dDOSTAWY," &
                                                     "SKUP=@dSKUP," &
                                                     "KOLEXCEL01=@dKOLEXCEL01," &
                                                     "KOLEXCEL02=@dKOLEXCEL02," &
                                                     "KOLEXCEL03=@dKOLEXCEL03," &
                                                     "KOLEXCEL04=@dKOLEXCEL04," &
                                                     "KOLEXCEL05=@dKOLEXCEL05," &
                                                     "KOLEXCEL06=@dKOLEXCEL06," &
                                                     "KOLEXCEL07=@dKOLEXCEL07," &
                                                     "KOLEXCEL08=@dKOLEXCEL08," &
                                                     "KOLEXCEL09=@dKOLEXCEL09," &
                                                     "KOLEXCEL10=@dKOLEXCEL10," &
                                                     "EXCEL01=@dEXCEL01," &
                                                     "EXCEL02=@dEXCEL02," &
                                                     "EXCEL03=@dEXCEL03," &
                                                     "EXCEL04=@dEXCEL04," &
                                                     "EXCEL05=@dEXCEL05," &
                                                     "EXCEL06=@dEXCEL06," &
                                                     "EXCEL07=@dEXCEL07," &
                                                     "EXCEL08=@dEXCEL08," &
                                                     "EXCEL09=@dEXCEL09," &
                                                     "EXCEL10=@dEXCEL10," &
                                                     "WERSJA=@dWERSJA " &
                                        "WHERE (PRACOWNIK=@orygSymbol) AND " &
                                              "(DATARAPORTU=@orygData) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sInsertSQLDaneDoRaportu As String = "INSERT INTO DaneDoRaportu (" &
                                                     "PRACOWNIK," &
                                                     "DATARAPORTU," &
                                                     "KLBEZDRUKARKI," &
                                                     "OFERTY," &
                                                     "TESTYWSTAWIONE," &
                                                     "CZYSZCZENIE," &
                                                     "DOSTAWY," &
                                                     "SKUP," &
                                                     "KOLEXCEL01," &
                                                     "KOLEXCEL02," &
                                                     "KOLEXCEL03," &
                                                     "KOLEXCEL04," &
                                                     "KOLEXCEL05," &
                                                     "KOLEXCEL06," &
                                                     "KOLEXCEL07," &
                                                     "KOLEXCEL08," &
                                                     "KOLEXCEL09," &
                                                     "KOLEXCEL10," &
                                                     "EXCEL01," &
                                                     "EXCEL02," &
                                                     "EXCEL03," &
                                                     "EXCEL04," &
                                                     "EXCEL05," &
                                                     "EXCEL06," &
                                                     "EXCEL07," &
                                                     "EXCEL08," &
                                                     "EXCEL09," &
                                                     "EXCEL10," &
                                                     "WERSJA " &
                                                        ") " &
                                                "VALUES  (" &
                                                     "@dPRACOWNIK," &
                                                     "@dDATARAPORTU," &
                                                     "@dKLBEZDRUKARKI," &
                                                     "@dOFERTY," &
                                                     "@dTESTYWSTAWIONE," &
                                                     "@dCZYSZCZENIE," &
                                                     "@dDOSTAWY," &
                                                     "@dSKUP," &
                                                     "@dKOLEXCEL01," &
                                                     "@dKOLEXCEL02," &
                                                     "@dKOLEXCEL03," &
                                                     "@dKOLEXCEL04," &
                                                     "@dKOLEXCEL05," &
                                                     "@dKOLEXCEL06," &
                                                     "@dKOLEXCEL07," &
                                                     "@dKOLEXCEL08," &
                                                     "@dKOLEXCEL09," &
                                                     "@dKOLEXCEL10," &
                                                     "@dEXCEL01," &
                                                     "@dEXCEL02," &
                                                     "@dEXCEL03," &
                                                     "@dEXCEL04," &
                                                     "@dEXCEL05," &
                                                     "@dEXCEL06," &
                                                     "@dEXCEL07," &
                                                     "@dEXCEL08," &
                                                     "@dEXCEL09," &
                                                     "@dEXCEL10," &
                                                     "@dWERSJA " &
                                                         ")"

    'SQLDeleteDaneDoRaportu(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateDaneDoRaportu(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertDaneDoRaportu(sqlCommandInsert, sqlDataAdapter)


    Public Sub SQLDeleteDaneDoRaportu(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 10, "PRACOWNIK")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygData", SqlDbType.VarChar, 10, "DATARAPORTU")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original

        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    'Public ddrPracownik As String        'PRACOWNIK   Text, 10
    'Public ddrDataRaportu As String      'DATARAPORTU Text, 7   RRRR-MM
    'Public ddrKlBezDrukarki As Integer   'KLBEZDRUKARKI  Integer
    'Public ddrOferty As Integer          'OFERTY         Integer
    'Public ddrTestyWstawionw As Integer  'TESTYWSTAWIONE Integer
    'Public ddrCzyszczenie As Integer     'CZYSZCZENIE    Integer
    'Public ddrDostawy As Integer         'DOSTAWY        Integer
    'Public ddrSkup As Integer            'SKUP           Integer

    'Public ddrKolExcel01 As String            'KolExCEL01           String
    'Public ddrKolExcel02 As String            'KolExCEL02           String
    'Public ddrKolExcel03 As String            'KolExCEL03           String
    'Public ddrKolExcel04 As String            'KolExCEL04           String
    'Public ddrKolExcel05 As String            'KolExCEL05           String
    'Public ddrKolExcel06 As String            'KolExCEL06           String
    'Public ddrKolExcel07 As String            'KolExCEL07           String
    'Public ddrKolExcel08 As String            'KolExCEL08           String
    'Public ddrKolExcel09 As String            'KolExCEL09           String
    'Public ddrKolExcel10 As String            'KolExCEL10           String

    'Public ddrExcel01 As Integer            'EXCEL01           Integer
    'Public ddrExcel02 As Integer            'EXCEL02           Integer
    'Public ddrExcel03 As Integer            'EXCEL03           Integer
    'Public ddrExcel04 As Integer            'EXCEL04           Integer
    'Public ddrExcel05 As Integer            'EXCEL05           Integer
    'Public ddrExcel06 As Integer            'EXCEL06           Integer
    'Public ddrExcel07 As Integer            'EXCEL07           Integer
    'Public ddrExcel08 As Integer            'EXCEL08           Integer
    'Public ddrExcel09 As Integer            'EXCEL09           Integer
    'Public ddrExcel10 As Integer            'EXCEL10           Integer

    'Public ddrWersja As Integer          'WERSJA         Integer

    Public Sub SQLUpdateDaneDoRaportu(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        'sqlCommandUpdate.Parameters.Add("@dPracownik", SqlDbType.VarChar, 10, "PRACOWNIK")
        'sqlCommandUpdate.Parameters.Add("@dDataRaportu", SqlDbType.VarChar, 10, "DATARAPORTU")
        sqlCommandUpdate.Parameters.Add("@dKlBezDrukarki", SqlDbType.Int, Nothing, "KLBEZDRUKARKI")
        sqlCommandUpdate.Parameters.Add("@dOferty", SqlDbType.Int, Nothing, "OFERTY")
        sqlCommandUpdate.Parameters.Add("@dTestyWstawione", SqlDbType.Int, Nothing, "TESTYWSTAWIONE")
        sqlCommandUpdate.Parameters.Add("@dCzyszczenie", SqlDbType.Int, Nothing, "CZYSZCZENIE")
        sqlCommandUpdate.Parameters.Add("@dDostawy", SqlDbType.Int, Nothing, "Dostawy")
        sqlCommandUpdate.Parameters.Add("@dSkup", SqlDbType.Int, Nothing, "SKUP")

        sqlCommandUpdate.Parameters.Add("@dKolExcel01", SqlDbType.VarChar, 50, "KOLEXCEL01")
        sqlCommandUpdate.Parameters.Add("@dKolExcel02", SqlDbType.VarChar, 50, "KOLEXCEL02")
        sqlCommandUpdate.Parameters.Add("@dKolExcel03", SqlDbType.VarChar, 50, "KOLEXCEL03")
        sqlCommandUpdate.Parameters.Add("@dKolExcel04", SqlDbType.VarChar, 50, "KOLEXCEL04")
        sqlCommandUpdate.Parameters.Add("@dKolExcel05", SqlDbType.VarChar, 50, "KOLEXCEL05")
        sqlCommandUpdate.Parameters.Add("@dKolExcel06", SqlDbType.VarChar, 50, "KOLEXCEL06")
        sqlCommandUpdate.Parameters.Add("@dKolExcel07", SqlDbType.VarChar, 50, "KOLEXCEL07")
        sqlCommandUpdate.Parameters.Add("@dKolExcel08", SqlDbType.VarChar, 50, "KOLEXCEL08")
        sqlCommandUpdate.Parameters.Add("@dKolExcel09", SqlDbType.VarChar, 50, "KOLEXCEL09")
        sqlCommandUpdate.Parameters.Add("@dKolExcel10", SqlDbType.VarChar, 50, "KOLEXCEL10")

        sqlCommandUpdate.Parameters.Add("@dExcel01", SqlDbType.Int, Nothing, "EXCEL01")
        sqlCommandUpdate.Parameters.Add("@dExcel02", SqlDbType.Int, Nothing, "EXCEL02")
        sqlCommandUpdate.Parameters.Add("@dExcel03", SqlDbType.Int, Nothing, "EXCEL03")
        sqlCommandUpdate.Parameters.Add("@dExcel04", SqlDbType.Int, Nothing, "EXCEL04")
        sqlCommandUpdate.Parameters.Add("@dExcel05", SqlDbType.Int, Nothing, "EXCEL05")
        sqlCommandUpdate.Parameters.Add("@dExcel06", SqlDbType.Int, Nothing, "EXCEL06")
        sqlCommandUpdate.Parameters.Add("@dExcel07", SqlDbType.Int, Nothing, "EXCEL07")
        sqlCommandUpdate.Parameters.Add("@dExcel08", SqlDbType.Int, Nothing, "EXCEL08")
        sqlCommandUpdate.Parameters.Add("@dExcel09", SqlDbType.Int, Nothing, "EXCEL09")
        sqlCommandUpdate.Parameters.Add("@dExcel10", SqlDbType.Int, Nothing, "EXCEL10")

        sqlCommandUpdate.Parameters.Add("@dWersja", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 10, "PRACOWNIK")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygData", SqlDbType.VarChar, 10, "DATARAPORTU")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original

        sqlDataAdapter.UpdateCommand = sqlCommandUpdate
    End Sub

    Public Sub SQLInsertDaneDoRaportu(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter

        sqlParam = sqlCommandInsert.Parameters.Add("@dPracownik", SqlDbType.VarChar, 10, "PRACOWNIK")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dDataRaportu", SqlDbType.VarChar, 10, "DATARAPORTU")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dKlBezDrukarki", SqlDbType.Int, Nothing, "KLBEZDRUKARKI")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dOferty", SqlDbType.Int, Nothing, "OFERTY")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dTestyWstawione", SqlDbType.Int, Nothing, "TESTYWSTAWIONE")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dCzyszczenie", SqlDbType.Int, Nothing, "CZYSZCZENIE")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dDostawy", SqlDbType.Int, Nothing, "Dostawy")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dSkup", SqlDbType.Int, Nothing, "SKUP")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@dKolExcel01", SqlDbType.VarChar, 50, "KOLEXCEL01")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dKolExcel02", SqlDbType.VarChar, 50, "KOLEXCEL02")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dKolExcel03", SqlDbType.VarChar, 50, "KOLEXCEL03")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dKolExcel04", SqlDbType.VarChar, 50, "KOLEXCEL04")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dKolExcel05", SqlDbType.VarChar, 50, "KOLEXCEL05")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dKolExcel06", SqlDbType.VarChar, 50, "KOLEXCEL06")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dKolExcel07", SqlDbType.VarChar, 50, "KOLEXCEL07")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dKolExcel08", SqlDbType.VarChar, 50, "KOLEXCEL08")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dKolExcel09", SqlDbType.VarChar, 50, "KOLEXCEL09")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dKolExcel10", SqlDbType.VarChar, 50, "KOLEXCEL10")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@dExcel01", SqlDbType.Int, Nothing, "EXCEL01")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dExcel02", SqlDbType.Int, Nothing, "EXCEL02")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dExcel03", SqlDbType.Int, Nothing, "EXCEL03")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dExcel04", SqlDbType.Int, Nothing, "EXCEL04")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dExcel05", SqlDbType.Int, Nothing, "EXCEL05")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dExcel06", SqlDbType.Int, Nothing, "EXCEL06")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dExcel07", SqlDbType.Int, Nothing, "EXCEL07")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dExcel08", SqlDbType.Int, Nothing, "EXCEL08")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dExcel09", SqlDbType.Int, Nothing, "EXCEL09")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@dExcel10", SqlDbType.Int, Nothing, "EXCEL10")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@dWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub

    Public Sub DBDeleteDaneDoRaportu(ByRef dbCommandDelete As OleDb.OleDbCommand,
                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
        Dim dbParam As OleDb.OleDbParameter

        dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 10, "PRACOWNIK")
        dbParam.SourceVersion = DataRowVersion.Original
        dbParam = dbCommandDelete.Parameters.Add("@orygData", OleDb.OleDbType.VarChar, 10, "DATARAPORTU")
        dbParam.SourceVersion = DataRowVersion.Original
        dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
        dbParam.SourceVersion = DataRowVersion.Original

        dbDataAdapter.DeleteCommand = dbCommandDelete
    End Sub

    Public Sub DBUpdateDaneDoRaportu(ByRef dbCommandUpdate As OleDb.OleDbCommand,
                               ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
        Dim dbParam As New OleDb.OleDbParameter

        'dbCommandUpdate.Parameters.Add("@dPracownik", OleDb.OleDbType.VarChar, 10, "PRACOWNIK")
        'dbCommandUpdate.Parameters.Add("@dDataRaportu", OleDb.OleDbType.VarChar, 10, "DATARAPORTU")
        dbCommandUpdate.Parameters.Add("@dKlBezDrukarki", OleDb.OleDbType.Integer, Nothing, "KLBEZDRUKARKI")
        dbCommandUpdate.Parameters.Add("@dOferty", OleDb.OleDbType.Integer, Nothing, "OFERTY")
        dbCommandUpdate.Parameters.Add("@dTestyWstawione", OleDb.OleDbType.Integer, Nothing, "TESTYWSTAWIONE")
        dbCommandUpdate.Parameters.Add("@dCzyszczenie", OleDb.OleDbType.Integer, Nothing, "CZYSZCZENIE")
        dbCommandUpdate.Parameters.Add("@dDostawy", OleDb.OleDbType.Integer, Nothing, "Dostawy")
        dbCommandUpdate.Parameters.Add("@dSkup", OleDb.OleDbType.Integer, Nothing, "SKUP")

        dbCommandUpdate.Parameters.Add("@dKolExcel01", OleDb.OleDbType.VarChar, 50, "KOLEXCEL01")
        dbCommandUpdate.Parameters.Add("@dKolExcel02", OleDb.OleDbType.VarChar, 50, "KOLEXCEL02")
        dbCommandUpdate.Parameters.Add("@dKolExcel03", OleDb.OleDbType.VarChar, 50, "KOLEXCEL03")
        dbCommandUpdate.Parameters.Add("@dKolExcel04", OleDb.OleDbType.VarChar, 50, "KOLEXCEL04")
        dbCommandUpdate.Parameters.Add("@dKolExcel05", OleDb.OleDbType.VarChar, 50, "KOLEXCEL05")
        dbCommandUpdate.Parameters.Add("@dKolExcel06", OleDb.OleDbType.VarChar, 50, "KOLEXCEL06")
        dbCommandUpdate.Parameters.Add("@dKolExcel07", OleDb.OleDbType.VarChar, 50, "KOLEXCEL07")
        dbCommandUpdate.Parameters.Add("@dKolExcel08", OleDb.OleDbType.VarChar, 50, "KOLEXCEL08")
        dbCommandUpdate.Parameters.Add("@dKolExcel09", OleDb.OleDbType.VarChar, 50, "KOLEXCEL09")
        dbCommandUpdate.Parameters.Add("@dKolExcel10", OleDb.OleDbType.VarChar, 50, "KOLEXCEL10")

        dbCommandUpdate.Parameters.Add("@dExcel01", OleDb.OleDbType.Integer, Nothing, "EXCEL01")
        dbCommandUpdate.Parameters.Add("@dExcel02", OleDb.OleDbType.Integer, Nothing, "EXCEL02")
        dbCommandUpdate.Parameters.Add("@dExcel03", OleDb.OleDbType.Integer, Nothing, "EXCEL03")
        dbCommandUpdate.Parameters.Add("@dExcel04", OleDb.OleDbType.Integer, Nothing, "EXCEL04")
        dbCommandUpdate.Parameters.Add("@dExcel05", OleDb.OleDbType.Integer, Nothing, "EXCEL05")
        dbCommandUpdate.Parameters.Add("@dExcel06", OleDb.OleDbType.Integer, Nothing, "EXCEL06")
        dbCommandUpdate.Parameters.Add("@dExcel07", OleDb.OleDbType.Integer, Nothing, "EXCEL07")
        dbCommandUpdate.Parameters.Add("@dExcel08", OleDb.OleDbType.Integer, Nothing, "EXCEL08")
        dbCommandUpdate.Parameters.Add("@dExcel09", OleDb.OleDbType.Integer, Nothing, "EXCEL09")
        dbCommandUpdate.Parameters.Add("@dExcel10", OleDb.OleDbType.Integer, Nothing, "EXCEL10")

        dbCommandUpdate.Parameters.Add("@dWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")

        'parametry filtrowania
        dbParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 10, "PRACOWNIK")
        dbParam.SourceVersion = DataRowVersion.Original
        dbParam = dbCommandUpdate.Parameters.Add("@orygData", OleDb.OleDbType.VarChar, 10, "DATARAPORTU")
        dbParam.SourceVersion = DataRowVersion.Original
        dbParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
        dbParam.SourceVersion = DataRowVersion.Original

        dbDataAdapter.UpdateCommand = dbCommandUpdate
    End Sub

    Public Sub DBInsertDaneDoRaportu(ByRef dbCommandInsert As OleDb.OleDbCommand,
                               ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
        Dim dbParam As New OleDb.OleDbParameter

        dbParam = dbCommandInsert.Parameters.Add("@dPracownik", OleDb.OleDbType.VarChar, 10, "PRACOWNIK")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dDataRaportu", OleDb.OleDbType.VarChar, 10, "DATARAPORTU")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dKlBezDrukarki", OleDb.OleDbType.Integer, Nothing, "KLBEZDRUKARKI")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dOferty", OleDb.OleDbType.Integer, Nothing, "OFERTY")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dTestyWstawione", OleDb.OleDbType.Integer, Nothing, "TESTYWSTAWIONE")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dCzyszczenie", OleDb.OleDbType.Integer, Nothing, "CZYSZCZENIE")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dDostawy", OleDb.OleDbType.Integer, Nothing, "Dostawy")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dSkup", OleDb.OleDbType.Integer, Nothing, "SKUP")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@dKolExcel01", OleDb.OleDbType.VarChar, 50, "KOLEXCEL01")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dKolExcel02", OleDb.OleDbType.VarChar, 50, "KOLEXCEL02")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dKolExcel03", OleDb.OleDbType.VarChar, 50, "KOLEXCEL03")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dKolExcel04", OleDb.OleDbType.VarChar, 50, "KOLEXCEL04")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dKolExcel05", OleDb.OleDbType.VarChar, 50, "KOLEXCEL05")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dKolExcel06", OleDb.OleDbType.VarChar, 50, "KOLEXCEL06")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dKolExcel07", OleDb.OleDbType.VarChar, 50, "KOLEXCEL07")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dKolExcel08", OleDb.OleDbType.VarChar, 50, "KOLEXCEL08")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dKolExcel09", OleDb.OleDbType.VarChar, 50, "KOLEXCEL09")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dKolExcel10", OleDb.OleDbType.VarChar, 50, "KOLEXCEL10")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@dExcel01", OleDb.OleDbType.Integer, Nothing, "EXCEL01")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dExcel02", OleDb.OleDbType.Integer, Nothing, "EXCEL02")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dExcel03", OleDb.OleDbType.Integer, Nothing, "EXCEL03")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dExcel04", OleDb.OleDbType.Integer, Nothing, "EXCEL04")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dExcel05", OleDb.OleDbType.Integer, Nothing, "EXCEL05")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dExcel06", OleDb.OleDbType.Integer, Nothing, "EXCEL06")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dExcel07", OleDb.OleDbType.Integer, Nothing, "EXCEL07")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dExcel08", OleDb.OleDbType.Integer, Nothing, "EXCEL08")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dExcel09", OleDb.OleDbType.Integer, Nothing, "EXCEL09")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@dExcel10", OleDb.OleDbType.Integer, Nothing, "EXCEL10")
        dbParam.SourceVersion = DataRowVersion.Current


        dbParam = dbCommandInsert.Parameters.Add("@dWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
        dbParam.SourceVersion = DataRowVersion.Current

        dbDataAdapter.InsertCommand = dbCommandInsert
    End Sub












































    '---------------------------------------------------------------------
    'Wersja
    Public oWersjaBazyDanych As Integer     'IDENT     Integer

    Public sConnectionWersja As String = ConnectionStrings()
    Public HQConnectionWersja As String = HQConnectionStrings()

    Public sSelectSQLWersja As String = "SELECT * FROM Wersja"

    Public sUpdateSQLWersja As String = "UPDATE Wersja SET " &
                                                 "IDENT=@oIdent"

    Public sInsertSQLWersja As String = "INSERT INTO Wersja (" &
                                                 "IDENT " &
                                                 ") " &
                                        "VALUES  (" &
                                                 "@oIdent " &
                                                 ")"


    '**************************************
    '** definiowanie Connections Strings
    '**************************************

    Public Function DataBaseConnection() As String
        Dim ConnString As String = ""
        If ParL_DataBaseType = "MS ACCESS" Then
            'w ACCESS wersja przechowywana jest w Tabeli WERSJA
            'w polu IDENT, integer
            ConnString = ConnStringAccess()
        Else
            'w SQL wersja i nazwa programu przechowywana jest w procedurach wbudowanych
            ConnString = ConnStringMSSQL()
        End If
        Return ConnString
    End Function


    Public Function ConnectionStrings() As String
        Dim cs As String = ""
        If ParL_DataBaseType = "MS ACCESS" Then
            cs = ConnStringAccess()
        Else
            cs = ConnStringMSSQL()
        End If
        '--------------------------------------
        sConnectionUzytkownicy = cs
        sConnectionKlienci = cs
        sConnectionKontakty = cs
        sConnectionRodzajeKontaktow = cs
        sConnectionOsoby = cs
        sConnectionObroty = cs
        sConnectionObrotyMies = cs
        sConnectionAkcjeOpis = cs
        sConnectionAkcjeSpec = cs
        sConnectionAnalizyZakupu = cs
        sConnectionDaneDoRaportu = cs
        sConnectionSlownikCoDalej = cs
        sConnectionCoDalej = cs
        sConnectionBranze = cs
        sConnectionPodBranze = cs
        '----------------------------------------
        Return cs
    End Function


    Public Function ConnStringAccess() As String
        If OSArchitectureIs64bit() Then
            ''ACE
            'If you are an application developer using OLEDB, set the Provider argument 
            'of the ConnectionString property to “Microsoft.ACE.OLEDB.12.0”. 
            Return ("Provider=""Microsoft.ACE.OLEDB.12.0"";" &
                "Data Source=""" & ParL_KatZbiorow & "\" & oPlikKartoteki & """;" &
                "Persist Security Info=False")

            'If you are an application developer using ODBC to connect to Microsoft Office Access data, 
            'set the Connection String to “Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=path to mdb/accdb file”
            'If you are an application developer using ODBC to connect to Microsoft Office Excel data, 
            'set the Connection String to “Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=path to xls/xlsx/xlsm/xlsb file”
            'Return ("Driver={Microsoft Access Driver (*.mdb, *.accdb)};" & _
            '        "DBQ=""" & ParL_KatZbiorow & "\" & oPlikKartoteki & """ ")
        Else
            'Jet 4,0
            Return ("Provider=""Microsoft.Jet.OLEDB.4.0"";" &
                "Data Source=""" & ParL_KatZbiorow & "\" & oPlikKartoteki & """;" &
                "Persist Security Info=False")
        End If
    End Function

    Public Function ConnStringMSSQL() As String
        Dim MyCS As String

        'baza SQL Serwer 2000
        'MyCS = "Data Source=" & oLocalHostName & ";" & _
        '       "Integrated Security=SSPI;" & _
        '       "Initial Catalog=" & oBazaDanychSlowniki & ""

        If ParL_Autentykacja = "Windows" Then
            MyCS = "Data Source=" & ParL_Serwer & ";" &
                   "Integrated Security=SSPI;" &
                   "Initial Catalog=" & ParL_DataBase & ""
        Else
            'MyCS = "Data Source=" & ParL_Serwer & ";" & _
            '       "User ID=" & ParL_Login & ";" & _
            '       "Integrated Security=SSPI;" & _
            '       "Initial Catalog=" & oBazaDanychSlowniki & ""

            'MyCS = "Provider=SQLOLEDB.1;" & _
            '       "Persist Security Info=False;" & _
            '       "User ID=WCz;" & _
            '       "Initial Catalog=Softart;" & _
            '       "Data Source=SERWER1;" & _
            '       "Use Procedure for Prepare=1;" & _
            '       "Auto Translate=True;" & _
            '       "Packet Size=4096;" & _
            '       "Workstation ID=WCZNOTEBOOK;" & _
            '       "Use Encryption for Data=False;" & _
            '       "Tag with column collation when possible=False"

            MyCS = "Persist Security Info=False;" &
                   "User ID=" & ParL_Login & ";" &
                   "Password=" & ParL_Password & ";" &
                   "Initial Catalog=" & ParL_DataBase & ";" &
                   "Data Source=" & ParL_Serwer & ";" &
                   "Packet Size=4096"
        End If

        Return (MyCS)
    End Function


    '**************************************
    '** sprawdzanie srodowiska
    '**************************************

    Public Function SrodowiskoOK() As Boolean
        If Len(Dir(ParL_KatZbiorow & "\" & oPlikKartoteki)) Then
            Return (True)
        Else
            MessageBox.Show("W podanym katalogu bazy danych programu" & vbCrLf & ParL_KatZbiorow & vbCrLf & "Nie znalaz³em prawid³owych plików z danymi programu...",
                "Sprawdzanie œrodowiska",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1)
            Return (False)
        End If
    End Function







    '**************************************
    '** sprawdzanie wersji ACCESS/MSSQL
    '**************************************

    Public Function WersjaBazyDanych() As Integer
        If ParL_DataBaseType = "MS ACCESS" Then
            'w ACCESS wersja przechowywana jest w Tabeli WERSJA
            'w polu IDENT, integer
            Return WersjaBazyDanychACCESS()
        Else
            'w SQL wersja i nazwa programu przechowywana jest w procedurach wbudowanych
            Return WersjaBazyDanychMSSQL()
        End If
    End Function


    Public Function WersjaBazyDanychACCESS() As Integer
        Dim sConnectionWersja As String = ConnStringAccess()

        Dim dbConnectionWersja As New OleDb.OleDbConnection(sConnectionWersja)
        Dim dbCommandSelectWersja As New OleDb.OleDbCommand(sSelectSQLWersja, dbConnectionWersja)
        Dim dbDataAdapterWersja As New OleDb.OleDbDataAdapter(dbCommandSelectWersja)
        Dim DataSetWersja As New DataSet
        Dim DataViewWersja As New DataView

        Dim oWersjaDB As Integer = 0

        DataSetWersja.Locale = New System.Globalization.CultureInfo("pl-PL")
        Dim MapowanieTabeli As System.Data.Common.DataTableMapping
        dbDataAdapterWersja.TableMappings.Clear()
        MapowanieTabeli = dbDataAdapterWersja.TableMappings.Add("Table", "TABELA_Wersja")
        '------------------------------------------
        'wypelnij DATASET
        oWersjaBazyDanych = 0
        Try
            dbConnectionWersja.Open()
        Catch Ex As System.Exception
        End Try

        If dbConnectionWersja.State = ConnectionState.Open Then
            Try
                dbDataAdapterWersja.FillSchema(DataSetWersja, SchemaType.Mapped, "TABELA_Wersja")
                dbDataAdapterWersja.Fill(DataSetWersja)
                dbConnectionWersja.Close()
                'definiuj DataView
                DataViewWersja = New DataView(DataSetWersja.Tables(0))
                If DataViewWersja.Count > 0 Then
                    Dim ii As Integer
                    For ii = 0 To DataViewWersja.Count - 1
                        oWersjaDB = DataViewWersja.Item(ii).Item(0)
                    Next
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (oWersjaDB)
    End Function

    Public Function WersjaBazyDanychMSSQL() As Int16
        Dim oWersjaDB As Int16 = 0
        Dim oNazwaDB As String = ""
        '---------------------------------------
        'sprawdz czy to prawid³owa baza danych SoftArt -  ma procedure softart_nazwaprogramu
        Dim ir As Long = 0
        Dim OK As Boolean = False
        Dim Odpowiedz As String = ""

        Dim qConnectionString As String = ConnectionStrings()
        Dim qSelectSQL As String = "SELECT * FROM sysobjects " &
                                   "WHERE xtype='P'  " &
                                   "ORDER BY name"
        Dim qConnection As New SqlClient.SqlConnection(qConnectionString)
        Dim qCommandSelect As New SqlClient.SqlCommand(qSelectSQL, qConnection)
        Dim qDataAdapter As New SqlClient.SqlDataAdapter(qCommandSelect)
        Dim qDataSetTabela As New DataSet
        Dim qDataViewTabela As New DataView

        Dim MapowanieTabeli As System.Data.Common.DataTableMapping
        MapowanieTabeli = qDataAdapter.TableMappings.Add("Table", "TABELA")
        qDataSetTabela.Locale = New System.Globalization.CultureInfo("pl-PL")
        '------------------------------------------
        'wypelnij DATASET
        Try
            qConnection.Open()
        Catch Ex As System.Exception
            'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
            '    System.Windows.Forms.MessageBoxButtons.OK, _
            '    MessageBoxIcon.Information, _
            '    MessageBoxDefaultButton.Button1)
        End Try

        If qConnection.State = ConnectionState.Open Then
            Try
                qDataAdapter.FillSchema(qDataSetTabela, SchemaType.Mapped, "TABELA")
                qDataAdapter.Fill(qDataSetTabela)
                qConnection.Close()

                'definiuj DataView
                qDataViewTabela = New DataView(qDataSetTabela.Tables(0))
                If qDataViewTabela.Count > 0 Then
                    'sprawdz czy to baza danych softart
                    For ir = 0 To qDataViewTabela.Count - 1
                        If qDataViewTabela.Item(ir).Item("NAME") = "softart_nazwaprogramu" Then
                            OK = True
                            Exit For
                        End If
                    Next
                End If
            Catch Ex As System.Exception
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & Ex.Message, "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            End Try

            'sprawdz czy to baza danych programu SOFTART-MAGAZYN
            If OK Then  'to jest baza danych softart
                Dim nConnectionString As String = ConnectionStrings()
                Dim nStoredProc As String = "softart_nazwaprogramu"
                Dim nConnection As New SqlClient.SqlConnection(nConnectionString)
                Dim nCommand As New SqlClient.SqlCommand
                Dim nParam As SqlClient.SqlParameter
                nCommand.CommandText = nStoredProc  'xSelectSQL
                nCommand.CommandType = CommandType.StoredProcedure
                nParam = nCommand.Parameters.Add(New SqlClient.SqlParameter("@name", SqlDbType.VarChar, 50))
                nParam.Direction = ParameterDirection.Output
                nCommand.Connection = nConnection
                '---------------------------------------
                Dim wConnectionString As String = ConnectionStrings()
                Dim wStoredProc As String = "softart_wersjaprogramu"
                Dim wConnection As New SqlClient.SqlConnection(wConnectionString)
                Dim wCommand As New SqlClient.SqlCommand
                Dim wParam As SqlClient.SqlParameter
                wCommand.CommandText = wStoredProc  'xSelectSQL
                wCommand.CommandType = CommandType.StoredProcedure
                wParam = wCommand.Parameters.Add(New SqlClient.SqlParameter("@wersja", SqlDbType.Int, Nothing))
                wParam.Direction = ParameterDirection.Output
                wCommand.Connection = wConnection
                '---------------------------------------
                Try
                    nConnection.Open()
                    nCommand.ExecuteNonQuery()
                    nConnection.Close()
                    'podstaw parametry wyjsciowe
                    oNazwaDB = CType(nCommand.Parameters("@name").Value, String)
                    If oNazwaDB = "KartotekaKlientowPryzmat" Then
                        wConnection.Open()
                        wCommand.ExecuteNonQuery()
                        wConnection.Close()
                        'podstaw parametry wyjsciowe
                        oWersjaDB = CType(wCommand.Parameters("@wersja").Value, Int16)
                    End If

                Catch ex As Exception
                    'MessageBox.Show(ex.Message)
                End Try
            Else
            End If
        End If
        Return (oWersjaDB)
    End Function

    '********************************************************************
    '*** testowanie wyst¹pienia b³êdu wspóbie¿noœci
    '*******************************************************************


    Public Sub OnRowUpdated(ByVal sender As Object, ByVal args As OleDb.OleDbRowUpdatedEventArgs)
        Dim argsKomenda As String = ""
        'sprawdzamy liczbe aktualizowanych rekordow
        If args.RecordsAffected = 0 Then
            Select Case args.StatementType
                Case StatementType.Delete
                    argsKomenda = "DELETE"
                Case StatementType.Insert
                    argsKomenda = "INSERT"
                Case StatementType.Select
                    argsKomenda = "SELECT"
                Case StatementType.Update
                    argsKomenda = "UPDATE"
            End Select

            Dim aTXT As String
            'aTXT = args.Command.ToString
            aTXT = args.Errors.ToString
            aTXT = args.Row.RowError.ToString
            aTXT = args.Status.ToString
            aTXT = args.TableMapping.ToString

            'opis bledu = args.Errors.ToString()
            MessageBox.Show("Podczas wykonywania komendy " & argsKomenda & vbCrLf &
                                        "wyst¹pi³ b³¹d wspólnego dostêpu do bazy danych..." & vbCrLf &
                                        "Aktualizacje rekordu " & args.Row(0) & " zostan¹ utracone" & vbCrLf &
                                        "i pobrany bêdzie bie¿¹cy rekord z bazy danych..." & vbCrLf & vbCrLf &
                                        "Proszê powtórzyæ aktualizacje i spróbowaæ ponownie zapisaæ rekord do bazy.",
                "B³¹d podczas aktualizacji bazy :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)

            MessageBox.Show("Podczas wykonywania komendy " & argsKomenda & vbCrLf &
                                        "Errors=" & args.Errors.ToString & vbCrLf &
                                        "RowError=" & args.Row.RowError.ToString & vbCrLf &
                                        "Status=" & args.Status.ToString & vbCrLf &
                                        "StatementType=" & args.StatementType.ToString & vbCrLf &
                                        "TableMapping=" & args.TableMapping.ToString,
                "B³¹d podczas aktualizacji bazy :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)

            'args.Row.RowError ="Optimistic Concurrency Violation Encountered"
            args.Status = UpdateStatus.SkipCurrentRow

            oWystapilConcurrencyException = True
        End If
    End Sub


    Public Sub sqlOnRowUpdated(ByVal sender As Object, ByVal args As SqlClient.SqlRowUpdatedEventArgs)
        Dim argsKomenda As String = ""
        'sprawdzamy liczbe aktualizowanych rekordow
        If args.RecordsAffected = 0 Then
            Select Case args.StatementType
                Case StatementType.Delete
                    argsKomenda = "DELETE"
                Case StatementType.Insert
                    argsKomenda = "INSERT"
                Case StatementType.Select
                    argsKomenda = "SELECT"
                Case StatementType.Update
                    argsKomenda = "UPDATE"
            End Select

            Dim aTXT As String = ""
            aTXT = IIf(args.Command.ToString Is Nothing, "", args.Command.ToString)
            aTXT = args.Errors.ToString
            aTXT = args.Row.RowError.ToString
            aTXT = args.Status.ToString
            aTXT = args.TableMapping.ToString

            'opis bledu = args.Errors.ToString()
            MessageBox.Show("Podczas wykonywania komendy " & argsKomenda & vbCrLf &
                                        "wyst¹pi³ b³¹d wspólnego dostêpu do bazy danych..." & vbCrLf &
                                        "Aktualizacje rekordu " & args.Row(0) & " zostan¹ utracone" & vbCrLf &
                                        "i pobrany bêdzie bie¿¹cy rekord z bazy danych..." & vbCrLf & vbCrLf &
                                        "Proszê powtórzyæ aktualizacje i spróbowaæ ponownie zapisaæ rekord do bazy.",
                "B³¹d podczas aktualizacji bazy :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)

            MessageBox.Show("Podczas wykonywania komendy " & argsKomenda & vbCrLf &
                                        "Command=" & args.Command.ToString & vbCrLf &
                                        "Errors=" & args.Errors.ToString & vbCrLf &
                                        "RowError=" & args.Row.RowError.ToString & vbCrLf &
                                        "Status=" & args.Status.ToString & vbCrLf &
                                        "StatementType=" & args.StatementType.ToString & vbCrLf &
                                        "TableMapping=" & args.TableMapping.ToString,
                "B³¹d podczas aktualizacji bazy :",
                System.Windows.Forms.MessageBoxButtons.OK,
                MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)

            'args.Row.RowError ="Optimistic Concurrency Violation Encountered"
            args.Status = UpdateStatus.SkipCurrentRow

            oWystapilConcurrencyException = True
        End If
    End Sub







End Module
