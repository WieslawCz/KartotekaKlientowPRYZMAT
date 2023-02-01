Module modStrukturyKlienci


    '---------------------------------------------------------------------
    'Katalog Klientow
    Public oIdentKli As String         'IDENTKLIENTA   Text, 6
    Public oNrTOFIKli As Long          'NRTOFI         Long
    Public oNrTOFITxtKli As String     'NRTOFITXT      Text 500

    Public oKartaPaybackKli As Boolean 'KARTAPAYBACK   Logical
    Public oNryKartPaybackKli As String 'NRYKARTPAYBACK   memo

    Public oNazwa1Kli As String        'NAZWA1         Text, 100
    Public oKodPoczKli As String       'KODPOCZTOWY    Text, 10
    Public oMiejscKli As String        'MIEJSCOWOSC    Text, 5033
    Public oUlicaKli As String         'ULICA          Text, 50
    Public oNumNrDomuKli As Integer    'NUMNRDOMU      INTEGER
    Public oNrDomuKli As String        'NRDOMU         Text, 10
    Public oIDDomuKli As String        'IDDOMU         Text, 10
    Public oNIPKli As String           'NIP            Text, 15
    Public oTelefonKli As String       'TELEFON        Text, 50
    Public oFaxKli As String           'FAX            Text, 50
    Public oeMailKli As String         'EMAIL          Text, 50
    Public oWpisalKli As String        'KTOWPISAL      Text, 10   uzytkownik
    Public oRejonKli As String         'REJKLIENTA     Text, 20   
    Public oPKontaktKli As String      'PKONTAKT       Text, 10   data rrrr-mm-dd

    Public oBranzaKli As String        'BRANZA         Text, 100
    Public oPodBranzaKli As String     'PODBRANZA      Text, 100
    Public oLiczbaZaktrudnionychKli As Integer   'LICZBAZATRUDNIONYCH       INTEGER
    Public oLiczbaPracZdalnieKli As Integer      'LICZBAPRACZDALNIE         INTEGER

    '..............................
    Public oWykazUrzadzenKli As String              'WYKLAZURZADZEN    Memo

    Public oSposobWyboruDostawcyKli As String       'SPOSOBWYBORUDOSTAWCY    Text 30
    Public oZainstalowanoMonitorKli As Boolean      'ZAINSTALOWANOMONITOR    Logical
    Public oLiczbaUrzadzenKli As Integer            'LICZBAURZADZEN     Integer
    Public oLiczbaWydrukowKli As Integer            'LICZBAWYDRUKOW     Integer
    Public oPotencjalDrukuKli As String             'POTENCJALDRUKU     Text 2

    Public oAktZakupMatEkspKli As Boolean           'AKTZAKUPMATEKSP    Logical
    Public oAktRozliczaStronyDrukuKli As Boolean    'AKTROZLICZASTRONYDRUKU    Logical
    Public oAktKorzystaZNajmuKli As Boolean         'AKTKORZYSTAZNAJMU  Logical
    Public oZaintRozliczaniemStronDrukuKli As Boolean   'ZAINTROZLICZANIEMSTRONDRUKU    Logical
    Public oZaintKorzystaniemZNajmuKli As Boolean   'ZAINTKORZYSTANIEMZNAJMU    Logical

    Public oDataWeryfSegmentacjiKli As String          'DATAWERYFSEGMENTACJI  Text 10

    Public oPlanDlugookresowyKli As Integer         'PLANDLUGOOKRESOWY     Integer
    Public oPlanKrotkookresowyKli As Integer        'PLANKROTKOOKRESOWY    Integer
    '......................................
    Public oRodzSprzetuLKli As Boolean 'RODZSPRZETUL Logical
    Public oRodzSprzetuAKli As Boolean 'RODZSPRZETUA Logical
    Public oRodzSprzetuTKli As Boolean 'RODZSPRZETUT Logical
    Public oOfertaTZlozeniaKli As String        'OFERTATZLOZENIA         Text, 4
    Public oOfertaPlikKli As String    'OFERTAPLIK     Text, 100
    Public oUwagikli As String         'UWAGI          Memo

    Public oSkupUwagikli As String        'SKUPUWAGI      Memo
    Public oSprzedOpiekunkli As String    'SPRZEDOPIEKUN    Text, 10   uzytkownik

    Public oSprzedOKontaktRKli As String  'SPRZEDOKONTAKTR  Text, 4    rok
    Public oSprzedOKontaktTKli As String  'SPRZEDOKONTAKTT  Integer    nr tygodnia
    Public oSprzedOKontaktDKli As String  'SPRZEDOKONTAKTD  Text, 12   dzien tygodnia
    Public oSprzedNKontaktRKli As String  'SPRZEDNKONTAKTR  Text, 4    rok
    Public oSprzedNKontaktTKli As String  'SPRZEDNKONTAKTT  Integer    nr tygodnia
    Public oSprzedNKontaktDKli As String  'SKUPNKONTAKTT    Text, 12    nr tygodnia

    Public oSprzedUwagiKli As String      'SPRZEDUWAGI      Memo
    Public oWersjaKli As Integer          'WERSJA           Integer
    Public oMarkerKli As Boolean          'MARKER           boolean
    Public oMarkerPKli As Boolean         'MARKERP          boolean
    Public oAktywnyKli As Boolean         'AKTYWNY          boolean

    Public oZakupPopRokKli As Double       'ZAKUPPOPROK        Double
    Public oUslugiDodKli As String         'USLUGIDOD          Text, 200
    Public oRZURap01 As String              'RZURAP01     Text, 100
    Public oRZURap02 As String              'RZURAP02     Text, 100
    Public oRZURap03 As String              'RZURAP03     Text, 100
    Public oRZURap04 As String              'RZURAP04     Text, 100
    Public oRZURap05 As String              'RZURAP05     Text, 100
    Public oRZURap06 As String              'RZURAP06     Text, 100
    Public oRZURap07 As String              'RZURAP07     Text, 100
    Public oRZURap08 As String              'RZURAP08     Text, 100
    Public oRZURap09 As String              'RZURAP09     Text, 100











    'zmienne do niezależnej identyfikacji
    Public o2IdentKli As String         'IDENTKLIENTA   Text, 6
    Public o2NrTOFITxtKli As String     'NRTOFITXT      Text  500
    Public o2Nazwa1Kli As String        'NAZWA1         Text, 100
    Public o2KodPoczKli As String       'KODPOCZTOWY    Text, 10
    Public o2MiejscKli As String        'MIEJSCOWOSC    Text, 50
    Public o2UlicaKli As String         'ULICA          Text, 50



    Public sConnectionKlienci As String = ConnectionStrings()
    Public HQConnectionKlienci As String = HQConnectionStrings()



    Public sSelectSQLKlienci As String = "SELECT " &
                                           "IDENTKLIENTA AS NRKARTY, " &
                                           "NIP," &
                                           "NRTOFITXT," &
                                           "KARTAPAYBACK," &
                                           "NRYKARTPAYBACK," &
                                           "NAZWA1 AS NAZWAKLIENTA," &
                                           "MIEJSCOWOSC," &
                                           "KODPOCZTOWY," &
                                           "ULICA," &
                                           "NUMNRDOMU," &
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
                                     "RZURAP09, " &
                                           "WERSJA " &
                                        "FROM Klienci "

    Public sSelectSQLKlienciAktywni As String = sSelectSQLKlienci & " WHERE (AKTYWNY='TRUE') "



    Public sDeleteSQLKlienci As String = "DELETE FROM Klienci " &
                                        "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLKlienci As String = "UPDATE Klienci SET " &
                                                 "NIP=@oNIPKli," &
                                                 "NRTOFITXT=@oNrTofiTxtKli," &
                                                 "KARTAPAYBACK=@oKARTAPAYBACKKli," &
                                                 "NRYKARTPAYBACK=@oNRYKARTPAYBACKKli," &
                                                 "NAZWA1=@oNazwa1Kli," &
                                                 "MIEJSCOWOSC=@oMiejscKli," &
                                                 "KODPOCZTOWY=@oKodPoczKli," &
                                                 "ULICA=@oUlicaKli," &
                                                 "NUMNRDOMU=@oNumNrDomuKli," &
                                                 "NRDOMU=@oNrDomuKli," &
                                                 "IDDOMU=@oIDDomuKli," &
                                                 "REJKLIENTA=@oRejonKli," &
                                                 "TELEFON=@oTelefonKli," &
                                                 "FAX=@oFaxKli," &
                                                 "EMAIL=@oeMailKli," &
                                               "BRANZA=@oBRANZAKli," &
                                               "PODBRANZA=@oPODBRANZAKli," &
                                               "LICZBAZATRUDNIONYCH=@oLICZBAZATRUDNIONYCHKli," &
                                               "LICZBAPRACZDALNIE=@oLICZBAPRACZDALNIEKli," &
                                       "WYKAZURZADZEN=@oWYKAZURZADZENKli," &
                                         "SPOSOBWYBORUDOSTAWCY=@oSPOSOBWYBORUDOSTAWCYKli," &
                                         "ZAINSTALOWANOMONITOR=@oZAINSTALOWANOMONITORKli," &
                                         "LICZBAURZADZEN=@oLICZBAURZADZENKli," &
                                         "LICZBAWYDRUKOW=@oLICZBAWYDRUKOWKli," &
                                         "POTENCJALDRUKU=@oPOTENCJALDRUKUKli," &
                                       "AKTZAKUPMATEKSP=@oAKTZAKUPMATEKSPKli," &
                                       "AKTROZLICZASTRONYDRUKU=@oAKTROZLICZASTRONYDRUKUKli," &
                                       "AKTKORZYSTAZNAJMU=@oAKTKORZYSTAZNAJMUKli," &
                                       "ZAINTROZLICZANIEMSTRONDRUKU=@oZAINTROZLICZANIEMSTRONDRUKUKli," &
                                       "ZAINTKORZYSTANIEMZNAJMU=@oZAINTKORZYSTANIEMZNAJMUKli," &
                                         "DATAWERYFSEGMENTACJI=@oDATAWERYFSEGMENTACJIKli," &
                                         "PLANDLUGOOKRESOWY=@oPLANDLUGOOKRESOWYKli," &
                                         "PLANKROTKOOKRESOWY=@oPLANKROTKOOKRESOWYKli," &
                                                     "RODZSPRZETUL=@oRodzSprzetuLKli," &
                                                     "RODZSPRZETUA=@oRodzSprzetuAKli," &
                                                     "RODZSPRZETUT=@oRodzSprzetuTKli," &
                                                     "OFERTATZLOZENIA=@oOfertaTZLOZENIAKli," &
                                                     "OFERTAPLIK=@oOfertaPLIKKli," &
                                                     "PKONTAKT=@oPKontaktKli," &
                                                     "SKUPUWAGI=@oSkupUwagiKli," &
                                                     "SPRZEDOPIEKUN=@oSprzedOpiekunKli," &
                                                     "SPRZEDOKONTAKTR=@oSprzedOKontaktRKli," &
                                                     "SPRZEDOKONTAKTT=@oSprzedOKontaktTKli," &
                                                     "SPRZEDOKONTAKTD=@oSprzedOKontaktDKli," &
                                                     "SPRZEDNKONTAKTR=@oSprzedNKontaktRKli," &
                                                     "SPRZEDNKONTAKTT=@oSprzedNKontaktTKli," &
                                                     "SPRZEDNKONTAKTD=@oSprzedNKontaktDKli," &
                                                     "SPRZEDUWAGI=@oSprzedUwagiKli," &
                                                     "KTOWPISAL=@oWpisalKli," &
                                                     "UWAGI=@oUwagiKli," &
                                                     "MARKER=@oMarkerKli, " &
                                                     "MARKERP=@oMarkerPKli, " &
                                                     "AKTYWNY=@oAKTYWNYKli, " &
                                           "ZAKUPPOPROK=@oZAKUPPOPROKKli, " &
                                           "USLUGIDOD=@oUSLUGIDODKli, " &
                                           "RZURAP01=@oRZURAP01, " &
                                           "RZURAP02=@oRZURAP02, " &
                                           "RZURAP03=@oRZURAP03, " &
                                           "RZURAP04=@oRZURAP04, " &
                                           "RZURAP05=@oRZURAP05, " &
                                           "RZURAP06=@oRZURAP06, " &
                                           "RZURAP07=@oRZURAP07, " &
                                           "RZURAP08=@oRZURAP08, " &
                                           "RZURAP09=@oRZURAP09, " &
                                                "WERSJA=@oWersjaKli " &
                                        "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                              "(WERSJA=@orygWersja)"



    Public sInsertSQLKlienci As String = "INSERT INTO Klienci " &
                                         "(IDENTKLIENTA, " &
                                           "NIP," &
                                           "NRTOFITXT," &
                                           "KARTAPAYBACK," &
                                           "NRYKARTPAYBACK," &
                                           "NAZWA1," &
                                           "MIEJSCOWOSC," &
                                           "KODPOCZTOWY," &
                                           "ULICA," &
                                           "NUMNRDOMU," &
                                           "NRDOMU," &
                                           "IDDOMU," &
                                           "REJKLIENTA," &
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
                                               "SKUPUWAGI," &
                                               "SPRZEDOPIEKUN," &
                                               "SPRZEDOKONTAKTR," &
                                               "SPRZEDOKONTAKTT," &
                                               "SPRZEDOKONTAKTD," &
                                               "SPRZEDNKONTAKTR," &
                                               "SPRZEDNKONTAKTT," &
                                               "SPRZEDNKONTAKTD," &
                                               "SPRZEDUWAGI," &
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
                                           "RZURAP09, " &
                                               "WERSJA " &
                                           ") " &
                                "VALUES  (@oIdentKli," &
                                             "@oNIPKli," &
                                             "@oNrTofiTxtKli," &
                                             "@oKARTAPAYBACKKli," &
                                             "@oNRYKARTPAYBACKKli," &
                                             "@oNazwa1Kli," &
                                             "@oMiejscKli," &
                                             "@oKodPoczKli," &
                                             "@oUlicaKli," &
                                             "@oNumNrDomuKli," &
                                             "@oNrDomuKli," &
                                             "@oIDDomuKli," &
                                             "@oRejonKli," &
                                             "@oTelefonKli," &
                                             "@oFaxKli," &
                                             "@oeMailKli," &
                                               "@oBRANZAKli," &
                                               "@oPODBRANZAKli," &
                                               "@oLICZBAZATRUDNIONYCHKli," &
                                               "@oLICZBAPRACZDALNIEKli," &
                                       "@oWYKAZURZADZENKli," &
                                         "@oSPOSOBWYBORUDOSTAWCYKli," &
                                         "@oZAINSTALOWANOMONITORKli," &
                                         "@oLICZBAURZADZENKli," &
                                         "@oLICZBAWYDRUKOWKli," &
                                         "@oPOTENCJALDRUKUKli," &
                                       "@oAKTZAKUPMATEKSPKli," &
                                       "@oAKTROZLICZASTRONYDRUKUKli," &
                                       "@oAKTKORZYSTAZNAJMUKli," &
                                       "@oZAINTROZLICZANIEMSTRONDRUKUKli," &
                                       "@oZAINTKORZYSTANIEMZNAJMUKli," &
                                         "@oDATAWERYFSEGMENTACJIKli," &
                                         "@oPLANDLUGOOKRESOWYKli," &
                                         "@oPLANKROTKOOKRESOWYKli," &
                                             "@oRodzSprzetuLKli," &
                                             "@oRodzSprzetuAKli," &
                                             "@oRodzSprzetuTKli," &
                                             "@oOfertaTZLOZENIAKli," &
                                             "@oOfertaPLIKKli," &
                                             "@oPKontaktKli," &
                                             "@oSkupUwagiKli," &
                                             "@oSprzedOpiekunKli," &
                                             "@oSprzedOKontaktRKli," &
                                             "@oSprzedOKontaktTKli," &
                                             "@oSprzedOKontaktDKli," &
                                             "@oSprzedNKontaktRKli," &
                                             "@oSprzedNKontaktTKli," &
                                             "@oSprzedNKontaktDKli," &
                                             "@oSprzedUwagiKli," &
                                             "@oWpisalKli," &
                                             "@oUwagiKli," &
                                             "@oMarkerKli, " &
                                             "@oMarkerPKli, " &
                                             "@oAKTYWNYKli, " &
                                           "@oZAKUPPOPROKKli, " &
                                           "@oUSLUGIDODKli, " &
                                           "@oRZURAP01, " &
                                           "@oRZURAP02, " &
                                           "@oRZURAP03, " &
                                           "@oRZURAP04, " &
                                           "@oRZURAP05, " &
                                           "@oRZURAP06, " &
                                           "@oRZURAP07, " &
                                           "@oRZURAP08, " &
                                           "@oRZURAP09, " &
                                             "@oWersjaKli " &
                                         ")"




    Public Function NowePolaTabeliKlienci(ByVal TlumaczonaNazwa As String) As String
        'tlumaczy ORYGINALNE na NOWE
        Dim NazwaPola As String = TlumaczonaNazwa
        Select Case UCase(TlumaczonaNazwa)
            Case "IDENTKLIENTA"
                NazwaPola = "NRKARTY"
            Case "NAZWA1"
                NazwaPola = "NAZWAKLIENTA"
            Case "REJKLIENTA"
                NazwaPola = "REJONKLIENTA"

            Case "SKUPUWAGI"
                NazwaPola = "PROMOTORUWAGI"

            Case "SPRZEDOPIEKUN"
                NazwaPola = "OPIEKUN"
            Case "SPRZEDOKONTAKTR"
                NazwaPola = "OPIEKUNOSTKONTAKTROK"
            Case "SPRZEDOKONTAKTT"
                NazwaPola = "OPIEKUNOSTKONTAKTTYDZ"
            Case "SPRZEDOKONTAKTD"
                NazwaPola = "OPIEKUNOSTKONTAKTDZIEN"
            Case "SPRZEDNKONTAKTR"
                NazwaPola = "OPIEKUNKOLKONTAKTROK"
            Case "SPRZEDNKONTAKTT"
                NazwaPola = "OPIEKUNKOLKONTAKTTYDZ"
            Case "SPRZEDNKONTAKTD"
                NazwaPola = "OPIEKUNKOLKONTAKTDZIEN"

            Case "SPRZEDUWAGI"
                NazwaPola = "OPIEKUNUWAGI"
        End Select
        Return (NazwaPola)
    End Function


    Public Function OryginalnePolaTabeliKlienci(ByVal TlumaczonaNazwa As String) As String
        'tlumaczy NOWE na ORYGINALNE
        Dim NazwaPola As String = TlumaczonaNazwa
        Select Case UCase(TlumaczonaNazwa)
            Case "NRKARTY"
                NazwaPola = "IDENTKLIENTA"
            Case "NAZWAKLIENTA"
                NazwaPola = "NAZWA1"
            Case "REJONKLIENTA"
                NazwaPola = "REJKLIENTA"

            Case "PROMOTORUWAGI"
                NazwaPola = "SKUPUWAGI"
            Case "OPIEKUN"
                NazwaPola = "SPRZEDOPIEKUN"
            Case "OPIEKUNOSTKONTAKTROK"
                NazwaPola = "SPRZEDOKONTAKTR"
            Case "OPIEKUNOSTKONTAKTTYDZ"
                NazwaPola = "SPRZEDOKONTAKTT"
            Case "OPIEKUNOSTKONTAKTDZIEN"
                NazwaPola = "SPRZEDOKONTAKTD"
            Case "OPIEKUNKOLKONTAKTROK"
                NazwaPola = "SPRZEDNKONTAKTR"
            Case "OPIEKUNKOLKONTAKTTYDZ"
                NazwaPola = "SPRZEDNKONTAKTT"
            Case "OPIEKUNKOLKONTAKTDZIEN"
                NazwaPola = "SPRZEDNKONTAKTD"

            Case "OPIEKUNUWAGI"
                NazwaPola = "SPRZEDUWAGI"
        End Select
        Return (NazwaPola)
    End Function



    'SQLDeleteKlienci(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateKlienci(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertKlienci(sqlCommandInsert, sqlDataAdapter)


    Public Sub SQLDeleteKlienci(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParamDel As SqlClient.SqlParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        sqlParamDel = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 10, "NRKARTY")
        sqlParamDel.SourceVersion = DataRowVersion.Original
        sqlParamDel = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParamDel.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.DeleteCommand = sqlCommandDelete

    End Sub

    Public Sub SQLUpdateKlienci(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParamUpd As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda UPDATE
        'parametry aktualizacji
        'sqlcommandUpdate.Parameters.Add("@oIdentKli", sqldbtype.varchar, 6, "NRKARTY")
        sqlCommandUpdate.Parameters.Add("@oNIPKli", SqlDbType.VarChar, 15, "NIP")
        sqlCommandUpdate.Parameters.Add("@oNrTofiTxtKli", SqlDbType.VarChar, 500, "NRTOFITXT")
        sqlCommandUpdate.Parameters.Add("@oKartaPaybackKli", SqlDbType.Bit, Nothing, "KARTAPAYBACK")
        sqlCommandUpdate.Parameters.Add("@oNryKartPaybackKli", SqlDbType.Text, Nothing, "NRYKARTPAYBACK")

        sqlCommandUpdate.Parameters.Add("@oNazwa1Kli", SqlDbType.VarChar, 100, "NAZWAKLIENTA")
        sqlCommandUpdate.Parameters.Add("@oMiejscKli", SqlDbType.VarChar, 50, "MIEJSCOWOSC")
        sqlCommandUpdate.Parameters.Add("@oKodPoczKli", SqlDbType.VarChar, 10, "KODPOCZTOWY")
        sqlCommandUpdate.Parameters.Add("@oUlicaKli", SqlDbType.VarChar, 50, "ULICA")
        sqlCommandUpdate.Parameters.Add("@oNumNrDomuKli", SqlDbType.Int, Nothing, "NUMNRDOMU")
        sqlCommandUpdate.Parameters.Add("@oNrDomuKli", SqlDbType.VarChar, 10, "NRDOMU")
        sqlCommandUpdate.Parameters.Add("@oIDDomuKli", SqlDbType.VarChar, 10, "IDDOMU")
        sqlCommandUpdate.Parameters.Add("@oRejonKli", SqlDbType.VarChar, 20, "REJONKLIENTA")
        sqlCommandUpdate.Parameters.Add("@oTelefonKli", SqlDbType.VarChar, 50, "TELEFON")
        sqlCommandUpdate.Parameters.Add("@oFaxKli", SqlDbType.VarChar, 50, "FAX")
        sqlCommandUpdate.Parameters.Add("@oeMailKli", SqlDbType.VarChar, 50, "EMAIL")

        sqlCommandUpdate.Parameters.Add("@oBRANZAKli", SqlDbType.VarChar, 100, "BRANZA")
        sqlCommandUpdate.Parameters.Add("@oPODBRANZAKli", SqlDbType.VarChar, 100, "PODBRANZA")
        sqlCommandUpdate.Parameters.Add("@oLICZBAZATRUDNIONYCHKli", SqlDbType.Int, Nothing, "LICZBAZATRUDNIONYCH")
        sqlCommandUpdate.Parameters.Add("@oLICZBAPRACZDALNIEKli", SqlDbType.Int, Nothing, "LICZBAPRACZDALNIE")


        ''..............................
        'Public oWykazUrzadzenKli As String              'WYKLAZURZADZEN    Memo

        'Public oSposobWyboruDostawcyKli As String       'SPOSOBWYBORUDOSTAWCY    Text 30
        'Public oZainstalowanoMonitorKli As Boolean      'ZAINSTALOWANOMONITOR    Logical
        'Public oLiczbaUrzadzenKli As Integer            'LICZBAURZADZEN     Integer
        'Public oLiczbaWydrukowKli As Integer            'LICZBAWYDRUKOW     Integer
        'Public oPotencjalDrukuKli As String             'POTENCJALDRUKU     Text 2

        'Public oAktZakupMatEkspKli As Boolean               'AKTZAKUPMATEKSP    Logical
        'Public oAktRozliczaStronyDrukuKli As Boolean        'AKTROZLICZASTRONYDRUKU    Logical
        'Public oAktKorzystaZNajmuKli As Boolean             'AKTKORZYSTAZNAJMU  Logical
        'Public oZaintRozliczaniemStronDrukuKli As Boolean   'ZAINTROZLICZANIEMSTRONDRUKU    Logical
        'Public oZaintKorzystaniemZNajmuKli As Boolean       'ZAINTKORZYSTANIEMZNAJMU    Logical

        'Public oDataWeryfSegmentacji As String          'DATAWERYFSEGMENTACJI  Text 30

        'Public oPlanDlugookresowyKli As Integer         'PLANDLUGOOKRESOWY     Integer
        'Public oPlanKrotkookresowyKli As Integer        'PLANKROTKOOKRESOWY    Integer
        ''......................................
        sqlCommandUpdate.Parameters.Add("@oWYKAZURZADZENKli", SqlDbType.Text, Nothing, "WYKAZURZADZEN")
        sqlCommandUpdate.Parameters.Add("@oSPOSOBWYBORUDOSTAWCYKli", SqlDbType.VarChar, 30, "SPOSOBWYBORUDOSTAWCY")
        sqlCommandUpdate.Parameters.Add("@oZAINSTALOWANOMONITORKli", SqlDbType.Bit, Nothing, "ZAINSTALOWANOMONITOR")
        sqlCommandUpdate.Parameters.Add("@oLICZBAURZADZENKli", SqlDbType.Int, Nothing, "LICZBAURZADZEN")
        sqlCommandUpdate.Parameters.Add("@oLICZBAWYDRUKOWKli", SqlDbType.Int, Nothing, "LICZBAWYDRUKOW")
        sqlCommandUpdate.Parameters.Add("@oPOTENCJALDRUKUKli", SqlDbType.VarChar, 2, "POTENCJALDRUKU")

        sqlCommandUpdate.Parameters.Add("@oAKTZAKUPMATEKSPKli", SqlDbType.Bit, Nothing, "AKTZAKUPMATEKSP")
        sqlCommandUpdate.Parameters.Add("@oAKTROZLICZASTRONYDRUKUKli", SqlDbType.Bit, Nothing, "AKTROZLICZASTRONYDRUKU")
        sqlCommandUpdate.Parameters.Add("@oAKTKORZYSTAZNAJMUKli", SqlDbType.Bit, Nothing, "AKTKORZYSTAZNAJMU")
        sqlCommandUpdate.Parameters.Add("@oZAINTROZLICZANIEMSTRONDRUKUKli", SqlDbType.Bit, Nothing, "ZAINTROZLICZANIEMSTRONDRUKU")
        sqlCommandUpdate.Parameters.Add("@oZAINTKORZYSTANIEMZNAJMUKli", SqlDbType.Bit, Nothing, "ZAINTKORZYSTANIEMZNAJMU")

        sqlCommandUpdate.Parameters.Add("@oDATAWERYFSEGMENTACJIKli", SqlDbType.VarChar, 10, "DATAWERYFSEGMENTACJI")

        sqlCommandUpdate.Parameters.Add("@oPLANDLUGOOKRESOWYKli", SqlDbType.Int, Nothing, "PLANDLUGOOKRESOWY")
        sqlCommandUpdate.Parameters.Add("@oPLANKROTKOOKRESOWYKli", SqlDbType.Int, Nothing, "PLANKROTKOOKRESOWY")
        ''......................................
        sqlCommandUpdate.Parameters.Add("@oRodzSprzetuLKli", SqlDbType.Bit, Nothing, "RODZSPRZETUL")
        sqlCommandUpdate.Parameters.Add("@oRodzSprzetuAKli", SqlDbType.Bit, Nothing, "RODZSPRZETUA")
        sqlCommandUpdate.Parameters.Add("@oRodzSprzetuTKli", SqlDbType.Bit, Nothing, "RODZSPRZETUT")

        sqlCommandUpdate.Parameters.Add("@oOfertaTZLOZENIAKli", SqlDbType.VarChar, 4, "OFERTATZLOZENIA")
        sqlCommandUpdate.Parameters.Add("@oOfertaPLIKKli", SqlDbType.VarChar, 100, "OFERTAPLIK")
        sqlCommandUpdate.Parameters.Add("@oPKontaktKli", SqlDbType.VarChar, 10, "PKONTAKT")

        sqlCommandUpdate.Parameters.Add("@oSkupUwagiKli", SqlDbType.Text, Nothing, "PROMOTORUWAGI")
        sqlCommandUpdate.Parameters.Add("@oSprzedOpiekunKli", SqlDbType.VarChar, 10, "OPIEKUN")
        sqlCommandUpdate.Parameters.Add("@oSprzedOKontaktRKli", SqlDbType.VarChar, 4, "OPIEKUNOSTKONTAKTROK")
        sqlCommandUpdate.Parameters.Add("@oSprzedOKontaktTKli", SqlDbType.Int, Nothing, "OPIEKUNOSTKONTAKTTYDZ")
        sqlCommandUpdate.Parameters.Add("@oSprzedOKontaktDKli", SqlDbType.VarChar, 12, "OPIEKUNOSTKONTAKTDZIEN")

        sqlCommandUpdate.Parameters.Add("@oSprzedNKontaktRKli", SqlDbType.VarChar, 4, "OPIEKUNKOLKONTAKTROK")
        sqlCommandUpdate.Parameters.Add("@oSprzedNKontaktTKli", SqlDbType.Int, Nothing, "OPIEKUNKOLKONTAKTTYDZ")
        sqlCommandUpdate.Parameters.Add("@oSprzedNKontaktDKli", SqlDbType.VarChar, 12, "OPIEKUNKOLKONTAKTDZIEN")
        sqlCommandUpdate.Parameters.Add("@oSprzedUwagiKli", SqlDbType.Text, Nothing, "OPIEKUNUWAGI")
        sqlCommandUpdate.Parameters.Add("@oWpisalKli", SqlDbType.VarChar, 10, "KTOWPISAL")
        sqlCommandUpdate.Parameters.Add("@oUwagiKli", SqlDbType.Text, Nothing, "UWAGI")

        sqlCommandUpdate.Parameters.Add("@oMarkerKli", SqlDbType.Bit, Nothing, "MARKER")
        sqlCommandUpdate.Parameters.Add("@oMarkerPKli", SqlDbType.Bit, Nothing, "MARKERP")
        sqlCommandUpdate.Parameters.Add("@oAKTYWNYKli", SqlDbType.Bit, Nothing, "AKTYWNY")
        sqlCommandUpdate.Parameters.Add("@oZAKUPPOPROKKli", SqlDbType.Float, Nothing, "ZAKUPPOPROK")
        sqlCommandUpdate.Parameters.Add("@oUSLUGIDODKli", SqlDbType.VarChar, 200, "USLUGIDOD")
        sqlCommandUpdate.Parameters.Add("@oRZURAP01", SqlDbType.VarChar, 100, "RZURAP01")
        sqlCommandUpdate.Parameters.Add("@oRZURAP02", SqlDbType.VarChar, 100, "RZURAP02")
        sqlCommandUpdate.Parameters.Add("@oRZURAP03", SqlDbType.VarChar, 100, "RZURAP03")
        sqlCommandUpdate.Parameters.Add("@oRZURAP04", SqlDbType.VarChar, 100, "RZURAP04")
        sqlCommandUpdate.Parameters.Add("@oRZURAP05", SqlDbType.VarChar, 100, "RZURAP05")
        sqlCommandUpdate.Parameters.Add("@oRZURAP06", SqlDbType.VarChar, 100, "RZURAP06")
        sqlCommandUpdate.Parameters.Add("@oRZURAP07", SqlDbType.VarChar, 100, "RZURAP07")
        sqlCommandUpdate.Parameters.Add("@oRZURAP08", SqlDbType.VarChar, 100, "RZURAP08")
        sqlCommandUpdate.Parameters.Add("@oRZURAP09", SqlDbType.VarChar, 100, "RZURAP09")

        sqlCommandUpdate.Parameters.Add("@oWersjaKli", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParamUpd = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.VarChar, 10, "NRKARTY")
        sqlParamUpd.SourceVersion = DataRowVersion.Original
        sqlParamUpd = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParamUpd.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.UpdateCommand = sqlCommandUpdate
    End Sub


    Public Sub SQLInsertKlienci(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParamIns As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda INSERT
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oIdentKli", SqlDbType.VarChar, 6, "NRKARTY")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oNIPKli", SqlDbType.VarChar, 15, "NIP")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oNrTofiTxtKli", SqlDbType.VarChar, 500, "NRTOFITXT")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oKartaPaybackKli", SqlDbType.Bit, Nothing, "KARTAPAYBACK")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oNryKartPaybackKli", SqlDbType.Text, Nothing, "NRYKARTPAYBACK")
        sqlParamIns.SourceVersion = DataRowVersion.Current

        sqlParamIns = sqlCommandInsert.Parameters.Add("@oNazwa1Kli", SqlDbType.VarChar, 100, "NAZWAKLIENTA")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oMiejscKli", SqlDbType.VarChar, 50, "MIEJSCOWOSC")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oKodPoczKli", SqlDbType.VarChar, 10, "KODPOCZTOWY")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oUlicaKli", SqlDbType.VarChar, 50, "ULICA")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oNumNrDomuKli", SqlDbType.Int, Nothing, "NUMNRDOMU")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oNrDomuKli", SqlDbType.VarChar, 10, "NRDOMU")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oIDDomuKli", SqlDbType.VarChar, 10, "IDDOMU")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oRejonKli", SqlDbType.VarChar, 20, "REJONKLIENTA")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oTelefonKli", SqlDbType.VarChar, 50, "TELEFON")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oFaxKli", SqlDbType.VarChar, 50, "FAX")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oeMailKli", SqlDbType.VarChar, 50, "EMAIL")
        sqlParamIns.SourceVersion = DataRowVersion.Current

        sqlParamIns = sqlCommandInsert.Parameters.Add("@oBRANZAKli", SqlDbType.VarChar, 100, "BRANZA")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oPODBRANZAKli", SqlDbType.VarChar, 100, "PODBRANZA")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oLICZBAZATRUDNIONYCHKli", SqlDbType.Int, Nothing, "LICZBAZATRUDNIONYCH")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oLICZBAPRACZDALNIEKli", SqlDbType.Int, Nothing, "LICZBAPRACZDALNIE")
        sqlParamIns.SourceVersion = DataRowVersion.Current

        ''..............................
        'Public oWykazUrzadzenKli As String              'WYKLAZURZADZEN    Memo

        'Public oSposobWyboruDostawcyKli As String       'SPOSOBWYBORUDOSTAWCY    Text 30
        'Public oZainstalowanoMonitorKli As Boolean      'ZAINSTALOWANOMONITOR    Logical
        'Public oLiczbaUrzadzenKli As Integer            'LICZBAURZADZEN     Integer
        'Public oLiczbaWydrukowKli As Integer            'LICZBAWYDRUKOW     Integer
        'Public oPotencjalDrukuKli As String             'POTENCJALDRUKU     Text 2

        'Public oAktZakupMatEkspKli As Boolean               'AKTZAKUPMATEKSP    Logical
        'Public oAktRozliczaStronyDrukuKli As Boolean        'AKTROZLICZASTRONYDRUKU    Logical
        'Public oAktKorzystaZNajmuKli As Boolean             'AKTKORZYSTAZNAJMU  Logical
        'Public oZaintRozliczaniemStronDrukuKli As Boolean   'ZAINTROZLICZANIEMSTRONDRUKU    Logical
        'Public oZaintKorzystaniemZNajmuKli As Boolean       'ZAINTKORZYSTANIEMZNAJMU    Logical

        'Public oDataWeryfSegmentacji As String          'DATAWERYFSEGMENTACJI  Text 10

        'Public oPlanDlugookresowyKli As Integer         'PLANDLUGOOKRESOWY     Integer
        'Public oPlanKrotkookresowyKli As Integer        'PLANKROTKOOKRESOWY    Integer
        ''......................................
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oWYKAZURZADZENKli", SqlDbType.Text, Nothing, "WYKAZURZADZEN")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oSPOSOBWYBORUDOSTAWCYKli", SqlDbType.VarChar, 30, "SPOSOBWYBORUDOSTAWCY")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oZAINSTALOWANOMONITORKli", SqlDbType.Bit, Nothing, "ZAINSTALOWANOMONITOR")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oLICZBAURZADZENKli", SqlDbType.Int, Nothing, "LICZBAURZADZEN")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oLICZBAWYDRUKOWKli", SqlDbType.Int, Nothing, "LICZBAWYDRUKOW")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oPOTENCJALDRUKUKli", SqlDbType.VarChar, 2, "POTENCJALDRUKU")
        sqlParamIns.SourceVersion = DataRowVersion.Current

        sqlParamIns = sqlCommandInsert.Parameters.Add("@oAKTZAKUPMATEKSPKli", SqlDbType.Bit, Nothing, "AKTZAKUPMATEKSP")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oAKTROZLICZASTRONYDRUKUKli", SqlDbType.Bit, Nothing, "AKTROZLICZASTRONYDRUKU")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oAKTKORZYSTAZNAJMUKli", SqlDbType.Bit, Nothing, "AKTKORZYSTAZNAJMU")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oZAINTROZLICZANIEMSTRONDRUKUKli", SqlDbType.Bit, Nothing, "ZAINTROZLICZANIEMSTRONDRUKU")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oZAINTKORZYSTANIEMZNAJMUKli", SqlDbType.Bit, Nothing, "ZAINTKORZYSTANIEMZNAJMU")
        sqlParamIns.SourceVersion = DataRowVersion.Current

        sqlParamIns = sqlCommandInsert.Parameters.Add("@oDATAWERYFSEGMENTACJIKli", SqlDbType.VarChar, 10, "DATAWERYFSEGMENTACJI")
        sqlParamIns.SourceVersion = DataRowVersion.Current

        sqlParamIns = sqlCommandInsert.Parameters.Add("@oPLANDLUGOOKRESOWYKli", SqlDbType.Int, Nothing, "PLANDLUGOOKRESOWY")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oPLANKROTKOOKRESOWYKli", SqlDbType.Int, Nothing, "PLANKROTKOOKRESOWY")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        ''......................................


        sqlParamIns = sqlCommandInsert.Parameters.Add("@oRodzSprzetuLKli", SqlDbType.Bit, Nothing, "RODZSPRZETUL")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oRodzSprzetuAKli", SqlDbType.Bit, Nothing, "RODZSPRZETUA")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oRodzSprzetuTKli", SqlDbType.Bit, Nothing, "RODZSPRZETUT")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oOfertaTZLOZENIAKli", SqlDbType.VarChar, 4, "OFERTATZLOZENIA")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oOfertaPLIKKli", SqlDbType.VarChar, 100, "OFERTAPLIK")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oPKontaktKli", SqlDbType.VarChar, 10, "PKONTAKT")
        sqlParamIns.SourceVersion = DataRowVersion.Current

        sqlParamIns = sqlCommandInsert.Parameters.Add("@oSkupUwagiKli", SqlDbType.Text, Nothing, "PROMOTORUWAGI")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oSprzedOpiekunKli", SqlDbType.VarChar, 10, "OPIEKUN")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oSprzedOKontaktRKli", SqlDbType.VarChar, 4, "OPIEKUNOSTKONTAKTROK")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oSprzedOKontaktTKli", SqlDbType.Int, Nothing, "OPIEKUNOSTKONTAKTTYDZ")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oSprzedOKontaktDKli", SqlDbType.VarChar, 12, "OPIEKUNOSTKONTAKTDZIEN")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oSprzedNKontaktRKli", SqlDbType.VarChar, 4, "OPIEKUNKOLKONTAKTROK")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oSprzedNKontaktTKli", SqlDbType.Int, Nothing, "OPIEKUNKOLKONTAKTTYDZ")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oSprzedNKontaktDKli", SqlDbType.VarChar, 12, "OPIEKUNKOLKONTAKTDZIEN")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oSprzedUwagiKli", SqlDbType.Text, Nothing, "OPIEKUNUWAGI")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oWpisalKli", SqlDbType.VarChar, 10, "KTOWPISAL")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oUwagiKli", SqlDbType.Text, Nothing, "UWAGI")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oMarkerKli", SqlDbType.Bit, Nothing, "MARKER")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oMarkerPKli", SqlDbType.Bit, Nothing, "MARKERP")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oAKTYWNYKli", SqlDbType.Bit, Nothing, "AKTYWNY")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oZAKUPPOPROKKli", SqlDbType.Float, Nothing, "ZAKUPPOPROK")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oUSLUGIDODKli", SqlDbType.VarChar, 200, "USLUGIDOD")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oRZURAP01", SqlDbType.VarChar, 100, "RZURAP01")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oRZURAP02", SqlDbType.VarChar, 100, "RZURAP02")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oRZURAP03", SqlDbType.VarChar, 100, "RZURAP03")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oRZURAP04", SqlDbType.VarChar, 100, "RZURAP04")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oRZURAP05", SqlDbType.VarChar, 100, "RZURAP05")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oRZURAP06", SqlDbType.VarChar, 100, "RZURAP06")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oRZURAP07", SqlDbType.VarChar, 100, "RZURAP07")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oRZURAP08", SqlDbType.VarChar, 100, "RZURAP08")
        sqlParamIns.SourceVersion = DataRowVersion.Current
        sqlParamIns = sqlCommandInsert.Parameters.Add("@oRZURAP09", SqlDbType.VarChar, 100, "RZURAP09")
        sqlParamIns.SourceVersion = DataRowVersion.Current

        sqlParamIns = sqlCommandInsert.Parameters.Add("@oWersjaKli", SqlDbType.Int, Nothing, "WERSJA")
        sqlParamIns.SourceVersion = DataRowVersion.Current

        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub
































    'Public Sub DBDeleteKlienci(ByRef dbCommandDelete As OleDb.OleDbCommand,
    '                       ByRef dbDataAdapter As OleDb.OleDbDataAdapter)

    '    Dim dbParamDel As OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda DELETE
    '    'parametry filtrowania
    '    dbParamDel = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 6, "NRKARTY")
    '    dbParamDel.SourceVersion = DataRowVersion.Original
    '    dbParamDel = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParamDel.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.DeleteCommand = dbCommandDelete
    'End Sub

    'Public Sub DBUpdateKlienci(ByRef dbCommandUpdate As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParamUpd As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda UPDATE
    '    'parametry aktualizacji
    '    'dbCommandUpdate.Parameters.Add("@oIdentKli", OleDb.OleDbType.varchar, 6, "NRKARTY")
    '    dbCommandUpdate.Parameters.Add("@oNIPKli", OleDb.OleDbType.VarChar, 15, "NIP")
    '    dbCommandUpdate.Parameters.Add("@oNrTofiTxtKli", OleDb.OleDbType.VarChar, 500, "NRTOFITXT")
    '    dbCommandUpdate.Parameters.Add("@oKartaPaybackKli", OleDb.OleDbType.Boolean, Nothing, "KARTAPAYBACK")
    '    dbCommandUpdate.Parameters.Add("@oNryKartPaybackKli", OleDb.OleDbType.WChar, Nothing, "NRYKARTPAYBACK")

    '    dbCommandUpdate.Parameters.Add("@oWYKAZURZADZENKli", OleDb.OleDbType.WChar, Nothing, "WYKAZURZADZEN")

    '    dbCommandUpdate.Parameters.Add("@oNAZWA1Kli", OleDb.OleDbType.VarChar, 100, "NAZWAKLIENTA")
    '    dbCommandUpdate.Parameters.Add("@oMiejscKli", OleDb.OleDbType.VarChar, 50, "MIEJSCOWOSC")
    '    dbCommandUpdate.Parameters.Add("@oKodPoczKli", OleDb.OleDbType.VarChar, 10, "KODPOCZTOWY")
    '    dbCommandUpdate.Parameters.Add("@oUlicaKli", OleDb.OleDbType.VarChar, 50, "ULICA")
    '    dbCommandUpdate.Parameters.Add("@oNumNrDomuKli", OleDb.OleDbType.Integer, Nothing, "NUMNRDOMU")
    '    dbCommandUpdate.Parameters.Add("@oNrDomuKli", OleDb.OleDbType.VarChar, 10, "NRDOMU")
    '    dbCommandUpdate.Parameters.Add("@oIDDomuKli", OleDb.OleDbType.VarChar, 10, "IDDOMU")
    '    dbCommandUpdate.Parameters.Add("@oRejonKli", OleDb.OleDbType.VarChar, 20, "REJONKLIENTA")
    '    dbCommandUpdate.Parameters.Add("@oTelefonKli", OleDb.OleDbType.VarChar, 50, "TELEFON")
    '    dbCommandUpdate.Parameters.Add("@oFaxKli", OleDb.OleDbType.VarChar, 50, "FAX")
    '    dbCommandUpdate.Parameters.Add("@oeMailKli", OleDb.OleDbType.VarChar, 50, "EMAIL")
    '    dbCommandUpdate.Parameters.Add("@oRodzSprzetuLKli", OleDb.OleDbType.Boolean, Nothing, "RODZSPRZETUL")
    '    dbCommandUpdate.Parameters.Add("@oRodzSprzetuAKli", OleDb.OleDbType.Boolean, Nothing, "RODZSPRZETUA")
    '    dbCommandUpdate.Parameters.Add("@oRodzSprzetuTKli", OleDb.OleDbType.Boolean, Nothing, "RODZSPRZETUT")
    '    dbCommandUpdate.Parameters.Add("@oOfertaTZLOZENIAKli", OleDb.OleDbType.VarChar, 4, "OFERTATZLOZENIA")
    '    dbCommandUpdate.Parameters.Add("@oOfertaPLIKKli", OleDb.OleDbType.VarChar, 100, "OFERTAPLIK")
    '    dbCommandUpdate.Parameters.Add("@oPKontaktKli", OleDb.OleDbType.VarChar, 10, "PKONTAKT")
    '    dbCommandUpdate.Parameters.Add("@oSkupUwagiKli", OleDb.OleDbType.WChar, Nothing, "PROMOTORUWAGI")
    '    dbCommandUpdate.Parameters.Add("@oSprzedOpiekunKli", OleDb.OleDbType.VarChar, 10, "OPIEKUN")
    '    dbCommandUpdate.Parameters.Add("@oSprzedOKontaktRKli", OleDb.OleDbType.VarChar, 4, "OPIEKUNOSTKONTAKTROK")
    '    dbCommandUpdate.Parameters.Add("@oSprzedOKontaktTKli", OleDb.OleDbType.Integer, Nothing, "OPIEKUNOSTKONTAKTTYDZ")
    '    dbCommandUpdate.Parameters.Add("@oSprzedOKontaktDKli", OleDb.OleDbType.VarChar, 12, "OPIEKUNOSTKONTAKTDZIEN")
    '    dbCommandUpdate.Parameters.Add("@oSprzedNKontaktRKli", OleDb.OleDbType.VarChar, 4, "OPIEKUNKOLKONTAKTROK")
    '    dbCommandUpdate.Parameters.Add("@oSprzedNKontaktTKli", OleDb.OleDbType.Integer, Nothing, "OPIEKUNKOLKONTAKTTYDZ")
    '    dbCommandUpdate.Parameters.Add("@oSprzedNKontaktDKli", OleDb.OleDbType.VarChar, 12, "OPIEKUNKOLKONTAKTDZIEN")
    '    dbCommandUpdate.Parameters.Add("@oSprzedUwagiKli", OleDb.OleDbType.WChar, Nothing, "OPIEKUNUWAGI")
    '    dbCommandUpdate.Parameters.Add("@oWpisalKli", OleDb.OleDbType.VarChar, 10, "KTOWPISAL")
    '    dbCommandUpdate.Parameters.Add("@oUwagiKli", OleDb.OleDbType.VarChar, Nothing, "UWAGI")
    '    dbCommandUpdate.Parameters.Add("@oMarkerKli", OleDb.OleDbType.Boolean, Nothing, "MARKER")
    '    dbCommandUpdate.Parameters.Add("@oMarkerPKli", OleDb.OleDbType.Boolean, Nothing, "MARKERP")
    '    dbCommandUpdate.Parameters.Add("@oZAKUPPOPROKKli", OleDb.OleDbType.Double, Nothing, "ZAKUPPOPROK")
    '    dbCommandUpdate.Parameters.Add("@oUSLUGIDODKli", OleDb.OleDbType.VarChar, 200, "USLUGIDOD")
    '    dbCommandUpdate.Parameters.Add("@oRZURAP01", OleDb.OleDbType.VarChar, 100, "RZURAP01")
    '    dbCommandUpdate.Parameters.Add("@oRZURAP02", OleDb.OleDbType.VarChar, 100, "RZURAP02")
    '    dbCommandUpdate.Parameters.Add("@oRZURAP03", OleDb.OleDbType.VarChar, 100, "RZURAP03")
    '    dbCommandUpdate.Parameters.Add("@oRZURAP04", OleDb.OleDbType.VarChar, 100, "RZURAP04")
    '    dbCommandUpdate.Parameters.Add("@oRZURAP05", OleDb.OleDbType.VarChar, 100, "RZURAP05")
    '    dbCommandUpdate.Parameters.Add("@oRZURAP06", OleDb.OleDbType.VarChar, 100, "RZURAP06")
    '    dbCommandUpdate.Parameters.Add("@oRZURAP07", OleDb.OleDbType.VarChar, 100, "RZURAP07")
    '    dbCommandUpdate.Parameters.Add("@oRZURAP08", OleDb.OleDbType.VarChar, 100, "RZURAP08")
    '    dbCommandUpdate.Parameters.Add("@oRZURAP09", OleDb.OleDbType.VarChar, 100, "RZURAP09")

    '    dbCommandUpdate.Parameters.Add("@oWersjaKli", OleDb.OleDbType.Integer, Nothing, "WERSJA")

    '    'parametry filtrowania
    '    dbParamUpd = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.VarChar, 6, "NRKARTY")
    '    dbParamUpd.SourceVersion = DataRowVersion.Original
    '    dbParamUpd = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParamUpd.SourceVersion = DataRowVersion.Original
    '    dbDataAdapter.UpdateCommand = dbCommandUpdate
    'End Sub

    'Public Sub DBInsertKlienci(ByRef dbCommandInsert As OleDb.OleDbCommand,
    '                           ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
    '    Dim dbParamIns As New OleDb.OleDbParameter
    '    '------------------------------------------
    '    'komenda INSERT
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oIdentKli", OleDb.OleDbType.VarChar, 6, "NRKARTY")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oNIPKli", OleDb.OleDbType.VarChar, 15, "NIP")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oNrTofiTxtKli", OleDb.OleDbType.VarChar, 500, "NRTOFITXT")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oKartaPaybackKli", OleDb.OleDbType.Boolean, Nothing, "KARTAPAYBACK")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oNryKartPaybackKli", OleDb.OleDbType.WChar, Nothing, "NRYKARTPAYBACK")
    '    dbParamIns.SourceVersion = DataRowVersion.Current

    '    dbParamIns = dbCommandInsert.Parameters.Add("@oWYKAZURZADZENKli", OleDb.OleDbType.WChar, Nothing, "WYKAZURZADZEN")
    '    dbParamIns.SourceVersion = DataRowVersion.Current

    '    dbParamIns = dbCommandInsert.Parameters.Add("@oNAZWA1Kli", OleDb.OleDbType.VarChar, 100, "NAZWAKLIENTA")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oMiejscKli", OleDb.OleDbType.VarChar, 50, "MIEJSCOWOSC")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oKodPoczKli", OleDb.OleDbType.VarChar, 10, "KODPOCZTOWY")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oUlicaKli", OleDb.OleDbType.VarChar, 50, "ULICA")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oNumNrDomuKli", OleDb.OleDbType.Integer, Nothing, "NUMNRDOMU")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oNrDomuKli", OleDb.OleDbType.VarChar, 10, "NRDOMU")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oIDDomuKli", OleDb.OleDbType.VarChar, 10, "IDDOMU")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oRejonKli", OleDb.OleDbType.VarChar, 20, "REJONKLIENTA")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oTelefonKli", OleDb.OleDbType.VarChar, 50, "TELEFON")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oFaxKli", OleDb.OleDbType.VarChar, 50, "FAX")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oeMailKli", OleDb.OleDbType.VarChar, 50, "EMAIL")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oRodzSprzetuLKli", OleDb.OleDbType.Boolean, Nothing, "RODZSPRZETUL")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oRodzSprzetuAKli", OleDb.OleDbType.Boolean, Nothing, "RODZSPRZETUA")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oRodzSprzetuTKli", OleDb.OleDbType.Boolean, Nothing, "RODZSPRZETUT")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oOfertaTZLOZENIAKli", OleDb.OleDbType.VarChar, 4, "OFERTATZLOZENIA")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oOfertaPLIKKli", OleDb.OleDbType.VarChar, 100, "OFERTAPLIK")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oPKontaktKli", OleDb.OleDbType.VarChar, 10, "PKONTAKT")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oSkupUwagiKli", OleDb.OleDbType.WChar, Nothing, "PROMOTORUWAGI")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oSprzedOpiekunKli", OleDb.OleDbType.VarChar, 10, "OPIEKUN")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oSprzedOKontaktRKli", OleDb.OleDbType.VarChar, 4, "OPIEKUNOSTKONTAKTROK")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oSprzedOKontaktTKli", OleDb.OleDbType.Integer, Nothing, "OPIEKUNOSTKONTAKTTYDZ")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oSprzedOKontaktDKli", OleDb.OleDbType.VarChar, 12, "OPIEKUNOSTKONTAKTDZIEN")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oSprzedNKontaktTKli", OleDb.OleDbType.VarChar, 4, "OPIEKUNKOLKONTAKTROK")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oSprzedNKontaktTKli", OleDb.OleDbType.Integer, Nothing, "OPIEKUNKOLKONTAKTTYDZ")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oSprzedNKontaktDKli", OleDb.OleDbType.VarChar, 12, "OPIEKUNKOLKONTAKTDZIEN")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oSprzedUwagiKli", OleDb.OleDbType.WChar, Nothing, "OPIEKUNUWAGI")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oWpisalKli", OleDb.OleDbType.VarChar, 10, "KTOWPISAL")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oUwagiKli", OleDb.OleDbType.VarChar, Nothing, "UWAGI")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oMarkerKli", OleDb.OleDbType.Boolean, Nothing, "MARKER")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oMarkerPKli", OleDb.OleDbType.Boolean, Nothing, "MARKERP")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oZAKUPPOPROKKli", OleDb.OleDbType.Double, Nothing, "ZAKUPPOPROK")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oUSLUGIDODKli", OleDb.OleDbType.VarChar, 200, "USLUGIDOD")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oRZURAP01", OleDb.OleDbType.VarChar, 100, "RZURAP01")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oRZURAP02", OleDb.OleDbType.VarChar, 100, "RZURAP02")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oRZURAP03", OleDb.OleDbType.VarChar, 100, "RZURAP03")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oRZURAP04", OleDb.OleDbType.VarChar, 100, "RZURAP04")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oRZURAP05", OleDb.OleDbType.VarChar, 100, "RZURAP05")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oRZURAP06", OleDb.OleDbType.VarChar, 100, "RZURAP06")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oRZURAP07", OleDb.OleDbType.VarChar, 100, "RZURAP07")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oRZURAP08", OleDb.OleDbType.VarChar, 100, "RZURAP08")
    '    dbParamIns.SourceVersion = DataRowVersion.Current
    '    dbParamIns = dbCommandInsert.Parameters.Add("@oRZURAP09", OleDb.OleDbType.VarChar, 100, "RZURAP09")
    '    dbParamIns.SourceVersion = DataRowVersion.Current

    '    dbParamIns = dbCommandInsert.Parameters.Add("@oWersjaKli", OleDb.OleDbType.Integer, Nothing, "WERSJA")
    '    dbParamIns.SourceVersion = DataRowVersion.Current

    '    dbDataAdapter.InsertCommand = dbCommandInsert
    'End Sub























End Module
