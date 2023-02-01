Module modStrukturyFiltrowanie





    '----------------------------------------------------------------
    ' zmienne wykorzystywane do zlozonego filtrowania Katalogu Klientow

    'ANALIZY
    Public fiIloscAA As String
    Public fiIloscBB As String

    Public fiIlZak00 As String
    Public fiIlZak01 As String
    Public fiIlZak02 As String
    Public fiIlZak03 As String
    Public fiIlZak04 As String
    Public fiIlZak05 As String
    Public fiIlZak06 As String
    Public fiIlZak07 As String
    Public fiIlZak08 As String
    Public fiIlZak09 As String
    Public fiIlZak10 As String
    Public fiIlZak11 As String

    Public fiIlZak12 As String
    Public fiIlZak13 As String
    Public fiIlZak14 As String
    Public fiIlZak15 As String
    Public fiIlZak16 As String
    Public fiIlZak17 As String
    Public fiIlZak18 As String
    Public fiIlZak19 As String
    Public fiIlZak20 As String
    Public fiIlZak21 As String
    Public fiIlZak22 As String
    Public fiIlZak23 As String



    'Kontakty
    Public fiIdentKon As String            'IDENTKLIENTA     Text, 6
    Public fiDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
    Public fiPracownikKon As String        'PRACOWNIK        Text, 2     uzytkownik
    Public fiTematKon As String            'TEMAT            Text, 50
    Public fiUwagiKon As String            'UWAGI            Memo













    'kLIENCI
    '..............................
    Public fiIdentKli As String       'IDENTKLIENTA   Text, 6
    Public fiNrTOFIKli As String      '!!!!  NRTOFI         long, 5
    Public fiNrTOFITxtKli As String   'NRTOFITXT      Text, 500

    Public fiKartaPaybackKli As String '!!!!  KARTAPAYBACK   Logical
    Public fiNryKartPaybackKli As String 'NRYKARTPAYBACK   memo

    Public fiNazwa1Kli As String      'NAZWA1         Text, 50
    Public fiKodPoczKli As String     'KODPOCZTOWY    Text, 10
    Public fiMiejscKli As String      'MIEJSCOWOSC    Text, 50
    Public fiUlicaKli As String       'ULICA          Text, 50
    Public fiNumNrDomuKli As String   '!!!!  NUMNRDOMU      Integer
    Public fiNrDomuKli As String      'NRDOMU         Text, 10
    Public fiIDDomuKli As String      'IDDOMU         Text, 10
    Public fiNIPKli As String         'NIP            Text, 15
    Public fiTelefonKli As String     'TELEFON        Text, 50
    Public fiFaxKli As String         'FAX            Text, 50
    Public fieMailKli As String       'EMAIL          Text, 50
    Public fiWpisalKli As String      'KTOWPISAL      Text, 10   uzytkownik
    Public fiRejonKli As String         'REJKLIENTA     Text, 20   
    Public fiPKontaktKli As String      'PKONTAKT       Text, 10   data rrrr-mm-dd

    Public fiBranzaKli As String        'BRANZA         Text, 100
    Public fiPodBranzaKli As String     'PODBRANZA      Text, 100
    Public fiLiczbaZaktrudnionychKli As String   'LICZBAZATRUDNIONYCH       INTEGER
    Public fiLiczbaPracZdalnieKli As String      'LICZBAPRACZDALNIE         INTEGER
    '......................................
    Public fiWykazUrzadzenKli As String              'WYKLAZURZADZEN    Memo

    Public fiSposobWyboruDostawcyKli As String       'SPOSOBWYBORUDOSTAWCY    Text 30
    Public fiZainstalowanoMonitorKli As String      '!!!!  ZAINSTALOWANOMONITOR    Logical
    Public fiLiczbaUrzadzenKli As String            '!!!!  LICZBAURZADZEN     Integer
    Public fiLiczbaWydrukowKli As String            '!!!!  LICZBAWYDRUKOW     Integer
    Public fiPotencjalDrukuKli As String             'POTENCJALDRUKU     Text 2

    Public fiAktZakupMatEkspKli As String           '!!!!  AKTZAKUPMATEKSP    Logical
    Public fiAktRozliczaStronyDrukuKli As String    '!!!!  AKTROZLICZASTRONYDRUKU    Logical
    Public fiAktKorzystaZNajmuKli As String         '!!!!  AKTKORZYSTAZNAJMU  Logical
    Public fiZaintRozliczaniemStronDrukuKli As String   '!!!!  ZAINTROZLICZANIEMSTRONDRUKU    Logical
    Public fiZaintKorzystaniemZNajmuKli As String   '!!!!  ZAINTKORZYSTANIEMZNAJMU    Logical

    Public fiDataWeryfSegmentacjiKli As String          'DATAWERYFSEGMENTACJI  Text 10

    Public fiPlanDlugookresowyKli As String         '!!!!  PLANDLUGOOKRESOWY     Integer
    Public fiPlanKrotkookresowyKli As String        '!!!!  PLANKROTKOOKRESOWY    Integer
    '......................................
    Public fiRodzSprzetuLKli As String  '!!!!  RODZSPRZETUL Logical
    Public fiRodzSprzetuAKli As String  '!!!!  RODZSPRZETUA Logical
    Public fiRodzSprzetuTKli As String  '!!!!  RODZSPRZETUT Logical
    Public fiOfertaTZlozeniaKli As String        'OFERTATZLOZENIA         Text, 4
    Public fiOfertaPlikKli As String    'OFERTAPLIK     Text, 100
    Public fiUwagikli As String         'UWAGI          Memo
    Public fiSkupOpiekunKli As String    'SKUPOPIEKUN    Text, 10   uzytkownik
    Public fiSkupWartoscKli As String    'SKUPWARTOSC    Text, 50

    Public fiSprzedOpiekunkli As String    'SPRZEDOPIEKUN    Text, 10   uzytkownik

    Public fiSprzedOKontaktRKli As String  'SPRZEDOKONTAKRT  Text, 4    rok
    Public fiSprzedOKontaktTKli As String  '!!!!  SPRZEDOKONTAKTT  integer    nr tygodnia
    Public fiSprzedOKontaktDKli As String  'SPRZEDOKONTAKTD  Text, 12   dzien tygodnia
    Public fiSprzedNKontaktRKli As String  'SPRZEDNKONTAKRT  Text, 4    rok
    Public fiSprzedNKontaktTKli As String  '!!!!  SPRZEDNKONTAKTT  Integer    nr tygodnia
    Public fiSprzedNKontaktDKli As String  'SPRZEDNKONTAKTD  Text, 12    nr tygodnia

    Public fiSprzedUwagiKli As String      'SPRZEDUWAGI      Memo
    Public fiMarker As String              '!!!!  MARKER       Bit
    Public fiMarkerP As String             '!!!!  MARKERP      Bit
    Public fiAktywnyKli As String          '!!!!  AKTYWNY          boolean

    Public fiZakupPopRokKli As String       '!!!!  ZAKUPPOPROK        Double
    Public fiUslugiDodKli As String         'USLUGIDOD          Text, 200
    Public fiRZURap01 As String              'RZURAP01     Text, 100
    Public fiRZURap02 As String              'RZURAP02     Text, 100
    Public fiRZURap03 As String              'RZURAP03     Text, 100
    Public fiRZURap04 As String              'RZURAP04     Text, 100
    Public fiRZURap05 As String              'RZURAP05     Text, 100
    Public fiRZURap06 As String              'RZURAP06     Text, 100
    Public fiRZURap07 As String              'RZURAP07     Text, 100
    Public fiRZURap08 As String              'RZURAP08     Text, 100
    Public fiRZURap09 As String              'RZURAP09     Text, 100

    'Public fiWersjaKli As string         'WERSJA           Integer




End Module
