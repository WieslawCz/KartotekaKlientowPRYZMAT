Module modDefinicje
    Public Const SoftarteMail As String = "czajkowski@softart.com.pl"
    Public Const SoftartWWW As String = "http://www.softart.com.pl"

    'zmienne globalne do wyliczania przedzia³ów czasowych
    Public oCzasStart As Date = DateTime.Now
    Public oCzasStop As Date = DateTime.Now

    Public oKomunikat As String = ""
    Public oWybralem As String = ""
    Public oGeneratorWydruku As String = ""
    Public oOperacja As String = ""        'identyfikator operacji DODAJ/USUN/EDYTUJ/DRUKUJ...
    Public oAktualizuj As Boolean = True   'wskaznik czy aktualizowac
    Public oCoDalej As String = ""
    Public oCoMamRobic As String = ""      'identyfikator operacji EDYTUJ/WYBIERAJ
    Public oTabela As String = ""          'Tabela
    Public oPlikNaDysku As String = ""     'nazwa pliku przekazywana
    Public oWystapilConcurrencyException As Boolean = False
    Public oData As String = ""            'do definiowania daty
    Public oWeek As Integer = 0            'tydzien w roku
    Public oPracownik As String = ""       'pracownik
    Public oKlient As String = ""
    Public oNumer As String
    Public oNumerKartyPayBack As String = ""
    Public oRodzajKontaktu As String = ""
    Public oIdentCoDalej As String = ""
    Public oIdBranzy As String = ""
    Public oIdPodBranzy As String = ""

    Public oNazwaSchematu As String = ""
    Public oSchematFiltrowania As String = ""
    Public oLokalnyFiltr As String = ""
    Public oAkcja As String = ""
    Public oWybranoAkcjeMarketingowa As Boolean = False

    'kolory na ekranie
    Public SoftartCursorBackColor As Color = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(192, Byte))
    Public SoftartCursorForeColor As Color = System.Drawing.Color.Black
    Public SoftartEditableBackColor As Color = System.Drawing.SystemColors.Window
    Public SoftartEditableForeColor As Color = System.Drawing.Color.Purple
    Public SoftartNonEditableBackColor As Color = System.Drawing.SystemColors.Control
    Public SoftartNonEditableForeColor As Color = System.Drawing.Color.Purple
    Public SoftartLabelBackColor As Color = System.Drawing.SystemColors.Control
    Public SoftartLabelForeColor As Color = System.Drawing.Color.Navy

    'standardowe wielkoœci czcionki do wyswwietlanai w formach/Oknach
    Public Softart_FormFontName As String = "Calibri"
    Public Softart_FormFontSize As Integer = 9.0!

    'standardowe wielkoœci czcionki do wyswwietlanai w formach/Oknach
    Public Softart_FormFontSizeSmall As Integer = 8.25!
    Public Softart_FormFontSizeMedium As Integer = 10.0!
    Public Softart_FormFontSizeBig As Integer = 12.0!

    'do testowania koloru tabeli
    Public SoftartColorTest As Char = ""

    'definicje pol dla generatora wydrukow
    Public GenNazwa As String = ""
    Public GenTyp As String = ""
    Public GenRozmiar As String = ""
    Public GenNaglowek As String = ""
    Public GenWyrownanie As String = ""
    Public GenMaxIlWierszy As String = ""


End Module
