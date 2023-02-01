Module modUprawnienia

    'uprawnienia Administratora
    Public Const Program_AdministratorLogin As String = "SUPERVISOR"
    Public Const Program_AdministratorHaslo As String = "SoftartOpole"
    Public Const Program_AdministratorHaslo2 As String = "PRY2M@T"
    Public Const Program_AdministratorNazwa As String = "Administrator Systemu"



    'identyfikacja uzytkownika
    Public Program_IdUzytkownika As String = ""
    Public Program_NazwaUzytkownika As String = ""
    Public Program_FunkcjaUzytkownika As String = ""
    Public Program_HasloUzytkownika As String = ""
    Public Program_UprawnieniaUzytkownika As String = ""
    Public Program_STOP As Boolean = False



    'Role w programie
    Public Const Program_RolaAdministrator As String = "A"
    Public Const Program_RolaSzef As String = "S"
    Public Const Program_RolaPracownikUprzywilejowany As String = "U"
    Public Const Program_RolaPracownik As String = "P"



    '************************************************
    ' zapisuje do katalogu UPRAWNIENIA informacje o roli uzytkownika
    ' kolumNA uprawnienia MA 100 ZNAKÓW
    ' GENERUJE LOSOWO 100 LITER i wpisuje kolejno do pola
    ' ZMIENIA pole nr 58 na oznaczenie uprawnienia
    '************************************************
    Const _Uprawnnienia_NrZnaku As Integer = 58

    Public Function ZapiszUprawnieniaUzytkownika(ByVal RolaLitera As String) As String
        Dim Upraw As String = ""
        Dim i As Integer = 0
        For i = 1 To 100
            Randomize()     'generuj ....
            Upraw &= Chr(Asc("A") - 1 + (Int(Rnd() * 26)))       '"ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Next
        Return (Mid(Upraw, 1, _Uprawnnienia_NrZnaku - 1) & RolaLitera & Mid(Upraw, _Uprawnnienia_NrZnaku + 1))
    End Function


    Public Function OdczytajUprawnieniaUzytkownika(ByVal strUprawnienia As String) As String
        Return (Mid(strUprawnienia, _Uprawnnienia_NrZnaku, 1))
    End Function



    '************************************************
    ' Testowanie uprawnien uzytkownika
    '************************************************

    Public Function _MaUprawnieniaAdministratora() As Boolean
        Return ((Program_IdUzytkownika = Program_AdministratorLogin) Or _
                (OdczytajUprawnieniaUzytkownika(Program_UprawnieniaUzytkownika) = Program_RolaAdministrator))
    End Function

    Public Function _MaUprawnieniaSzefa() As Boolean
        Return ((OdczytajUprawnieniaUzytkownika(Program_UprawnieniaUzytkownika) = Program_RolaSzef))
    End Function

    Public Function _MaUprawnieniaPracownikaUprzywilejowanego() As Boolean
        Return ((OdczytajUprawnieniaUzytkownika(Program_UprawnieniaUzytkownika) = Program_RolaPracownikUprzywilejowany))
    End Function

    Public Function _MaUprawnieniaPracownika() As Boolean
        Return ((OdczytajUprawnieniaUzytkownika(Program_UprawnieniaUzytkownika) = Program_RolaPracownik))
    End Function







    'Public Function _MaUprawnieniaAdministratora() As Boolean
    '    Return ((Program_IdUzytkownika = Program_AdministratorLogin))
    'End Function

    'Public Function _MaUprawnieniaSzefa() As Boolean
    '    Return ((OdczytajUprawnieniaUzytkownika(Program_UprawnieniaUzytkownika) = Program_RolaSzef))
    'End Function

    'Public Function _MaUprawnieniaPracownikaUprzywilejowanego() As Boolean
    '    Return ((OdczytajUprawnieniaUzytkownika(Program_UprawnieniaUzytkownika) = Program_RolaPracownikUprzywilejowany))
    'End Function

    'Public Function _MaUprawnieniaPracownika() As Boolean
    '    Return ((OdczytajUprawnieniaUzytkownika(Program_UprawnieniaUzytkownika) = Program_RolaPracownik))
    'End Function

End Module
