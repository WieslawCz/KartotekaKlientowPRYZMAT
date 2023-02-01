Module modStrukturyObroty





    '---------------------------------------------------------------------
    'Obroty
    Public oIdentObr As String            'IDENTKLIENTA     Text, 6
    Public oDataObr As String             'DATA             Text,10

    Public oLWartSprzedObr As Double      'LWARTSPRZED      Double
    Public oLIloSprzedObr As Double       'LILOPRZED        Double
    Public oLMarSprzedObr As Double       'LMARPRZED        Double

    Public oAWartSprzedObr As Double      'AWARTSPRZED      Double
    Public oAIloSprzedObr As Double       'AILOSPRZED       Double
    Public oAMARSprzedObr As Double       'AMARSPRZED       Double

    Public oLOWartSprzedObr As Double      'LOWARTSPRZED      Double
    Public oLOIloSprzedObr As Double       'LOILOPRZED        Double
    Public oLOMARSprzedObr As Double       'LOMARPRZED        Double

    Public oAOWartSprzedObr As Double      'AOWARTSPRZED      Double
    Public oAOIloSprzedObr As Double       'AOILOSPRZED       Double
    Public oAOMARSprzedObr As Double       'AOMARSPRZED       Double

    Public oNAJEMWartSprzedObr As Double      'NAJEMWARTSPRZED      Double
    Public oNAJEMIloSprzedObr As Double       'NAJEMILOPRZED        Double
    Public oNAJEMMARSprzedObr As Double       'NAJEMMARPRZED        Double

    Public oSTRONYWartSprzedObr As Double      'STRONYWARTSPRZED      Double
    Public oSTRONYIloSprzedObr As Double       'STRONYILOPRZED        Double
    Public oSTRONYMARSprzedObr As Double       'STRONYMARPRZED        Double

    Public oDRUKARKIWartSprzedObr As Double      'DRUKARKIWARTSPRZED      Double
    Public oDRUKARKIIloSprzedObr As Double       'DRUKARKIILOPRZED        Double
    Public oDRUKARKIMARSprzedObr As Double       'DRUKARKIMARPRZED        Double

    Public oSKUPWartSprzedObr As Double      'SKUPWARTSPRZED      Double
    Public oSKUPIloSprzedObr As Double       'SKUPILOPRZED        Double
    Public oSKUPMARSprzedObr As Double       'SKUPMARPRZED        Double

    Public oWersjaObr As Integer          'WERSJA           Integer

    Public sConnectionObroty As String = ConnectionStrings()
    Public HQConnectionObroty As String = HQConnectionStrings()

    Public sSelectSQLObroty As String = "SELECT " &
                                                 "IDENTKLIENTA," &
                                                 "DATA," &
                                                 "AILOSPRZED," &
                                                 "AWARTSPRZED," &
                                                 "AMARSPRZED," &
                                               "LILOSPRZED," &
                                               "LWARTSPRZED," &
                                               "LMARSPRZED," &
                                                   "AOILOSPRZED," &
                                                   "AOWARTSPRZED," &
                                                   "AOMARSPRZED," &
                                                 "LOILOSPRZED," &
                                                 "LOWARTSPRZED," &
                                                 "LOMARSPRZED," &
                                               "NAJEMILOSPRZED," &
                                               "NAJEMWARTSPRZED," &
                                               "NAJEMMARSPRZED," &
                                                 "STRONYILOSPRZED," &
                                                 "STRONYWARTSPRZED," &
                                                 "STRONYMARSPRZED," &
                                                   "DRUKARKIILOSPRZED," &
                                                   "DRUKARKIWARTSPRZED," &
                                                   "DRUKARKIMARSPRZED," &
                                                     "SKUPILOSPRZED," &
                                                     "SKUPWARTSPRZED," &
                                                     "SKUPMARSPRZED," &
                                                 "WERSJA " &
                                        "FROM Obroty "
    '----------"FROM Obroty WHERE IDENTKLIENTA='" & oIdentKli & "'"

    Public sDeleteSQLObroty As String = "DELETE FROM Obroty " &
                                        "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                              "(DATA=@orygData) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLObroty As String = "UPDATE Obroty SET " &
                                                 "AILOSPRZED=@oAIloSprzedObr," &
                                                 "AWARTSPRZED=@oAWartSprzedObr," &
                                                 "AMARSPRZED=@oAMARSprzedObr," &
                                                 "LILOSPRZED=@oLIloSprzedObr," &
                                                 "LWARTSPRZED=@oLWartSprzedObr," &
                                                 "LMARSPRZED=@oLMARSprzedObr," &
                                                   "AOILOSPRZED=@oAOIloSprzedObr," &
                                                   "AOWARTSPRZED=@oAOWartSprzedObr," &
                                                   "AOMARSPRZED=@oAOMARSprzedObr," &
                                                   "LOILOSPRZED=@oLOIloSprzedObr," &
                                                   "LOWARTSPRZED=@oLOWartSprzedObr," &
                                                   "LOMARSPRZED=@oLOMARSprzedObr," &
                                               "NAJEMILOSPRZED=@oNAJEMILOSPRZEDObr," &
                                               "NAJEMWARTSPRZED=@oNAJEMWARTSPRZEDObr," &
                                               "NAJEMMARSPRZED=@oNAJEMMARSPRZEDObr," &
                                                 "STRONYILOSPRZED=@oSTRONYILOSPRZEDObr," &
                                                 "STRONYWARTSPRZED=@oSTRONYWARTSPRZEDObr," &
                                                 "STRONYMARSPRZED=@oSTRONYMARSPRZEDObr," &
                                                   "DRUKARKIILOSPRZED=@oDRUKARKIILOSPRZEDObr," &
                                                   "DRUKARKIWARTSPRZED=@oDRUKARKIWARTSPRZEDObr," &
                                                   "DRUKARKIMARSPRZED=@oDRUKARKIMARSPRZEDObr," &
                                                     "SKUPILOSPRZED=@oSKUPILOSPRZEDObr," &
                                                     "SKUPWARTSPRZED=@oSKUPWARTSPRZEDObr," &
                                                     "SKUPMARSPRZED=@oSKUPMARSPRZEDObr," &
                                                 "WERSJA=@oWersjaObr " &
                                        "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                              "(DATA=@orygData) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sInsertSQLObroty As String = "INSERT INTO Obroty (" &
                                                 "IDENTKLIENTA," &
                                                 "DATA," &
                                                 "AILOSPRZED," &
                                                 "AWARTSPRZED," &
                                                 "AMARSPRZED," &
                                                 "LILOSPRZED," &
                                                 "LWARTSPRZED," &
                                                 "LMARSPRZED," &
                                                   "AOILOSPRZED," &
                                                   "AOWARTSPRZED," &
                                                   "AOMARSPRZED," &
                                                   "LOILOSPRZED," &
                                                   "LOWARTSPRZED," &
                                                   "LOMARSPRZED," &
                                               "NAJEMILOSPRZED," &
                                               "NAJEMWARTSPRZED," &
                                               "NAJEMMARSPRZED," &
                                                 "STRONYILOSPRZED," &
                                                 "STRONYWARTSPRZED," &
                                                 "STRONYMARSPRZED," &
                                                   "DRUKARKIILOSPRZED," &
                                                   "DRUKARKIWARTSPRZED," &
                                                   "DRUKARKIMARSPRZED," &
                                                     "SKUPILOSPRZED," &
                                                     "SKUPWARTSPRZED," &
                                                     "SKUPMARSPRZED," &
                                                 "WERSJA " &
                                                 ") " &
                                        "VALUES  (" &
                                                 "@oIdentObr," &
                                                 "@oDataObr," &
                                                 "@oAIloSprzedObr," &
                                                 "@oAWartSprzedObr," &
                                                 "@oAMARSprzedObr," &
                                                 "@oLIloSprzedObr," &
                                                 "@oLWartSprzedObr," &
                                                 "@oLMARSprzedObr," &
                                                   "@oAOIloSprzedObr," &
                                                   "@oAOWartSprzedObr," &
                                                   "@oAOMARSprzedObr," &
                                                   "@oLOIloSprzedObr," &
                                                   "@oLOWartSprzedObr," &
                                                   "@oLOMARSprzedObr," &
                                               "@oNAJEMILOSPRZEDObr," &
                                               "@oNAJEMWARTSPRZEDObr," &
                                               "@oNAJEMMARSPRZEDObr," &
                                                 "@oSTRONYILOSPRZEDObr," &
                                                 "@oSTRONYWARTSPRZEDObr," &
                                                 "@oSTRONYMARSPRZEDObr," &
                                                   "@oDRUKARKIILOSPRZEDObr," &
                                                   "@oDRUKARKIWARTSPRZEDObr," &
                                                   "@oDRUKARKIMARSPRZEDObr," &
                                                     "@oSKUPILOSPRZEDObr," &
                                                     "@oSKUPWARTSPRZEDObr," &
                                                     "@oSKUPMARSPRZEDObr," &
                                                 "@oWersjaObr " &
                                                 ")"

    'SQLDeleteObroty(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateObroty(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertObroty(sqlCommandInsert, sqlDataAdapter)


    Public Sub SQLDeleteObroty(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygData", SqlDbType.Char, 10, "DATA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    Public Sub SQLUpdateObroty(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda Update
        'parametry aktualizacji
        'sqlCommandUpdate.Parameters.Add("@oIdentObr", sqlDbType.Char, 6, "IDENTKLIENTA")
        'sqlCommandUpdate.Parameters.Add("@oDataObr", sqlDbType.Char, 10, "DATA")

        sqlCommandUpdate.Parameters.Add("@oAIloSprzedObr", SqlDbType.Float, Nothing, "AILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oAWartSprzedObr", SqlDbType.Float, Nothing, "AWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oAMARSprzedObr", SqlDbType.Float, Nothing, "AMARSPRZED")
        sqlCommandUpdate.Parameters.Add("@oLIloSprzedObr", SqlDbType.Float, Nothing, "LILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oLWartSprzedObr", SqlDbType.Float, Nothing, "LWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oLMARSprzedObr", SqlDbType.Float, Nothing, "LMARSPRZED")

        sqlCommandUpdate.Parameters.Add("@oAOIloSprzedObr", SqlDbType.Float, Nothing, "AOILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oAOWartSprzedObr", SqlDbType.Float, Nothing, "AOWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oAOMARSprzedObr", SqlDbType.Float, Nothing, "AOMARSPRZED")
        sqlCommandUpdate.Parameters.Add("@oLOIloSprzedObr", SqlDbType.Float, Nothing, "LOILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oLOWartSprzedObr", SqlDbType.Float, Nothing, "LOWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oLOMARSprzedObr", SqlDbType.Float, Nothing, "LOMARSPRZED")

        sqlCommandUpdate.Parameters.Add("@oNAJEMIloSprzedObr", SqlDbType.Float, Nothing, "NAJEMILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oNAJEMWartSprzedObr", SqlDbType.Float, Nothing, "NAJEMWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oNAJEMMARSprzedObr", SqlDbType.Float, Nothing, "NAJEMMARSPRZED")

        sqlCommandUpdate.Parameters.Add("@oSTRONYIloSprzedObr", SqlDbType.Float, Nothing, "STRONYILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oSTRONYWartSprzedObr", SqlDbType.Float, Nothing, "STRONYWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oSTRONYMARSprzedObr", SqlDbType.Float, Nothing, "STRONYMARSPRZED")

        sqlCommandUpdate.Parameters.Add("@oDRUKARKIIloSprzedObr", SqlDbType.Float, Nothing, "DRUKARKIILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oDRUKARKIWartSprzedObr", SqlDbType.Float, Nothing, "DRUKARKIWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oDRUKARKIMARSprzedObr", SqlDbType.Float, Nothing, "DRUKARKIMARSPRZED")

        sqlCommandUpdate.Parameters.Add("@oSKUPIloSprzedObr", SqlDbType.Float, Nothing, "SKUPILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oSKUPWartSprzedObr", SqlDbType.Float, Nothing, "SKUPWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oSKUPMARSprzedObr", SqlDbType.Float, Nothing, "SKUPMARSPRZED")

        sqlCommandUpdate.Parameters.Add("@oWersjaObr", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygData", SqlDbType.Char, 10, "DATA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.UpdateCommand = sqlCommandUpdate
    End Sub


    Public Sub SQLInsertObroty(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda INSERT
        sqlParam = sqlCommandInsert.Parameters.Add("@oIdentObr", SqlDbType.Char, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oDataObr", SqlDbType.Char, 10, "DATA")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oAIloSprzedObr", SqlDbType.Float, Nothing, "AILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oAWartSprzedObr", SqlDbType.Float, Nothing, "AWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oAMARSprzedObr", SqlDbType.Float, Nothing, "AMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oLIloSprzedObr", SqlDbType.Float, Nothing, "LILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oLWartSprzedObr", SqlDbType.Float, Nothing, "LWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oLMARSprzedObr", SqlDbType.Float, Nothing, "LMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oAOIloSprzedObr", SqlDbType.Float, Nothing, "AOILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oAOWartSprzedObr", SqlDbType.Float, Nothing, "AOWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oAOMARSprzedObr", SqlDbType.Float, Nothing, "AOMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oLOIloSprzedObr", SqlDbType.Float, Nothing, "LOILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oLOWartSprzedObr", SqlDbType.Float, Nothing, "LOWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oLOMARSprzedObr", SqlDbType.Float, Nothing, "LOMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oNAJEMIloSprzedObr", SqlDbType.Float, Nothing, "NAJEMILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oNAJEMWartSprzedObr", SqlDbType.Float, Nothing, "NAJEMWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oNAJEMMARSprzedObr", SqlDbType.Float, Nothing, "NAJEMMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oSTRONYIloSprzedObr", SqlDbType.Float, Nothing, "STRONYILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oSTRONYWartSprzedObr", SqlDbType.Float, Nothing, "STRONYWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oSTRONYMARSprzedObr", SqlDbType.Float, Nothing, "STRONYMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oDRUKARKIIloSprzedObr", SqlDbType.Float, Nothing, "DRUKARKIILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oDRUKARKIWartSprzedObr", SqlDbType.Float, Nothing, "DRUKARKIWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oDRUKARKIMARSprzedObr", SqlDbType.Float, Nothing, "DRUKARKIMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oSKUPIloSprzedObr", SqlDbType.Float, Nothing, "SKUPILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oSKUPWartSprzedObr", SqlDbType.Float, Nothing, "SKUPWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oSKUPMARSprzedObr", SqlDbType.Float, Nothing, "SKUPMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oWersjaObr", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub



    Public Sub DBDeleteObroty(ByRef dbCommandDelete As OleDb.OleDbCommand,
                                  ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
        Dim dbParam As OleDb.OleDbParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
        dbParam.SourceVersion = DataRowVersion.Original
        dbParam = dbCommandDelete.Parameters.Add("@orygData", OleDb.OleDbType.Char, 10, "DATA")
        dbParam.SourceVersion = DataRowVersion.Original
        dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
        dbParam.SourceVersion = DataRowVersion.Original
        dbDataAdapter.DeleteCommand = dbCommandDelete
    End Sub

    Public Sub DBUpdateObroty(ByRef dbCommandUpdate As OleDb.OleDbCommand,
                               ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
        Dim dbParam As New OleDb.OleDbParameter
        '------------------------------------------
        'komenda Update
        'parametry aktualizacji
        'dbCommandUpdate.Parameters.Add("@oIdentObr", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
        'dbCommandUpdate.Parameters.Add("@oDataObr", OleDb.OleDbType.Char, 10, "DATA")

        dbCommandUpdate.Parameters.Add("@oAIloSprzedObr", OleDb.OleDbType.Double, Nothing, "AILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oAWartSprzedObr", OleDb.OleDbType.Double, Nothing, "AWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oAMARSprzedObr", OleDb.OleDbType.Double, Nothing, "AMARSPRZED")
        dbCommandUpdate.Parameters.Add("@oLIloSprzedObr", OleDb.OleDbType.Double, Nothing, "LILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oLWartSprzedObr", OleDb.OleDbType.Double, Nothing, "LWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oLMarSprzedObr", OleDb.OleDbType.Double, Nothing, "LMARSPRZED")

        dbCommandUpdate.Parameters.Add("@oAOIloSprzedObr", OleDb.OleDbType.Double, Nothing, "AOILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oAOWartSprzedObr", OleDb.OleDbType.Double, Nothing, "AOWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oAOMARSprzedObr", OleDb.OleDbType.Double, Nothing, "AOMARSPRZED")
        dbCommandUpdate.Parameters.Add("@oLOIloSprzedObr", OleDb.OleDbType.Double, Nothing, "LOILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oLOWartSprzedObr", OleDb.OleDbType.Double, Nothing, "LOWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oLOMARSprzedObr", OleDb.OleDbType.Double, Nothing, "LOMARSPRZED")

        dbCommandUpdate.Parameters.Add("@oNAJEMIloSprzedObr", OleDb.OleDbType.Double, Nothing, "NAJEMILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oNAJEMWartSprzedObr", OleDb.OleDbType.Double, Nothing, "NAJEMWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oNAJEMMARSprzedObr", OleDb.OleDbType.Double, Nothing, "NAJEMMARSPRZED")

        dbCommandUpdate.Parameters.Add("@oSTRONYIloSprzedObr", OleDb.OleDbType.Double, Nothing, "STRONYILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oSTRONYWartSprzedObr", OleDb.OleDbType.Double, Nothing, "STRONYWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oSTRONYMARSprzedObr", OleDb.OleDbType.Double, Nothing, "STRONYMARSPRZED")

        dbCommandUpdate.Parameters.Add("@oDRUKARKIIloSprzedObr", OleDb.OleDbType.Double, Nothing, "DRUKARKIILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oDRUKARKIWartSprzedObr", OleDb.OleDbType.Double, Nothing, "DRUKARKIWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oDRUKARKIMARSprzedObr", OleDb.OleDbType.Double, Nothing, "DRUKARKIMARSPRZED")

        dbCommandUpdate.Parameters.Add("@oSKUPIloSprzedObr", OleDb.OleDbType.Double, Nothing, "SKUPILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oSKUPWartSprzedObr", OleDb.OleDbType.Double, Nothing, "SKUPWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oSKUPMARSprzedObr", OleDb.OleDbType.Double, Nothing, "SKUPMARSPRZED")

        dbCommandUpdate.Parameters.Add("@oWersjaObr", OleDb.OleDbType.Integer, Nothing, "WERSJA")

        'parametry filtrowania
        dbParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
        dbParam.SourceVersion = DataRowVersion.Original
        dbParam = dbCommandUpdate.Parameters.Add("@orygData", OleDb.OleDbType.Char, 10, "DATA")
        dbParam.SourceVersion = DataRowVersion.Original
        dbParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
        dbParam.SourceVersion = DataRowVersion.Original
        dbDataAdapter.UpdateCommand = dbCommandUpdate
    End Sub

    Public Sub DBInsertObroty(ByRef dbCommandInsert As OleDb.OleDbCommand,
                               ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
        Dim dbParam As New OleDb.OleDbParameter
        '------------------------------------------
        'komenda INSERT
        dbParam = dbCommandInsert.Parameters.Add("@oIdentObr", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oDataObr", OleDb.OleDbType.Char, 10, "DATA")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@oAIloSprzedObr", OleDb.OleDbType.Double, Nothing, "AILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oAWartSprzedObr", OleDb.OleDbType.Double, Nothing, "AWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oAMARSprzedObr", OleDb.OleDbType.Double, Nothing, "AMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oLIloSprzedObr", OleDb.OleDbType.Double, Nothing, "LILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oLWartSprzedObr", OleDb.OleDbType.Double, Nothing, "LWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oLMARSprzedObr", OleDb.OleDbType.Double, Nothing, "LMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@oAOIloSprzedObr", OleDb.OleDbType.Double, Nothing, "AOILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oAOWartSprzedObr", OleDb.OleDbType.Double, Nothing, "AOWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oAOMARSprzedObr", OleDb.OleDbType.Double, Nothing, "AOMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oLOIloSprzedObr", OleDb.OleDbType.Double, Nothing, "LOILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oLOWartSprzedObr", OleDb.OleDbType.Double, Nothing, "LOWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oLOMARSprzedObr", OleDb.OleDbType.Double, Nothing, "LOMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current


        dbParam = dbCommandInsert.Parameters.Add("@oNAJEMIloSprzedObr", OleDb.OleDbType.Double, Nothing, "NAJEMILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oNAJEMWartSprzedObr", OleDb.OleDbType.Double, Nothing, "NAJEMWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oNAJEMMARSprzedObr", OleDb.OleDbType.Double, Nothing, "NAJEMMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@oSTRONYIloSprzedObr", OleDb.OleDbType.Double, Nothing, "STRONYILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oSTRONYWartSprzedObr", OleDb.OleDbType.Double, Nothing, "STRONYWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oSTRONYMARSprzedObr", OleDb.OleDbType.Double, Nothing, "STRONYMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@oDRUKARKIIloSprzedObr", OleDb.OleDbType.Double, Nothing, "DRUKARKIILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oDRUKARKIWartSprzedObr", OleDb.OleDbType.Double, Nothing, "DRUKARKIWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oDRUKARKIMARSprzedObr", OleDb.OleDbType.Double, Nothing, "DRUKARKIMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@oSKUPIloSprzedObr", OleDb.OleDbType.Double, Nothing, "SKUPILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oSKUPWartSprzedObr", OleDb.OleDbType.Double, Nothing, "SKUPWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oSKUPMARSprzedObr", OleDb.OleDbType.Double, Nothing, "SKUPMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@oWersjaObr", OleDb.OleDbType.Integer, Nothing, "WERSJA")
        dbParam.SourceVersion = DataRowVersion.Current
        dbDataAdapter.InsertCommand = dbCommandInsert
    End Sub



    'Public oIdentObr As String            'IDENTKLIENTA     Text, 6
    'Public oDataObr As String             'DATA             Text,10

    'Public oLWartSprzedObr As Double      'LWARTSPRZED      Double
    'Public oLIloSprzedObr As Double       'LILOPRZED        Double
    'Public oLMarSprzedObr As Double       'LMARPRZED        Double

    'Public oAWartSprzedObr As Double      'AWARTSPRZED      Double
    'Public oAIloSprzedObr As Double       'AILOSPRZED       Double
    'Public oAMARSprzedObr As Double       'AMARSPRZED       Double

    'Public oLOWartSprzedObr As Double      'LOWARTSPRZED      Double
    'Public oLOIloSprzedObr As Double       'LOILOPRZED        Double
    'Public oLOMARSprzedObr As Double       'LOMARPRZED        Double

    'Public oAOWartSprzedObr As Double      'AOWARTSPRZED      Double
    'Public oAOIloSprzedObr As Double       'AOILOSPRZED       Double
    'Public oAOMARSprzedObr As Double       'AOMARSPRZED       Double

    'Public oNAJEMWartSprzedObr As Double      'NAJEMWARTSPRZED      Double
    'Public oNAJEMIloSprzedObr As Double       'NAJEMILOPRZED        Double
    'Public oNAJEMMARSprzedObr As Double       'NAJEMMARPRZED        Double

    'Public oSTRONYWartSprzedObr As Double      'STRONYWARTSPRZED      Double
    'Public oSTRONYIloSprzedObr As Double       'STRONYILOPRZED        Double
    'Public oSTRONYMARSprzedObr As Double       'STRONYMARPRZED        Double

    'Public oDRUKARKIWartSprzedObr As Double      'DRUKARKIWARTSPRZED      Double
    'Public oDRUKARKIIloSprzedObr As Double       'DRUKARKIILOPRZED        Double
    'Public oDRUKARKIMARSprzedObr As Double       'DRUKARKIMARPRZED        Double

    'Public oSKUPWartSprzedObr As Double      'SKUPWARTSPRZED      Double
    'Public oSKUPIloSprzedObr As Double       'SKUPILOPRZED        Double
    'Public oSKUPMARSprzedObr As Double       'SKUPMARPRZED        Double

    'Public oWersjaObr As Integer          'WERSJA           Integer

    Public Function IdentObroty(ByVal IdentK As String, ByVal IdentM As String) As Boolean
        Dim dbSelectSQLObroty As String = sSelectSQLObroty & " WHERE (IDENTKLIENTA='" & IdentK & "') AND (SUBSTRING(DATA,1,7)='" & IdentM & "') "

        Dim dbConnectionObroty As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectObroty As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterObroty As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionObroty As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectObroty As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterObroty As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetObroty As New DataSet
        Dim DataViewObroty As New DataView
        Dim ConnectionStateObroty As System.Data.ConnectionState

        Dim Wynik As Boolean = False

        oIdentObr = ""
        oDataObr = ""
        oLWartSprzedObr = 0
        oLIloSprzedObr = 0
        oLMarSprzedObr = 0
        oAWartSprzedObr = 0
        oAIloSprzedObr = 0
        oAMARSprzedObr = 0
        oLOWartSprzedObr = 0
        oLOIloSprzedObr = 0
        oLOMARSprzedObr = 0
        oAOWartSprzedObr = 0
        oAOIloSprzedObr = 0
        oAOMARSprzedObr = 0

        oNAJEMWartSprzedObr = 0
        oNAJEMIloSprzedObr = 0
        oNAJEMMARSprzedObr = 0

        oSTRONYWartSprzedObr = 0
        oSTRONYIloSprzedObr = 0
        oSTRONYMARSprzedObr = 0

        oDRUKARKIWartSprzedObr = 0
        oDRUKARKIIloSprzedObr = 0
        oDRUKARKIMARSprzedObr = 0

        oSKUPWartSprzedObr = 0
        oSKUPIloSprzedObr = 0
        oSKUPMARSprzedObr = 0

        oWersjaObr = 0

        DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionObroty = New OleDb.OleDbConnection(sConnectionKontakty)
            dbCommandSelectObroty = New OleDb.OleDbCommand(dbSelectSQLObroty, dbConnectionObroty)
            dbDataAdapterObroty = New OleDb.OleDbDataAdapter(dbCommandSelectObroty)

            'mapujemy domyslna nazwe tabeli
            Dim dbMapowanieTabeliObroty As System.Data.Common.DataTableMapping
            dbMapowanieTabeliObroty = dbDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionObroty.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObroty = dbConnectionObroty.State
            End Try
        Else
            sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectObroty = New SqlClient.SqlCommand(dbSelectSQLObroty, sqlConnectionObroty)
            sqlDataAdapterObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliObroty As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliObroty = sqlDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObroty.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObroty = sqlConnectionObroty.State
            End Try
        End If

        If ConnectionStateObroty = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    'dbDataAdapterObroty.Fill(DataSetObroty)
                    'dbConnectionObroty.Close()
                Else
                    sqlDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    sqlDataAdapterObroty.Fill(DataSetObroty)
                    sqlConnectionObroty.Close()
                End If

                'definiuj DataView
                DataViewObroty = New DataView(DataSetObroty.Tables(0))
                If DataViewObroty.Count > 0 Then
                    oIdentObr = GetTxtDRVField(DataViewObroty.Item(0), "IDENTKLIENTA")
                    oDataObr = GetTxtDRVField(DataViewObroty.Item(0), "DATA")
                    oLWartSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "LWARTSPRZED")
                    oLIloSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "LILOSPRZED")
                    oLMarSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "LMARSPRZED")
                    oAWartSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "AWARTSPRZED")
                    oAIloSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "AILOSPRZED")
                    oAMARSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "AMARSPRZED")
                    oLOWartSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "LOWARTSPRZED")
                    oLOIloSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "LOILOSPRZED")
                    oLOMARSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "LOMARSPRZED")
                    oAOWartSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "AOWARTSPRZED")
                    oAOIloSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "AOILOSPRZED")
                    oAOMARSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "AOMARSPRZED")

                    oNAJEMWartSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "NAJEMWARTSPRZED")
                    oNAJEMIloSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "NAJEMILOSPRZED")
                    oNAJEMMARSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "NAJEMMARSPRZED")

                    oSTRONYWartSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "STRONYWARTSPRZED")
                    oSTRONYIloSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "STRONYILOSPRZED")
                    oSTRONYMARSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "STRONYMARSPRZED")

                    oDRUKARKIWartSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "DRUKARKIWARTSPRZED")
                    oDRUKARKIIloSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "DRUKARKIILOSPRZED")
                    oDRUKARKIMARSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "DRUKARKIMARSPRZED")

                    oSKUPWartSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "SKUPWARTSPRZED")
                    oSKUPIloSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "SKUPILOSPRZED")
                    oSKUPMARSprzedObr = GetDblDRVField(DataViewObroty.Item(0), "SKUPMARSPRZED")

                    oWersjaObr = GetIntDRVField(DataViewObroty.Item(0), "WERSJA")

                    Wynik = True
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wynik)
    End Function


    Public Function WartoscObrotow(ByVal IdentK As String,
                                   ByVal IdentM As String) As Double
        Dim dbSelectSQLObroty As String = sSelectSQLObroty & " WHERE (IDENTKLIENTA='" & IdentK & "') AND (SUBSTRING(DATA,1,7)='" & IdentM & "') "

        Dim dbConnectionObroty As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectObroty As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterObroty As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionObroty As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectObroty As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterObroty As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetObroty As New DataSet
        Dim DataViewObroty As New DataView
        Dim ConnectionStateObroty As System.Data.ConnectionState
        Dim i As Integer = 0
        Dim Wartosc As Double = 0

        DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionObroty = New OleDb.OleDbConnection(sConnectionKontakty)
            dbCommandSelectObroty = New OleDb.OleDbCommand(dbSelectSQLObroty, dbConnectionObroty)
            dbDataAdapterObroty = New OleDb.OleDbDataAdapter(dbCommandSelectObroty)

            'mapujemy domyslna nazwe tabeli
            Dim dbMapowanieTabeliObroty As System.Data.Common.DataTableMapping
            dbMapowanieTabeliObroty = dbDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionObroty.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObroty = dbConnectionObroty.State
            End Try
        Else
            sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectObroty = New SqlClient.SqlCommand(dbSelectSQLObroty, sqlConnectionObroty)
            sqlDataAdapterObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliObroty As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliObroty = sqlDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObroty.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObroty = sqlConnectionObroty.State
            End Try
        End If

        If ConnectionStateObroty = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    'dbDataAdapterObroty.Fill(DataSetObroty)
                    'dbConnectionObroty.Close()
                Else
                    sqlDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    sqlDataAdapterObroty.Fill(DataSetObroty)
                    sqlConnectionObroty.Close()
                End If

                'definiuj DataView
                DataViewObroty = New DataView(DataSetObroty.Tables(0))
                If DataViewObroty.Count > 0 Then
                    For i = 0 To DataViewObroty.Count - 1
                        Wartosc += GetDblDRVField(DataViewObroty.Item(i), "LWARTSPRZED") +
                                   GetDblDRVField(DataViewObroty.Item(i), "AWARTSPRZED") +
                                   GetDblDRVField(DataViewObroty.Item(i), "LOWARTSPRZED") +
                                   GetDblDRVField(DataViewObroty.Item(i), "AOWARTSPRZED") +
                                   GetDblDRVField(DataViewObroty.Item(i), "NAJEMWARTSPRZED") +
                                   GetDblDRVField(DataViewObroty.Item(i), "STRONYWARTSPRZED") +
                                   GetDblDRVField(DataViewObroty.Item(i), "DRUKARKIWARTSPRZED") +
                                   GetDblDRVField(DataViewObroty.Item(i), "SKUPWARTSPRZED")
                    Next i
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wartosc)
    End Function


    Public Function WartoscObrotowMaterialow(ByVal IdentK As String,
                                   ByVal IdentM As String) As Double
        Dim dbSelectSQLObroty As String = sSelectSQLObroty & " WHERE (IDENTKLIENTA='" & IdentK & "') AND (SUBSTRING(DATA,1,7)='" & IdentM & "') "

        Dim dbConnectionObroty As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectObroty As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterObroty As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionObroty As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectObroty As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterObroty As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetObroty As New DataSet
        Dim DataViewObroty As New DataView
        Dim ConnectionStateObroty As System.Data.ConnectionState
        Dim i As Integer = 0
        Dim Wartosc As Double = 0

        DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionObroty = New OleDb.OleDbConnection(sConnectionKontakty)
            dbCommandSelectObroty = New OleDb.OleDbCommand(dbSelectSQLObroty, dbConnectionObroty)
            dbDataAdapterObroty = New OleDb.OleDbDataAdapter(dbCommandSelectObroty)

            'mapujemy domyslna nazwe tabeli
            Dim dbMapowanieTabeliObroty As System.Data.Common.DataTableMapping
            dbMapowanieTabeliObroty = dbDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionObroty.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObroty = dbConnectionObroty.State
            End Try
        Else
            sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectObroty = New SqlClient.SqlCommand(dbSelectSQLObroty, sqlConnectionObroty)
            sqlDataAdapterObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliObroty As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliObroty = sqlDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObroty.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObroty = sqlConnectionObroty.State
            End Try
        End If

        If ConnectionStateObroty = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    'dbDataAdapterObroty.Fill(DataSetObroty)
                    'dbConnectionObroty.Close()
                Else
                    sqlDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    sqlDataAdapterObroty.Fill(DataSetObroty)
                    sqlConnectionObroty.Close()
                End If

                'definiuj DataView
                DataViewObroty = New DataView(DataSetObroty.Tables(0))
                If DataViewObroty.Count > 0 Then
                    For i = 0 To DataViewObroty.Count - 1
                        Wartosc += GetDblDRVField(DataViewObroty.Item(i), "LWARTSPRZED") +
                                   GetDblDRVField(DataViewObroty.Item(i), "AWARTSPRZED") +
                                   GetDblDRVField(DataViewObroty.Item(i), "LOWARTSPRZED") +
                                   GetDblDRVField(DataViewObroty.Item(i), "AOWARTSPRZED")
                    Next i
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wartosc)
    End Function


    Public Function WartoscObrotowUslug(ByVal IdentK As String,
                                   ByVal IdentM As String) As Double
        Dim dbSelectSQLObroty As String = sSelectSQLObroty & " WHERE (IDENTKLIENTA='" & IdentK & "') AND (SUBSTRING(DATA,1,7)='" & IdentM & "') "

        Dim dbConnectionObroty As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectObroty As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterObroty As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionObroty As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectObroty As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterObroty As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetObroty As New DataSet
        Dim DataViewObroty As New DataView
        Dim ConnectionStateObroty As System.Data.ConnectionState
        Dim i As Integer = 0
        Dim Wartosc As Double = 0

        DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionObroty = New OleDb.OleDbConnection(sConnectionKontakty)
            dbCommandSelectObroty = New OleDb.OleDbCommand(dbSelectSQLObroty, dbConnectionObroty)
            dbDataAdapterObroty = New OleDb.OleDbDataAdapter(dbCommandSelectObroty)

            'mapujemy domyslna nazwe tabeli
            Dim dbMapowanieTabeliObroty As System.Data.Common.DataTableMapping
            dbMapowanieTabeliObroty = dbDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionObroty.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObroty = dbConnectionObroty.State
            End Try
        Else
            sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectObroty = New SqlClient.SqlCommand(dbSelectSQLObroty, sqlConnectionObroty)
            sqlDataAdapterObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliObroty As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliObroty = sqlDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObroty.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObroty = sqlConnectionObroty.State
            End Try
        End If

        If ConnectionStateObroty = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    'dbDataAdapterObroty.Fill(DataSetObroty)
                    'dbConnectionObroty.Close()
                Else
                    sqlDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    sqlDataAdapterObroty.Fill(DataSetObroty)
                    sqlConnectionObroty.Close()
                End If

                'definiuj DataView
                DataViewObroty = New DataView(DataSetObroty.Tables(0))
                If DataViewObroty.Count > 0 Then
                    For i = 0 To DataViewObroty.Count - 1
                        Wartosc += GetDblDRVField(DataViewObroty.Item(i), "NAJEMWARTSPRZED") +
                                   GetDblDRVField(DataViewObroty.Item(i), "STRONYWARTSPRZED") +
                                   GetDblDRVField(DataViewObroty.Item(i), "DRUKARKIWARTSPRZED")         '+ GetDblDRVField(DataViewObroty.Item(i), "SKUPWARTSPRZED")
                    Next i
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wartosc)
    End Function



    Public Function WartoscObrotowSkup(ByVal IdentK As String,
                                   ByVal IdentM As String) As Double
        Dim dbSelectSQLObroty As String = sSelectSQLObroty & " WHERE (IDENTKLIENTA='" & IdentK & "') AND (SUBSTRING(DATA,1,7)='" & IdentM & "') "

        Dim dbConnectionObroty As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectObroty As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterObroty As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionObroty As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectObroty As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterObroty As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetObroty As New DataSet
        Dim DataViewObroty As New DataView
        Dim ConnectionStateObroty As System.Data.ConnectionState
        Dim i As Integer = 0
        Dim Wartosc As Double = 0

        DataSetObroty.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionObroty = New OleDb.OleDbConnection(sConnectionKontakty)
            dbCommandSelectObroty = New OleDb.OleDbCommand(dbSelectSQLObroty, dbConnectionObroty)
            dbDataAdapterObroty = New OleDb.OleDbDataAdapter(dbCommandSelectObroty)

            'mapujemy domyslna nazwe tabeli
            Dim dbMapowanieTabeliObroty As System.Data.Common.DataTableMapping
            dbMapowanieTabeliObroty = dbDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionObroty.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObroty = dbConnectionObroty.State
            End Try
        Else
            sqlConnectionObroty = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectObroty = New SqlClient.SqlCommand(dbSelectSQLObroty, sqlConnectionObroty)
            sqlDataAdapterObroty = New SqlClient.SqlDataAdapter(sqlCommandSelectObroty)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliObroty As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliObroty = sqlDataAdapterObroty.TableMappings.Add("Table", "TABELA_Obroty")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObroty.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObroty = sqlConnectionObroty.State
            End Try
        End If

        If ConnectionStateObroty = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    'dbDataAdapterObroty.Fill(DataSetObroty)
                    'dbConnectionObroty.Close()
                Else
                    sqlDataAdapterObroty.FillSchema(DataSetObroty, SchemaType.Mapped, "TABELA_Obroty")
                    sqlDataAdapterObroty.Fill(DataSetObroty)
                    sqlConnectionObroty.Close()
                End If

                'definiuj DataView
                DataViewObroty = New DataView(DataSetObroty.Tables(0))
                If DataViewObroty.Count > 0 Then
                    For i = 0 To DataViewObroty.Count - 1
                        Wartosc += GetDblDRVField(DataViewObroty.Item(i), "SKUPWARTSPRZED")
                    Next i
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wartosc)
    End Function














    '---------------------------------------------------------------------
    'ObrotyMies
    Public oIdentMies As String            'IDENTKLIENTA     Text, 6
    Public oMiesiacMies As String          'MIESIAC          Text,7

    Public oLWartSprzedMies As Double      'LWARTSPRZED      Double
    Public oLIloSprzedMies As Double       'LILOSPRZED       Double
    Public oLMARSprzedMies As Double       'LMARSPRZED       Double

    Public oAWartSprzedMies As Double      'AWARTSPRZED      Double
    Public oAIloSprzedMies As Double       'AILOSPRZED       Double
    Public oAMARSprzedMies As Double       'AMARSPRZED       Double

    Public oLOWartSprzedMies As Double      'LOWARTSPRZED      Double
    Public oLOIloSprzedMies As Double       'LOILOSPRZED       Double
    Public oLOMARSprzedMies As Double       'LOMARSPRZED       Double

    Public oAOWartSprzedMies As Double      'AOWARTSPRZED      Double
    Public oAOIloSprzedMies As Double       'AOILOSPRZED       Double
    Public oAOMARSprzedMies As Double       'AOMARSPRZED       Double

    Public oNAJEMWartSprzedMies As Double      'NAJEMWARTSPRZED      Double
    Public oNAJEMIloSprzedMies As Double       'NAJEMILOPRZED        Double
    Public oNAJEMMARSprzedMies As Double       'NAJEMMARPRZED        Double

    Public oSTRONYWartSprzedMies As Double      'STRONYWARTSPRZED      Double
    Public oSTRONYIloSprzedMies As Double       'STRONYILOPRZED        Double
    Public oSTRONYMARSprzedMies As Double       'STRONYMARPRZED        Double

    Public oDRUKARKIWartSprzedMies As Double      'DRUKARKIWARTSPRZED      Double
    Public oDRUKARKIIloSprzedMies As Double       'DRUKARKIILOPRZED        Double
    Public oDRUKARKIMARSprzedMies As Double       'DRUKARKIMARPRZED        Double

    Public oSKUPWartSprzedMies As Double      'SKUPWARTSPRZED      Double
    Public oSKUPIloSprzedMies As Double       'SKUPILOPRZED        Double
    Public oSKUPMARSprzedMies As Double       'SKUPMARPRZED        Double

    Public oWersjaMies As Integer          'WERSJA           Integer

    Public sConnectionObrotyMies As String = ConnectionStrings()
    Public HQConnectionObrotyMies As String = HQConnectionStrings()

    Public sSelectSQLObrotyMies As String = "SELECT " &
                                                 "IDENTKLIENTA," &
                                                 "MIESIAC," &
                                                 "AILOSPRZED," &
                                                 "AWARTSPRZED," &
                                                 "AMARSPRZED," &
                                                 "LILOSPRZED," &
                                                 "LWARTSPRZED," &
                                                 "LMARSPRZED," &
                                                   "AOILOSPRZED," &
                                                   "AOWARTSPRZED," &
                                                   "AOMARSPRZED," &
                                                   "LOILOSPRZED," &
                                                   "LOWARTSPRZED," &
                                                   "LOMARSPRZED," &
                                               "NAJEMILOSPRZED," &
                                               "NAJEMWARTSPRZED," &
                                               "NAJEMMARSPRZED," &
                                                 "STRONYILOSPRZED," &
                                                 "STRONYWARTSPRZED," &
                                                 "STRONYMARSPRZED," &
                                                   "DRUKARKIILOSPRZED," &
                                                   "DRUKARKIWARTSPRZED," &
                                                   "DRUKARKIMARSPRZED," &
                                                     "SKUPILOSPRZED," &
                                                     "SKUPWARTSPRZED," &
                                                     "SKUPMARSPRZED," &
                                                 "WERSJA " &
                                            "FROM ObrotyMies "
    '--------"FROM ObrotyMies WHERE IDENTKLIENTA='" & oIdentKli & "'"

    Public sDeleteSQLObrotyMies As String = "DELETE FROM ObrotyMies " &
                                        "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                              "(MIESIAC=@orygMies) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sUpdateSQLObrotyMies As String = "UPDATE ObrotyMies SET " &
                                                 "AILOSPRZED=@oAIloSprzedMies," &
                                                 "AWARTSPRZED=@oAWartSprzedMies," &
                                                 "AMARSPRZED=@oAMARSprzedMies," &
                                                 "LILOSPRZED=@oLIloSprzedMies," &
                                                 "LWARTSPRZED=@oLWartSprzedMies," &
                                                 "LMARSPRZED=@oLMarSprzedMies," &
                                                   "AOILOSPRZED=@oAOIloSprzedMies," &
                                                   "AOWARTSPRZED=@oAOWartSprzedMies," &
                                                   "AOMARSPRZED=@oAOMarSprzedMies," &
                                                   "LOILOSPRZED=@oLOIloSprzedMies," &
                                                   "LOWARTSPRZED=@oLOWartSprzedMies," &
                                                   "LOMARSPRZED=@oLOMarSprzedMies," &
                                                "NAJEMILOSPRZED=@oNAJEMILOSPRZEDMies," &
                                               "NAJEMWARTSPRZED=@oNAJEMWARTSPRZEDMies," &
                                               "NAJEMMARSPRZED=@oNAJEMMARSPRZEDMies," &
                                                 "STRONYILOSPRZED=@oSTRONYILOSPRZEDMies," &
                                                 "STRONYWARTSPRZED=@oSTRONYWARTSPRZEDMies," &
                                                 "STRONYMARSPRZED=@oSTRONYMARSPRZEDMies," &
                                                   "DRUKARKIILOSPRZED=@oDRUKARKIILOSPRZEDMies," &
                                                   "DRUKARKIWARTSPRZED=@oDRUKARKIWARTSPRZEDMies," &
                                                   "DRUKARKIMARSPRZED=@oDRUKARKIMARSPRZEDMies," &
                                                     "SKUPILOSPRZED=@oSKUPILOSPRZEDMies," &
                                                     "SKUPWARTSPRZED=@oSKUPWARTSPRZEDMies," &
                                                     "SKUPMARSPRZED=@oSKUPMARSPRZEDMies," &
                                                "WERSJA=@oWersjaMies " &
                                        "WHERE (IDENTKLIENTA=@orygSymbol) AND " &
                                              "(MIESIAC=@orygMies) AND " &
                                              "(WERSJA=@orygWersja)"

    Public sInsertSQLObrotyMies As String = "INSERT INTO ObrotyMies (" &
                                                 "IDENTKLIENTA," &
                                                 "MIESIAC," &
                                                 "AILOSPRZED," &
                                                 "AWARTSPRZED," &
                                                 "AMARSPRZED," &
                                                 "LILOSPRZED," &
                                                 "LWARTSPRZED," &
                                                 "LMARSPRZED," &
                                                   "AOILOSPRZED," &
                                                   "AOWARTSPRZED," &
                                                   "AOMARSPRZED," &
                                                   "LOILOSPRZED," &
                                                   "LOWARTSPRZED," &
                                                   "LOMARSPRZED," &
                                               "NAJEMILOSPRZED," &
                                               "NAJEMWARTSPRZED," &
                                               "NAJEMMARSPRZED," &
                                                 "STRONYILOSPRZED," &
                                                 "STRONYWARTSPRZED," &
                                                 "STRONYMARSPRZED," &
                                                   "DRUKARKIILOSPRZED," &
                                                   "DRUKARKIWARTSPRZED," &
                                                   "DRUKARKIMARSPRZED," &
                                                     "SKUPILOSPRZED," &
                                                     "SKUPWARTSPRZED," &
                                                     "SKUPMARSPRZED," &
                                                 "WERSJA " &
                                                 ") " &
                                        "VALUES  (" &
                                                 "@oIdentMies," &
                                                 "@oMiesMies," &
                                                  "@oAIloSprzedMies," &
                                                  "@oAWartSprzedMies," &
                                                  "@oAMARSprzedMies," &
                                                  "@oLIloSprzedMies," &
                                                  "@oLWartSprzedMies," &
                                                  "@oLMARSprzedMies," &
                                                   "@oAOIloSprzedMies," &
                                                   "@oAOWartSprzedMies," &
                                                   "@oAOMARSprzedMies," &
                                                   "@oLOIloSprzedMies," &
                                                   "@oLOWartSprzedMies," &
                                                   "@oLOMARSprzedMies," &
                                               "@oNAJEMILOSPRZEDMies," &
                                               "@oNAJEMWARTSPRZEDMies," &
                                               "@oNAJEMMARSPRZEDMies," &
                                                 "@oSTRONYILOSPRZEDMies," &
                                                 "@oSTRONYWARTSPRZEDMies," &
                                                 "@oSTRONYMARSPRZEDMies," &
                                                   "@oDRUKARKIILOSPRZEDMies," &
                                                   "@oDRUKARKIWARTSPRZEDMies," &
                                                   "@oDRUKARKIMARSPRZEDMies," &
                                                     "@oSKUPILOSPRZEDMies," &
                                                     "@oSKUPWARTSPRZEDMies," &
                                                     "@oSKUPMARSPRZEDMies," &
                                                 "@oWersjaMies " &
                                                 ")"


    'SQLDeleteObrotyMies(sqlCommandDelete, sqlDataAdapter)
    'SQLUpdateObrotyMies(sqlCommandUpdate, sqlDataAdapter)
    'SQLInsertObrotyMies(sqlCommandInsert, sqlDataAdapter)


    Public Sub SQLDeleteObrotyMies(ByRef sqlCommandDelete As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        sqlParam = sqlCommandDelete.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygMies", SqlDbType.Char, 7, "MIESIAC")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandDelete.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.DeleteCommand = sqlCommandDelete
    End Sub

    Public Sub SQLUpdateObrotyMies(ByRef sqlCommandUpdate As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda Update
        'parametry aktualizacji
        'sqlCommandUpdate.Parameters.Add("@oIdentMies", sqlDbType.Char, 6, "IDENTKLIENTA")
        'sqlCommandUpdate.Parameters.Add("@oMiesMies", sqlDbType.Char, 7, "MIESIAC")

        sqlCommandUpdate.Parameters.Add("@oAIloSprzedMies", SqlDbType.Float, Nothing, "AILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oAWartSprzedMies", SqlDbType.Float, Nothing, "AWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oAMARSprzedMies", SqlDbType.Float, Nothing, "AMARSPRZED")
        sqlCommandUpdate.Parameters.Add("@oLIloSprzedMies", SqlDbType.Float, Nothing, "LILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oLWartSprzedMies", SqlDbType.Float, Nothing, "LWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oLMARSprzedMies", SqlDbType.Float, Nothing, "LMARSPRZED")

        sqlCommandUpdate.Parameters.Add("@oAOIloSprzedMies", SqlDbType.Float, Nothing, "AOILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oAOWartSprzedMies", SqlDbType.Float, Nothing, "AOWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oAOMARSprzedMies", SqlDbType.Float, Nothing, "AOMARSPRZED")
        sqlCommandUpdate.Parameters.Add("@oLOIloSprzedMies", SqlDbType.Float, Nothing, "LOILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oLOWartSprzedMies", SqlDbType.Float, Nothing, "LOWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oLOMARSprzedMies", SqlDbType.Float, Nothing, "LOMARSPRZED")

        sqlCommandUpdate.Parameters.Add("@oNAJEMIloSprzedMies", SqlDbType.Float, Nothing, "NAJEMILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oNAJEMWartSprzedMies", SqlDbType.Float, Nothing, "NAJEMWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oNAJEMMARSprzedMies", SqlDbType.Float, Nothing, "NAJEMMARSPRZED")

        sqlCommandUpdate.Parameters.Add("@oSTRONYIloSprzedMies", SqlDbType.Float, Nothing, "STRONYILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oSTRONYWartSprzedMies", SqlDbType.Float, Nothing, "STRONYWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oSTRONYMARSprzedMies", SqlDbType.Float, Nothing, "STRONYMARSPRZED")

        sqlCommandUpdate.Parameters.Add("@oDRUKARKIIloSprzedMies", SqlDbType.Float, Nothing, "DRUKARKIILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oDRUKARKIWartSprzedMies", SqlDbType.Float, Nothing, "DRUKARKIWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oDRUKARKIMARSprzedMies", SqlDbType.Float, Nothing, "DRUKARKIMARSPRZED")

        sqlCommandUpdate.Parameters.Add("@oSKUPIloSprzedMies", SqlDbType.Float, Nothing, "SKUPILOSPRZED")
        sqlCommandUpdate.Parameters.Add("@oSKUPWartSprzedMies", SqlDbType.Float, Nothing, "SKUPWARTSPRZED")
        sqlCommandUpdate.Parameters.Add("@oSKUPMARSprzedMies", SqlDbType.Float, Nothing, "SKUPMARSPRZED")

        sqlCommandUpdate.Parameters.Add("@oWersjaMies", SqlDbType.Int, Nothing, "WERSJA")

        'parametry filtrowania
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygSymbol", SqlDbType.Char, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygMies", SqlDbType.Char, 10, "MIESIAC")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlParam = sqlCommandUpdate.Parameters.Add("@orygWersja", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Original
        sqlDataAdapter.UpdateCommand = sqlCommandUpdate
    End Sub


    Public Sub SQLInsertObrotyMies(ByRef sqlCommandInsert As SqlClient.SqlCommand,
                               ByRef sqlDataAdapter As SqlClient.SqlDataAdapter)
        Dim sqlParam As New SqlClient.SqlParameter
        '------------------------------------------
        'komenda INSERT
        sqlParam = sqlCommandInsert.Parameters.Add("@oIdentMies", SqlDbType.Char, 6, "IDENTKLIENTA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oMiesMies", SqlDbType.Char, 10, "MIESIAC")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oAIloSprzedMies", SqlDbType.Float, Nothing, "AILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oAWartSprzedMies", SqlDbType.Float, Nothing, "AWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oAMARSprzedMies", SqlDbType.Float, Nothing, "AMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oLIloSprzedMies", SqlDbType.Float, Nothing, "LILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oLWartSprzedMies", SqlDbType.Float, Nothing, "LWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oLMARSprzedMies", SqlDbType.Float, Nothing, "LMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oAOIloSprzedMies", SqlDbType.Float, Nothing, "AOILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oAOWartSprzedMies", SqlDbType.Float, Nothing, "AOWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oAOMARSprzedMies", SqlDbType.Float, Nothing, "AOMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oLOIloSprzedMies", SqlDbType.Float, Nothing, "LOILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oLOWartSprzedMies", SqlDbType.Float, Nothing, "LOWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oLOMARSprzedMies", SqlDbType.Float, Nothing, "LOMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oNAJEMIloSprzedMies", SqlDbType.Float, Nothing, "NAJEMILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oNAJEMWartSprzedMies", SqlDbType.Float, Nothing, "NAJEMWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oNAJEMMARSprzedMies", SqlDbType.Float, Nothing, "NAJEMMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oSTRONYIloSprzedMies", SqlDbType.Float, Nothing, "STRONYILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oSTRONYWartSprzedMies", SqlDbType.Float, Nothing, "STRONYWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oSTRONYMARSprzedMies", SqlDbType.Float, Nothing, "STRONYMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oDRUKARKIIloSprzedMies", SqlDbType.Float, Nothing, "DRUKARKIILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oDRUKARKIWartSprzedMies", SqlDbType.Float, Nothing, "DRUKARKIWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oDRUKARKIMARSprzedMies", SqlDbType.Float, Nothing, "DRUKARKIMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oSKUPIloSprzedMies", SqlDbType.Float, Nothing, "SKUPILOSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oSKUPWartSprzedMies", SqlDbType.Float, Nothing, "SKUPWARTSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlParam = sqlCommandInsert.Parameters.Add("@oSKUPMARSprzedMies", SqlDbType.Float, Nothing, "SKUPMARSPRZED")
        sqlParam.SourceVersion = DataRowVersion.Current

        sqlParam = sqlCommandInsert.Parameters.Add("@oWersjaMies", SqlDbType.Int, Nothing, "WERSJA")
        sqlParam.SourceVersion = DataRowVersion.Current
        sqlDataAdapter.InsertCommand = sqlCommandInsert
    End Sub



    Public Sub DBDeleteObrotyMies(ByRef dbCommandDelete As OleDb.OleDbCommand,
                                  ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
        Dim dbParam As OleDb.OleDbParameter
        '------------------------------------------
        'komenda DELETE
        'parametry filtrowania
        dbParam = dbCommandDelete.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
        dbParam.SourceVersion = DataRowVersion.Original
        dbParam = dbCommandDelete.Parameters.Add("@orygMies", OleDb.OleDbType.Char, 7, "MIESIAC")
        dbParam.SourceVersion = DataRowVersion.Original
        dbParam = dbCommandDelete.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
        dbParam.SourceVersion = DataRowVersion.Original
        dbDataAdapter.DeleteCommand = dbCommandDelete
    End Sub

    Public Sub DBUpdateObrotyMies(ByRef dbCommandUpdate As OleDb.OleDbCommand,
                               ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
        Dim dbParam As New OleDb.OleDbParameter
        '------------------------------------------
        'komenda Update
        'parametry aktualizacji
        'dbCommandUpdate.Parameters.Add("@oIdentMies", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
        'dbCommandUpdate.Parameters.Add("@oMiesMies", OleDb.OleDbType.Char, 7, "MIESIAC")

        dbCommandUpdate.Parameters.Add("@oAIloSprzedMies", OleDb.OleDbType.Double, Nothing, "AILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oAWartSprzedMies", OleDb.OleDbType.Double, Nothing, "AWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oAMARSprzedMies", OleDb.OleDbType.Double, Nothing, "AMARSPRZED")
        dbCommandUpdate.Parameters.Add("@oLIloSprzedMies", OleDb.OleDbType.Double, Nothing, "LILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oLWartSprzedMies", OleDb.OleDbType.Double, Nothing, "LWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oLMarSprzedMies", OleDb.OleDbType.Double, Nothing, "LMARSPRZED")

        dbCommandUpdate.Parameters.Add("@oAOIloSprzedMies", OleDb.OleDbType.Double, Nothing, "AOILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oAOWartSprzedMies", OleDb.OleDbType.Double, Nothing, "AOWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oAOMARSprzedMies", OleDb.OleDbType.Double, Nothing, "AOMARSPRZED")
        dbCommandUpdate.Parameters.Add("@oLOIloSprzedMies", OleDb.OleDbType.Double, Nothing, "LOILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oLOWartSprzedMies", OleDb.OleDbType.Double, Nothing, "LOWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oLOMARSprzedMies", OleDb.OleDbType.Double, Nothing, "LOMARSPRZED")

        dbCommandUpdate.Parameters.Add("@oNAJEMIloSprzedMies", OleDb.OleDbType.Double, Nothing, "NAJEMILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oNAJEMWartSprzedMies", OleDb.OleDbType.Double, Nothing, "NAJEMWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oNAJEMMARSprzedMies", OleDb.OleDbType.Double, Nothing, "NAJEMMARSPRZED")

        dbCommandUpdate.Parameters.Add("@oSTRONYIloSprzedMies", OleDb.OleDbType.Double, Nothing, "STRONYILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oSTRONYWartSprzedMies", OleDb.OleDbType.Double, Nothing, "STRONYWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oSTRONYMARSprzedMies", OleDb.OleDbType.Double, Nothing, "STRONYMARSPRZED")

        dbCommandUpdate.Parameters.Add("@oDRUKARKIIloSprzedMies", OleDb.OleDbType.Double, Nothing, "DRUKARKIILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oDRUKARKIWartSprzedMies", OleDb.OleDbType.Double, Nothing, "DRUKARKIWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oDRUKARKIMARSprzedMies", OleDb.OleDbType.Double, Nothing, "DRUKARKIMARSPRZED")

        dbCommandUpdate.Parameters.Add("@oSKUPIloSprzedMies", OleDb.OleDbType.Double, Nothing, "SKUPILOSPRZED")
        dbCommandUpdate.Parameters.Add("@oSKUPWartSprzedMies", OleDb.OleDbType.Double, Nothing, "SKUPWARTSPRZED")
        dbCommandUpdate.Parameters.Add("@oSKUPMARSprzedMies", OleDb.OleDbType.Double, Nothing, "SKUPMARSPRZED")

        dbCommandUpdate.Parameters.Add("@oWersjaMies", OleDb.OleDbType.Integer, Nothing, "WERSJA")

        'parametry filtrowania
        dbParam = dbCommandUpdate.Parameters.Add("@orygSymbol", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
        dbParam.SourceVersion = DataRowVersion.Original
        dbParam = dbCommandUpdate.Parameters.Add("@orygMies", OleDb.OleDbType.Char, 10, "MIESIAC")
        dbParam.SourceVersion = DataRowVersion.Original
        dbParam = dbCommandUpdate.Parameters.Add("@orygWersja", OleDb.OleDbType.Integer, Nothing, "WERSJA")
        dbParam.SourceVersion = DataRowVersion.Original
        dbDataAdapter.UpdateCommand = dbCommandUpdate
    End Sub

    Public Sub DBInsertObrotyMies(ByRef dbCommandInsert As OleDb.OleDbCommand,
                               ByRef dbDataAdapter As OleDb.OleDbDataAdapter)
        Dim dbParam As New OleDb.OleDbParameter
        '------------------------------------------
        'komenda INSERT
        'dbCommandUpdate.Parameters.Add("@oIdentMies", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
        'dbCommandUpdate.Parameters.Add("@oMiesMies", OleDb.OleDbType.Char, 7, "MIESIAC")

        dbParam = dbCommandInsert.Parameters.Add("@oIdentMies", OleDb.OleDbType.Char, 6, "IDENTKLIENTA")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oMiesMies", OleDb.OleDbType.Char, 10, "MIESIAC")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@oAIloSprzedMies", OleDb.OleDbType.Double, Nothing, "AILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oAWartSprzedMies", OleDb.OleDbType.Double, Nothing, "AWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oAMARSprzedMies", OleDb.OleDbType.Double, Nothing, "AMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oLIloSprzedMies", OleDb.OleDbType.Double, Nothing, "LILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oLWartSprzedMies", OleDb.OleDbType.Double, Nothing, "LWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oLMarSprzedMies", OleDb.OleDbType.Double, Nothing, "LMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@oAOIloSprzedMies", OleDb.OleDbType.Double, Nothing, "AOILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oAOWartSprzedMies", OleDb.OleDbType.Double, Nothing, "AOWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oAOMARSprzedMies", OleDb.OleDbType.Double, Nothing, "AOMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oLOIloSprzedMies", OleDb.OleDbType.Double, Nothing, "LOILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oLOWartSprzedMies", OleDb.OleDbType.Double, Nothing, "LOWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oLOMARSprzedMies", OleDb.OleDbType.Double, Nothing, "LOMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@oNAJEMIloSprzedMies", OleDb.OleDbType.Double, Nothing, "NAJEMILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oNAJEMWartSprzedMies", OleDb.OleDbType.Double, Nothing, "NAJEMWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oNAJEMMARSprzedMies", OleDb.OleDbType.Double, Nothing, "NAJEMMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@oSTRONYIloSprzedMies", OleDb.OleDbType.Double, Nothing, "STRONYILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oSTRONYWartSprzedMies", OleDb.OleDbType.Double, Nothing, "STRONYWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oSTRONYMARSprzedMies", OleDb.OleDbType.Double, Nothing, "STRONYMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@oDRUKARKIIloSprzedMies", OleDb.OleDbType.Double, Nothing, "DRUKARKIILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oDRUKARKIWartSprzedMies", OleDb.OleDbType.Double, Nothing, "DRUKARKIWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oDRUKARKIMARSprzedMies", OleDb.OleDbType.Double, Nothing, "DRUKARKIMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@oSKUPIloSprzedMies", OleDb.OleDbType.Double, Nothing, "SKUPILOSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oSKUPWartSprzedMies", OleDb.OleDbType.Double, Nothing, "SKUPWARTSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current
        dbParam = dbCommandInsert.Parameters.Add("@oSKUPMARSprzedMies", OleDb.OleDbType.Double, Nothing, "SKUPMARSPRZED")
        dbParam.SourceVersion = DataRowVersion.Current

        dbParam = dbCommandInsert.Parameters.Add("@oWersjaMies", OleDb.OleDbType.Integer, Nothing, "WERSJA")
        dbParam.SourceVersion = DataRowVersion.Current

        dbDataAdapter.InsertCommand = dbCommandInsert
    End Sub



    'Public oIdentMies As String            'IDENTKLIENTA     Text, 6
    'Public oMiesiacMies As String          'MIESIAC          Text,7

    'Public oLWartSprzedMies As Double      'LWARTSPRZED      Double
    'Public oLIloSprzedMies As Double       'LILOSPRZED       Double
    'Public oLMARSprzedMies As Double       'LMARSPRZED       Double

    'Public oAWartSprzedMies As Double      'AWARTSPRZED      Double
    'Public oAIloSprzedMies As Double       'AILOSPRZED       Double
    'Public oAMARSprzedMies As Double       'AMARSPRZED       Double

    'Public oLOWartSprzedMies As Double      'LOWARTSPRZED      Double
    'Public oLOIloSprzedMies As Double       'LOILOSPRZED       Double
    'Public oLOMARSprzedMies As Double       'LOMARSPRZED       Double

    'Public oAOWartSprzedMies As Double      'AOWARTSPRZED      Double
    'Public oAOIloSprzedMies As Double       'AOILOSPRZED       Double
    'Public oAOMARSprzedMies As Double       'AOMARSPRZED       Double

    'Public oNAJEMWartSprzedMies As Double      'NAJEMWARTSPRZED      Double
    'Public oNAJEMIloSprzedMies As Double       'NAJEMILOPRZED        Double
    'Public oNAJEMMARSprzedMies As Double       'NAJEMMARPRZED        Double

    'Public oSTRONYWartSprzedMies As Double      'STRONYWARTSPRZED      Double
    'Public oSTRONYIloSprzedMies As Double       'STRONYILOPRZED        Double
    'Public oSTRONYMARSprzedMies As Double       'STRONYMARPRZED        Double

    'Public oDRUKARKIWartSprzedMies As Double      'DRUKARKIWARTSPRZED      Double
    'Public oDRUKARKIIloSprzedMies As Double       'DRUKARKIILOPRZED        Double
    'Public oDRUKARKIMARSprzedMies As Double       'DRUKARKIMARPRZED        Double

    'Public oSKUPWartSprzedMies As Double      'SKUPWARTSPRZED      Double
    'Public oSKUPIloSprzedMies As Double       'SKUPILOPRZED        Double
    'Public oSKUPMARSprzedMies As Double       'SKUPMARPRZED        Double

    'Public oWersjaMies As Integer          'WERSJA           Integer

    Public Function IdentObrotyMies(ByVal IdentK As String, ByVal IdentM As String) As Boolean
        Dim dbSelectSQLObrotyMies As String = sSelectSQLObrotyMies & " WHERE IDENTKLIENTA='" & IdentK & "' AND MIESIAC='" & IdentM & "' "

        Dim dbConnectionObrotyMies As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectObrotyMies As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterObrotyMies As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionObrotyMies As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectObrotyMies As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterObrotyMies As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetObrotyMies As New DataSet
        Dim DataViewObrotyMies As New DataView
        Dim ConnectionStateObrotyMies As System.Data.ConnectionState

        Dim Wynik As Boolean = False

        oIdentMies = ""
        oMiesiacMies = ""
        oLWartSprzedMies = 0
        oLIloSprzedMies = 0
        oLMARSprzedMies = 0
        oAWartSprzedMies = 0
        oAIloSprzedMies = 0
        oAMARSprzedMies = 0
        oLOWartSprzedMies = 0
        oLOIloSprzedMies = 0
        oLOMARSprzedMies = 0
        oAOWartSprzedMies = 0
        oAOIloSprzedMies = 0
        oAOMARSprzedMies = 0

        oNAJEMWartSprzedMies = 0
        oNAJEMIloSprzedMies = 0
        oNAJEMMARSprzedMies = 0

        oSTRONYWartSprzedMies = 0
        oSTRONYIloSprzedMies = 0
        oSTRONYMARSprzedMies = 0

        oDRUKARKIWartSprzedMies = 0
        oDRUKARKIIloSprzedMies = 0
        oDRUKARKIMARSprzedMies = 0

        oSKUPWartSprzedMies = 0
        oSKUPIloSprzedMies = 0
        oSKUPMARSprzedMies = 0

        oWersjaMies = 0

        DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionObrotyMies = New OleDb.OleDbConnection(sConnectionKontakty)
            dbCommandSelectObrotyMies = New OleDb.OleDbCommand(dbSelectSQLObrotyMies, dbConnectionObrotyMies)
            dbDataAdapterObrotyMies = New OleDb.OleDbDataAdapter(dbCommandSelectObrotyMies)

            'mapujemy domyslna nazwe tabeli
            Dim dbMapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            dbMapowanieTabeliObrotyMies = dbDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")
            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionObrotyMies.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObrotyMies = dbConnectionObrotyMies.State
            End Try
        Else
            sqlConnectionObrotyMies = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectObrotyMies = New SqlClient.SqlCommand(dbSelectSQLObrotyMies, sqlConnectionObrotyMies)
            sqlDataAdapterObrotyMies = New SqlClient.SqlDataAdapter(sqlCommandSelectObrotyMies)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliObrotyMies = sqlDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObrotyMies.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObrotyMies = sqlConnectionObrotyMies.State
            End Try
        End If

        If ConnectionStateObrotyMies = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    'dbDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    'dbConnectionObrotyMies.Close()
                Else
                    sqlDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    sqlDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    sqlConnectionObrotyMies.Close()
                End If

                'definiuj DataView
                DataViewObrotyMies = New DataView(DataSetObrotyMies.Tables(0))
                If DataViewObrotyMies.Count > 0 Then
                    oIdentMies = GetTxtDRVField(DataViewObrotyMies.Item(0), "IDENTKLIENTA")
                    oMiesiacMies = GetTxtDRVField(DataViewObrotyMies.Item(0), "MIESIAC")
                    oLWartSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "LWARTSPRZED")
                    oLIloSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "LILOSPRZED")
                    oLMARSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "LMARSPRZED")
                    oAWartSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "AWARTSPRZED")
                    oAIloSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "AILOSPRZED")
                    oAMARSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "AMARSPRZED")
                    oLOWartSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "LOWARTSPRZED")
                    oLOIloSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "LOILOSPRZED")
                    oLOMARSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "LOMARSPRZED")
                    oAOWartSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "AOWARTSPRZED")
                    oAOIloSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "AOILOSPRZED")
                    oAOMARSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "AOMARSPRZED")

                    oNAJEMWartSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "NAJEMWARTSPRZED")
                    oNAJEMIloSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "NAJEMILOSPRZED")
                    oNAJEMMARSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "NAJEMMARSPRZED")

                    oSTRONYWartSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "STRONYWARTSPRZED")
                    oSTRONYIloSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "STRONYILOSPRZED")
                    oSTRONYMARSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "STRONYMARSPRZED")

                    oDRUKARKIWartSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "DRUKARKIWARTSPRZED")
                    oDRUKARKIIloSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "DRUKARKIILOSPRZED")
                    oDRUKARKIMARSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "DRUKARKIMARSPRZED")

                    oSKUPWartSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "SKUPWARTSPRZED")
                    oSKUPIloSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "SKUPILOSPRZED")
                    oSKUPMARSprzedMies = GetDblDRVField(DataViewObrotyMies.Item(0), "SKUPMARSPRZED")

                    oWersjaMies = GetIntDRVField(DataViewObrotyMies.Item(0), "WERSJA")

                    Wynik = True
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wynik)
    End Function




    Public Function WartoscObrotowMies(ByVal IdentK As String,
                                       ByVal IdentM As String) As Double
        Dim dbSelectSQLObrotyMies As String = sSelectSQLObrotyMies & " WHERE IDENTKLIENTA='" & IdentK & "' AND MIESIAC='" & IdentM & "' "

        Dim dbConnectionObrotyMies As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectObrotyMies As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterObrotyMies As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionObrotyMies As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectObrotyMies As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterObrotyMies As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetObrotyMies As New DataSet
        Dim DataViewObrotyMies As New DataView
        Dim ConnectionStateObrotyMies As System.Data.ConnectionState
        Dim i As Integer = 0
        Dim Wartosc As Double = 0

        DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionObrotyMies = New OleDb.OleDbConnection(sConnectionKontakty)
            dbCommandSelectObrotyMies = New OleDb.OleDbCommand(dbSelectSQLObrotyMies, dbConnectionObrotyMies)
            dbDataAdapterObrotyMies = New OleDb.OleDbDataAdapter(dbCommandSelectObrotyMies)

            'mapujemy domyslna nazwe tabeli
            Dim dbMapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            dbMapowanieTabeliObrotyMies = dbDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")
            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionObrotyMies.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObrotyMies = dbConnectionObrotyMies.State
            End Try
        Else
            sqlConnectionObrotyMies = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectObrotyMies = New SqlClient.SqlCommand(dbSelectSQLObrotyMies, sqlConnectionObrotyMies)
            sqlDataAdapterObrotyMies = New SqlClient.SqlDataAdapter(sqlCommandSelectObrotyMies)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliObrotyMies = sqlDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObrotyMies.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObrotyMies = sqlConnectionObrotyMies.State
            End Try
        End If

        If ConnectionStateObrotyMies = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    'dbDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    'dbConnectionObrotyMies.Close()
                Else
                    sqlDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    sqlDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    sqlConnectionObrotyMies.Close()
                End If

                'definiuj DataView
                DataViewObrotyMies = New DataView(DataSetObrotyMies.Tables(0))
                If DataViewObrotyMies.Count > 0 Then
                    For i = 0 To DataViewObrotyMies.Count - 1
                        Wartosc += GetDblDRVField(DataViewObrotyMies.Item(i), "LWARTSPRZED") +
                                   GetDblDRVField(DataViewObrotyMies.Item(i), "AWARTSPRZED") +
                                   GetDblDRVField(DataViewObrotyMies.Item(i), "LOWARTSPRZED") +
                                   GetDblDRVField(DataViewObrotyMies.Item(i), "AOWARTSPRZED") +
                                   GetDblDRVField(DataViewObrotyMies.Item(i), "NAJEMWARTSPRZED") +
                                   GetDblDRVField(DataViewObrotyMies.Item(i), "STRONYWARTSPRZED") +
                                   GetDblDRVField(DataViewObrotyMies.Item(i), "DRUKARKIWARTSPRZED") +
                                   GetDblDRVField(DataViewObrotyMies.Item(i), "SKUPWARTSPRZED")
                    Next i
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wartosc)
    End Function




    Public Function WartoscObrotowMiesMaterialow(ByVal IdentK As String,
                                       ByVal IdentM As String) As Double
        Dim dbSelectSQLObrotyMies As String = sSelectSQLObrotyMies & " WHERE IDENTKLIENTA='" & IdentK & "' AND MIESIAC='" & IdentM & "' "

        Dim dbConnectionObrotyMies As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectObrotyMies As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterObrotyMies As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionObrotyMies As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectObrotyMies As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterObrotyMies As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetObrotyMies As New DataSet
        Dim DataViewObrotyMies As New DataView
        Dim ConnectionStateObrotyMies As System.Data.ConnectionState
        Dim i As Integer = 0
        Dim Wartosc As Double = 0

        DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionObrotyMies = New OleDb.OleDbConnection(sConnectionKontakty)
            dbCommandSelectObrotyMies = New OleDb.OleDbCommand(dbSelectSQLObrotyMies, dbConnectionObrotyMies)
            dbDataAdapterObrotyMies = New OleDb.OleDbDataAdapter(dbCommandSelectObrotyMies)

            'mapujemy domyslna nazwe tabeli
            Dim dbMapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            dbMapowanieTabeliObrotyMies = dbDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")
            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionObrotyMies.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObrotyMies = dbConnectionObrotyMies.State
            End Try
        Else
            sqlConnectionObrotyMies = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectObrotyMies = New SqlClient.SqlCommand(dbSelectSQLObrotyMies, sqlConnectionObrotyMies)
            sqlDataAdapterObrotyMies = New SqlClient.SqlDataAdapter(sqlCommandSelectObrotyMies)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliObrotyMies = sqlDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObrotyMies.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObrotyMies = sqlConnectionObrotyMies.State
            End Try
        End If

        If ConnectionStateObrotyMies = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    'dbDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    'dbConnectionObrotyMies.Close()
                Else
                    sqlDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    sqlDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    sqlConnectionObrotyMies.Close()
                End If

                'definiuj DataView
                DataViewObrotyMies = New DataView(DataSetObrotyMies.Tables(0))
                If DataViewObrotyMies.Count > 0 Then
                    For i = 0 To DataViewObrotyMies.Count - 1
                        Wartosc += GetDblDRVField(DataViewObrotyMies.Item(i), "LWARTSPRZED") +
                                   GetDblDRVField(DataViewObrotyMies.Item(i), "AWARTSPRZED") +
                                   GetDblDRVField(DataViewObrotyMies.Item(i), "LOWARTSPRZED") +
                                   GetDblDRVField(DataViewObrotyMies.Item(i), "AOWARTSPRZED")
                    Next i
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wartosc)
    End Function



    Public Function WartoscObrotowMiesUslug(ByVal IdentK As String,
                                       ByVal IdentM As String) As Double
        Dim dbSelectSQLObrotyMies As String = sSelectSQLObrotyMies & " WHERE IDENTKLIENTA='" & IdentK & "' AND MIESIAC='" & IdentM & "' "

        Dim dbConnectionObrotyMies As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectObrotyMies As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterObrotyMies As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionObrotyMies As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectObrotyMies As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterObrotyMies As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetObrotyMies As New DataSet
        Dim DataViewObrotyMies As New DataView
        Dim ConnectionStateObrotyMies As System.Data.ConnectionState
        Dim i As Integer = 0
        Dim Wartosc As Double = 0

        DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionObrotyMies = New OleDb.OleDbConnection(sConnectionKontakty)
            dbCommandSelectObrotyMies = New OleDb.OleDbCommand(dbSelectSQLObrotyMies, dbConnectionObrotyMies)
            dbDataAdapterObrotyMies = New OleDb.OleDbDataAdapter(dbCommandSelectObrotyMies)

            'mapujemy domyslna nazwe tabeli
            Dim dbMapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            dbMapowanieTabeliObrotyMies = dbDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")
            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionObrotyMies.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObrotyMies = dbConnectionObrotyMies.State
            End Try
        Else
            sqlConnectionObrotyMies = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectObrotyMies = New SqlClient.SqlCommand(dbSelectSQLObrotyMies, sqlConnectionObrotyMies)
            sqlDataAdapterObrotyMies = New SqlClient.SqlDataAdapter(sqlCommandSelectObrotyMies)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliObrotyMies = sqlDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObrotyMies.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObrotyMies = sqlConnectionObrotyMies.State
            End Try
        End If

        If ConnectionStateObrotyMies = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    'dbDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    'dbConnectionObrotyMies.Close()
                Else
                    sqlDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    sqlDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    sqlConnectionObrotyMies.Close()
                End If

                'definiuj DataView
                DataViewObrotyMies = New DataView(DataSetObrotyMies.Tables(0))
                If DataViewObrotyMies.Count > 0 Then
                    For i = 0 To DataViewObrotyMies.Count - 1
                        Wartosc += GetDblDRVField(DataViewObrotyMies.Item(i), "NAJEMWARTSPRZED") +
                                   GetDblDRVField(DataViewObrotyMies.Item(i), "STRONYWARTSPRZED") +
                                   GetDblDRVField(DataViewObrotyMies.Item(i), "DRUKARKIWARTSPRZED")             '+ GetDblDRVField(DataViewObrotyMies.Item(i), "SKUPWARTSPRZED")
                    Next i
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wartosc)
    End Function





    Public Function WartoscObrotowMiesSkup(ByVal IdentK As String,
                                       ByVal IdentM As String) As Double
        Dim dbSelectSQLObrotyMies As String = sSelectSQLObrotyMies & " WHERE IDENTKLIENTA='" & IdentK & "' AND MIESIAC='" & IdentM & "' "

        Dim dbConnectionObrotyMies As OleDb.OleDbConnection = Nothing
        Dim dbCommandSelectObrotyMies As OleDb.OleDbCommand = Nothing
        Dim dbDataAdapterObrotyMies As OleDb.OleDbDataAdapter = Nothing

        Dim sqlConnectionObrotyMies As SqlClient.SqlConnection = Nothing
        Dim sqlCommandSelectObrotyMies As SqlClient.SqlCommand = Nothing
        Dim sqlDataAdapterObrotyMies As SqlClient.SqlDataAdapter = Nothing

        Dim DataSetObrotyMies As New DataSet
        Dim DataViewObrotyMies As New DataView
        Dim ConnectionStateObrotyMies As System.Data.ConnectionState
        Dim i As Integer = 0
        Dim Wartosc As Double = 0

        DataSetObrotyMies.Locale = New System.Globalization.CultureInfo("pl-PL")
        If ParL_DataBaseType = "MS ACCESS" Then
            dbConnectionObrotyMies = New OleDb.OleDbConnection(sConnectionKontakty)
            dbCommandSelectObrotyMies = New OleDb.OleDbCommand(dbSelectSQLObrotyMies, dbConnectionObrotyMies)
            dbDataAdapterObrotyMies = New OleDb.OleDbDataAdapter(dbCommandSelectObrotyMies)

            'mapujemy domyslna nazwe tabeli
            Dim dbMapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            dbMapowanieTabeliObrotyMies = dbDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")
            '------------------------------------------
            'wypelnij DATASET
            Try
                dbConnectionObrotyMies.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObrotyMies = dbConnectionObrotyMies.State
            End Try
        Else
            sqlConnectionObrotyMies = New SqlClient.SqlConnection(sConnectionKontakty)
            sqlCommandSelectObrotyMies = New SqlClient.SqlCommand(dbSelectSQLObrotyMies, sqlConnectionObrotyMies)
            sqlDataAdapterObrotyMies = New SqlClient.SqlDataAdapter(sqlCommandSelectObrotyMies)

            'mapujemy domyslna nazwe tabeli
            Dim sqlMapowanieTabeliObrotyMies As System.Data.Common.DataTableMapping
            sqlMapowanieTabeliObrotyMies = sqlDataAdapterObrotyMies.TableMappings.Add("Table", "TABELA_ObrotyMies")
            '------------------------------------------
            'wypelnij DATASET
            Try
                sqlConnectionObrotyMies.Open()
            Catch Ex As System.Exception
            Finally
                ConnectionStateObrotyMies = sqlConnectionObrotyMies.State
            End Try
        End If

        If ConnectionStateObrotyMies = ConnectionState.Open Then
            Try
                If ParL_DataBaseType = "MS ACCESS" Then
                    'dbDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    'dbDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    'dbConnectionObrotyMies.Close()
                Else
                    sqlDataAdapterObrotyMies.FillSchema(DataSetObrotyMies, SchemaType.Mapped, "TABELA_ObrotyMies")
                    sqlDataAdapterObrotyMies.Fill(DataSetObrotyMies)
                    sqlConnectionObrotyMies.Close()
                End If

                'definiuj DataView
                DataViewObrotyMies = New DataView(DataSetObrotyMies.Tables(0))
                If DataViewObrotyMies.Count > 0 Then
                    For i = 0 To DataViewObrotyMies.Count - 1
                        Wartosc += GetDblDRVField(DataViewObrotyMies.Item(i), "SKUPWARTSPRZED")
                    Next i
                End If
            Catch Ex As System.Exception
            End Try
        End If
        Return (Wartosc)
    End Function














End Module
