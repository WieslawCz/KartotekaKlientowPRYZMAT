Public Class frmAnalizaCzestKontaktow

    Private Sub frmAnalizaCzestKontaktow_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.PictureBox1.Image = My.Resources.logomini
        'Me.TopMost = True
        If Len(_IdentKlienta) > 0 Then
            IdentKlienta(_IdentKlienta)

            lblIdent.Text = _IdentKlienta
            lblNazwaHandlowa.Text = oNazwa1Kli
            lblKod.Text = oKodPoczKli
            lblMiejscowosc.Text = oMiejscKli
            lblUlica.Text = oUlicaKli
            lblNrDomu.Text = oNrDomuKli
            lblIDDomu.Text = oIDDomuKli
        Else
            lblIdent.Text = ""
            lblNazwaHandlowa.Text = ""
            lblKod.Text = ""
            lblMiejscowosc.Text = ""
            lblUlica.Text = ""
            lblNrDomu.Text = ""
            lblIDDomu.Text = ""
        End If
        '----------------------
        'Zaloz robocze struktury...
        Dim robDataSet As New DataSet()
        Dim robTable As New DataTable()
        Dim robDataView As DataView = Nothing
        Dim robRow As DataRow = Nothing
        Dim robRowView As DataRowView = Nothing

        Dim robCol1 As DataColumn = robTable.Columns.Add("RODZAJ", GetType(System.String))
        Dim robCol2 As DataColumn = robTable.Columns.Add("RAZEM", GetType(System.Int32))
        Dim robCol3 As DataColumn = robTable.Columns.Add("T1", GetType(System.Int32))
        Dim robCol4 As DataColumn = robTable.Columns.Add("T2", GetType(System.Int32))
        Dim robCol5 As DataColumn = robTable.Columns.Add("M1", GetType(System.Int32))
        Dim robCol6 As DataColumn = robTable.Columns.Add("M2", GetType(System.Int32))
        Dim robCol7 As DataColumn = robTable.Columns.Add("M3", GetType(System.Int32))
        Dim robCol8 As DataColumn = robTable.Columns.Add("M4", GetType(System.Int32))
        Dim robCol9 As DataColumn = robTable.Columns.Add("M5", GetType(System.Int32))
        Dim robColA As DataColumn = robTable.Columns.Add("M6", GetType(System.Int32))
        Dim robColB As DataColumn = robTable.Columns.Add("R1", GetType(System.Int32))
        Dim robColC As DataColumn = robTable.Columns.Add("POZOSTALE", GetType(System.Int32))
        robDataSet.Tables.Add(robTable)
        'definiuj DataView
        robDataView = New DataView(robDataSet.Tables(0))
        robDataView.Sort = "RODZAJ"
        '-------------------------------------------
        Dim DataBiez As String = SysData()
        Dim DataT1 As String = WyliczDate(DataBiez, -7)
        Dim DataT2 As String = WyliczDate(DataBiez, -14)
        Dim DataM1 As String = PopMiesiac(DataBiez)
        Dim DataM2 As String = PopMiesiac(DataM1)
        Dim DataM3 As String = PopMiesiac(DataM2)
        Dim DataM4 As String = PopMiesiac(DataM3)
        Dim DataM5 As String = PopMiesiac(DataM4)
        Dim DataM6 As String = PopMiesiac(DataM5)
        Dim DataR1 As String = PopRok(DataBiez)

        ''---------------------------------------------------------------------
        ''Kontakty HANDLOWE - nowe 05.09.2015
        ''-----------------------------------------
        'Public oUniqIdKon As String           'UNIQID            varchar(40)
        'Public oIdentKon As String            'IDENTKLIENTA     Text, 6
        'Public oDataKon As String             'DATAKONTAKTU     Text, 10    data rrrr-mm-dd
        'Public oPracownikKon As String        'PRACOWNIK        Text, 10     uzytkownik
        'Public oTematKon As String            'TEMAT            Text, 50
        'Public oUwagiKon As String            'UWAGI            Memo
        'Public oWersjaKon As Integer          'WERSJA           Integer


        Dim drv As DataRowView
        Dim rowNo As Integer = 0
        For Each drv In _DataView
            oIdentKon = GetTxtDRVField(drv, "IDENTKLIENTA")
            oDataKon = GetTxtDRVField(drv, "DATAKONTAKTU")
            oPracownikKon = GetTxtDRVField(drv, "PRACOWNIK")
            oTematKon = GetTxtDRVField(drv, "TEMAT")
            oUwagiKon = GetTxtDRVField(drv, "UWAGI")

            If (Len(_IdentKlienta) = 0) Or (oIdentKon = _IdentKlienta) Then
                rowNo = robDataView.Find(oTematKon)
                If rowNo < 0 Then
                    ' wypelnij dataset informacjami o programie...
                    robRow = robDataSet.Tables(0).NewRow()
                    robRow("RODZAJ") = oTematKon
                    robRow("RAZEM") = 0
                    robRow("T1") = 0
                    robRow("T2") = 0
                    robRow("M1") = 0
                    robRow("M2") = 0
                    robRow("M3") = 0
                    robRow("M4") = 0
                    robRow("M5") = 0
                    robRow("M6") = 0
                    robRow("R1") = 0
                    robRow("POZOSTALE") = 0
                    robDataSet.Tables(0).Rows.Add(robRow)

                    rowNo = robDataView.Find(oTematKon)
                End If
                robRowView = robDataView.Item(rowNo)

                If oDataKon < DataR1 Then
                    robRowView("RAZEM") += 1
                    robRowView("POZOSTALE") += 1
                ElseIf oDataKon < DataM6 Then
                    robRowView("RAZEM") += 1
                    robRowView("R1") += 1
                ElseIf oDataKon < DataM5 Then
                    robRowView("RAZEM") += 1
                    robRowView("M6") += 1
                ElseIf oDataKon < DataM4 Then
                    robRowView("RAZEM") += 1
                    robRowView("M5") += 1
                ElseIf oDataKon < DataM3 Then
                    robRowView("RAZEM") += 1
                    robRowView("M4") += 1
                ElseIf oDataKon < DataM2 Then
                    robRowView("RAZEM") += 1
                    robRowView("M3") += 1
                ElseIf oDataKon < DataM1 Then
                    robRowView("RAZEM") += 1
                    robRowView("M2") += 1
                ElseIf oDataKon < DataT2 Then
                    robRowView("RAZEM") += 1
                    robRowView("M1") += 1
                ElseIf oDataKon < DataT1 Then
                    robRowView("RAZEM") += 1
                    robRowView("T2") += 1
                Else
                    robRowView("RAZEM") += 1
                    robRowView("T1") += 1
                End If

            End If
        Next
        '-----------------------------
        'wyswietl tabele wynikow
        'parametry DataGrid
        dagRaport.CaptionVisible = False
        dagRaport.ColumnHeadersVisible = True
        dagRaport.RowHeadersVisible = True
        dagRaport.BackgroundColor = System.Drawing.SystemColors.Control
        dagRaport.BorderStyle = BorderStyle.Fixed3D
        dagRaport.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top _
                                Or System.Windows.Forms.AnchorStyles.Bottom) _
                                Or System.Windows.Forms.AnchorStyles.Left) _
                                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        dagRaport.CaptionFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        dagRaport.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))

        'wypelnienie DataGrid
        'dagRaport.SetDataBinding(DataSetRaport, "TABELA_Raport")
        robDataView.AllowDelete = False
        robDataView.AllowNew = False
        dagRaport.DataSource = robDataView

        ' Add a GridTableStyle and set the MappingName to the name of the DataTable.
        Dim TSRaport As New DataGridTableStyle
        TSRaport.MappingName = robDataSet.Tables(0).TableName
        TSRaport.AlternatingBackColor = System.Drawing.SystemColors.Control
        TSRaport.BackColor = System.Drawing.SystemColors.ControlLight
        TSRaport.GridLineColor = System.Drawing.SystemColors.ControlDark
        TSRaport.ForeColor = System.Drawing.Color.Purple
        TSRaport.HeaderBackColor = System.Drawing.SystemColors.ScrollBar
        TSRaport.HeaderForeColor = System.Drawing.Color.Navy
        TSRaport.SelectionBackColor = SoftartCursorBackColor   'System.Drawing.Color.LightSeaGreen
        TSRaport.SelectionForeColor = SoftartCursorForeColor   'System.Drawing.Color.Yellow

        ''definiuj delegat
        'Dim deleg As delegateDefColor = AddressOf Koloruj

        TxtColumn(TSRaport, robDataSet.Tables(0).Columns(0).ColumnName, "Rodzaj kontaktu", 200, HorizontalAlignment.Left)
        NumColumn(TSRaport, robDataSet.Tables(0).Columns(1).ColumnName, "Razem  ", 50, HorizontalAlignment.Right, "F0")
        NumColumn(TSRaport, robDataSet.Tables(0).Columns(2).ColumnName, "1 Tydz. ", 50, HorizontalAlignment.Right, "F0")
        NumColumn(TSRaport, robDataSet.Tables(0).Columns(3).ColumnName, "2 Tydz. ", 50, HorizontalAlignment.Right, "F0")
        NumColumn(TSRaport, robDataSet.Tables(0).Columns(4).ColumnName, "1 Mies. ", 50, HorizontalAlignment.Right, "F0")
        NumColumn(TSRaport, robDataSet.Tables(0).Columns(5).ColumnName, "2 Mies. ", 50, HorizontalAlignment.Right, "F0")
        NumColumn(TSRaport, robDataSet.Tables(0).Columns(6).ColumnName, "3 Mies. ", 50, HorizontalAlignment.Right, "F0")
        NumColumn(TSRaport, robDataSet.Tables(0).Columns(7).ColumnName, "4 Mies. ", 50, HorizontalAlignment.Right, "F0")
        NumColumn(TSRaport, robDataSet.Tables(0).Columns(8).ColumnName, "5 Mies. ", 50, HorizontalAlignment.Right, "F0")
        NumColumn(TSRaport, robDataSet.Tables(0).Columns(9).ColumnName, "6 Mies. ", 50, HorizontalAlignment.Right, "F0")
        NumColumn(TSRaport, robDataSet.Tables(0).Columns(10).ColumnName, "1 Rok  ", 50, HorizontalAlignment.Right, "F0")
        NumColumn(TSRaport, robDataSet.Tables(0).Columns(11).ColumnName, "Pozostałe  ", 50, HorizontalAlignment.Right, "F0")

        ' Add the DataGridTableStyle instance to the GridTableStylesCollection. 
        dagRaport.TableStyles.Clear()
        dagRaport.TableStyles.Add(TSRaport)



    End Sub

    Private Sub cmdStop_Click(sender As Object, e As EventArgs) Handles cmdStop.Click
        Me.Close()
    End Sub

    Private Sub frmAnalizaCzestKontaktow_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        RysujGradient(Me)
    End Sub
End Class