'------------------------------------------
'UWAGA :
'Do referencji trzeba DODAC
'   Microsoft Excel 11.0 Object Library
'   Microsoft Office 11.0 Object Library
'-------------------------------------------
Imports System.Reflection ' For Missing.Value and BindingFlags
Imports System.Runtime.InteropServices ' For COMException
Imports Microsoft.Office.Interop.Excel

Module _modExcel
    '------------------------------------------
    'UWAGA :
    'Do referencji trzeba DODAC
    '   Microsoft Excel 11.0 Object Library
    '   Microsoft Office 11.0 Object Library
    '-------------------------------------------
    ' Office version    Office No   Framework       PIA
    ' Office XP (2002)   10
    ' Office 2003        11         Framework 1.1   O2003PIA.EXE
    ' Office 2007        12         Framework 1.1   O2007PIA.EXE
    ' Office 2010        14         Framework 2.0   O2010PIA.EXE

    'weryfikacja wersji PIA
    'http://www.peterlee.com.cn/blog.php?getArticleID=116&getYearMonth=20112





    ''-------------odbc
    ''Private EXCELConnection As String = "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & oPlikNaDysku & ";"
    'Private EXCELConnectionODBC As String = _
    '            "Driver={Microsoft Excel Driver (*.xls)};" & _
    '            "DSN=Pliki programu Excel;" & _
    '            "DBQ=" & oPlikNaDysku & ";" & _
    '            "DriverId=790;" & _
    '            "MaxBufferSize=2048;" & _
    '            "PageTimeout=5;"

    ''-------------oleDB
    'Private EXCELConnectionOLeDB As String = "provider=Microsoft.Jet.OLEDB.4.0; " & _
    '                                "data source=" & oPlikNaDysku & "; " & _
    '                                "Extended Properties='Excel 8.0; HDR=YES; IMEX=1';"

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




    '===================================
    ' Funkcje narzędziowe
    '===================================

    Private Function AdrColDlaExcela(ByVal colNo As Integer) As String
        Dim IdWierszy As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim colId As String = ""
        If colNo < Len(IdWierszy) Then
            colId = Mid(IdWierszy, colNo + 1, 1)
        Else
            colId = Mid(IdWierszy, Int(colNo / Len(IdWierszy)), 1) & Mid(IdWierszy, (colNo Mod Len(IdWierszy)) + 1, 1)
        End If
        Return (colId)
    End Function


    '===================================
    ' zapisz nagłówek tabeli
    '===================================

    Private Sub WpiszTytulDoArkusza(ByVal _Arkusz As _Worksheet, _
                           ByVal _Adres As String, _
                           ByVal _Zawartosc As String)

        Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
        'Dim ColRange As Microsoft.Office.Interop.Excel.Range
        'Dim RowRange As Microsoft.Office.Interop.Excel.Range

        range = _Arkusz.Range(_Adres, Missing.Value)
        If range Is Nothing Then
            MessageBox.Show("Nie mogę zdefiniować zakresu arkusza EXCEL", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            range.RowHeight = 15
            range.WrapText = False
            range.Font.Bold = True
            range.Value2 = _Zawartosc
        End If
    End Sub


    '===================================
    ' zapisz nagłówek tabeli
    '===================================

    Private Sub WpiszNaglowekDoArkusza(ByVal _Arkusz As _Worksheet, _
                           ByVal _Adres As String, _
                           ByVal _Zawartosc As String, _
                           ByVal _Szerokosc As Integer, _
                           ByVal _Wyrownanie As Integer)
        Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
        'Dim ColRange As Microsoft.Office.Interop.Excel.Range
        'Dim RowRange As Microsoft.Office.Interop.Excel.Range

        range = _Arkusz.Range(_Adres, Missing.Value)
        If range Is Nothing Then
            MessageBox.Show("Nie mogę zdefiniować zakresu arkusza EXCEL", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            'range.Cells.NumberFormat = "########0.0000"
            range.ColumnWidth = (48 / 64) * _Szerokosc    '48 Points = 64 Pixels
            range.RowHeight = 15
            range.WrapText = True
            If Not _Wyrownanie = Nothing Then
                range.HorizontalAlignment = _Wyrownanie
            End If
            range.Font.Bold = True
            range.Font.Color = System.Drawing.Color.Black.ToArgb
            range.Value2 = _Zawartosc
        End If
    End Sub



    Private Sub WpiszNaglowekDoArkusza(ByVal _Arkusz As _Worksheet, _
                       ByVal _Adres As String, _
                       ByVal _Zawartosc As String, _
                       ByVal _Szerokosc As Integer, _
                       ByVal _Wyrownanie As Integer, _
                       ByVal _Format As String)
        Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
        'Dim ColRange As Microsoft.Office.Interop.Excel.Range
        'Dim RowRange As Microsoft.Office.Interop.Excel.Range

        range = _Arkusz.Range(_Adres, Missing.Value)
        If range Is Nothing Then
            MessageBox.Show("Nie mogę zdefiniować zakresu arkusza EXCEL", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            If Len(_Format) > 0 Then
                'range.Cells.NumberFormat = _Format
            End If
            range.ColumnWidth = (48 / 64) * _Szerokosc    '48 Points = 64 Pixels
            range.RowHeight = 15
            range.WrapText = True
            If Not _Wyrownanie = Nothing Then
                range.HorizontalAlignment = _Wyrownanie
            End If
            range.Font.Bold = True
            range.Font.Color = System.Drawing.Color.Black.ToArgb
            range.Value2 = _Zawartosc
        End If
    End Sub


    '===================================
    ' zapisz pozycje tabeli
    '===================================

    Private Sub WpiszDoArkusza(ByVal _Arkusz As _Worksheet, _
                               ByVal _Adres As String, _
                               ByVal _Zawartosc As String, _
                               ByVal _Szerokosc As Integer, _
                               ByVal _Wyrownanie As Integer)
        Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
        'Dim ColRange As Microsoft.Office.Interop.Excel.Range
        'Dim RowRange As Microsoft.Office.Interop.Excel.Range

        range = _Arkusz.Range(_Adres, Missing.Value)
        If range Is Nothing Then
            MessageBox.Show("Nie mogę zdefiniować zakresu arkusza EXCEL", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            Try
                'range.Cells.NumberFormat = "@"
                'range.Cells.NumberFormat = "########0.0000"
                range.ColumnWidth = (48 / 64) * _Szerokosc    '48 Points = 64 Pixels
                range.RowHeight = 15
                range.WrapText = True
                If Not _Wyrownanie = Nothing Then
                    range.HorizontalAlignment = _Wyrownanie
                End If
                range.Font.Bold = False
                range.Font.Color = System.Drawing.Color.Navy.ToArgb
                range.Value2 = _Zawartosc

            Catch ex As Exception
                MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & ex.Message, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            End Try
        End If
    End Sub


    Private Sub WpiszDoArkusza(ByVal _Arkusz As _Worksheet, _
                           ByVal _Adres As String, _
                           ByVal _Zawartosc As Object, _
                           ByVal _Szerokosc As Integer, _
                           ByVal _Wyrownanie As Integer, _
                           ByVal _Format As String)
        Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
        'Dim ColRange As Microsoft.Office.Interop.Excel.Range
        'Dim RowRange As Microsoft.Office.Interop.Excel.Range

        range = _Arkusz.Range(_Adres, Missing.Value)
        If range Is Nothing Then
            MessageBox.Show("Nie mogę zdefiniować zakresu arkusza EXCEL", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            Try
                'range.Cells.NumberFormat = "@"
                If Len(_Format) > 0 Then
                    'range.Cells.NumberFormat = _Format
                End If
                range.ColumnWidth = (48 / 64) * _Szerokosc    '48 Points = 64 Pixels
                range.RowHeight = 15
                range.WrapText = True
                If Not _Wyrownanie = Nothing Then
                    range.HorizontalAlignment = _Wyrownanie
                End If
                range.Font.Bold = False
                range.Font.Color = System.Drawing.Color.Navy.ToArgb
                range.Value2 = _Zawartosc

            Catch ex As Exception
                MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & ex.Message, "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
            End Try
        End If
    End Sub










    'Public Sub ExportDataGrid2Excel2(ByRef myDataGrid As DataGrid)
    '    'On Error GoTo er
    '    Try

    '        Dim cm As CurrencyManager = CType(myDataGrid.BindingContext(myDataGrid.DataSource), CurrencyManager)
    '        Dim myDataGridRowsNo As Integer = cm.Count
    '        'Dim myDataGridColumnsNo As Integer = CType(myDataGrid.DataSource, DataTable).Columns.Count

    '        Dim myDataGridColumnsNo As Integer = CType(myDataGrid.DataSource, DataView).ToTable.Columns.Count

    '        Dim oExcel As Object = CreateObject("Excel.Application")

    '        'Dim oWorkBook As Object
    '        Dim oWorkSheet As Object

    '        Dim i As Integer = 0
    '        Dim k As Integer = 0
    '        Dim lRow As Long = 0
    '        Dim LastRow As Long = 0
    '        Dim LastCol As Long = 0

    '        Dim ExcelRow As Integer = 0
    '        Dim ExcelCol As Integer = 0
    '        Dim ExcelRowOffset As Integer = 2
    '        Dim ExcelColOffset As Integer = 0

    '        oExcel.Visible = False
    '        oExcel.Workbooks.Open(oKatProgramu & "\Nirmal.xls")
    '        oWorkSheet = oExcel.Workbooks("Nirmal.xls").sheets("Batch")

    '        'zapamietaj lokalizacej kursora
    '        LastRow = myDataGrid.CurrentCell.RowNumber 'Save Current row
    '        LastCol = myDataGrid.CurrentCell.ColumnNumber 'and column
    '        '-------------------------------------
    '        Dim row As Integer
    '        Dim col As Integer
    '        Dim cellValue As Object = Nothing
    '        For row = 0 To myDataGridRowsNo - 1
    '            ExcelRow = row + ExcelRowOffset
    '            For col = 0 To myDataGridColumnsNo - 1
    '                cellValue = myDataGrid(row, col)
    '                ExcelCol = col + ExcelColOffset

    '                oWorkSheet.Cells(ExcelRow, ExcelCol).Font.Bold = False
    '                oWorkSheet.Cells(ExcelRow, ExcelCol).Font.Color = System.Drawing.Color.Navy
    '                oWorkSheet.Cells(ExcelRow, ExcelCol).value = cellValue.ToString
    '            Next col
    '        Next row
    '        '-------------------------------------
    '        myDataGrid.CurrentCell = New DataGridCell(LastRow, LastCol)




    '        oExcel.Workbooks("Nirmal.xls").Save()
    '        oExcel.Workbooks("Nirmal.xls").Close(savechanges:=True)
    '        oExcel.Quit()

    '    Catch ex As Exception
    '        MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & ex.Message, "Uwaga :", _
    '            System.Windows.Forms.MessageBoxButtons.OK, _
    '            MessageBoxIcon.Information, _
    '            MessageBoxDefaultButton.Button1)
    '    End Try

    '    'If Err.Number = 1004 Then
    '    '    Exit Sub
    '    'End If
    'End Sub


    '===================================
    ' przepisujemy dane z DataGrid do EXCELa
    '===================================

    Public Sub ExportDataGrid2Excel(ByRef myDataGrid As DataGridView, _
                                    ByRef myDataView As DataView, _
                                    ByVal myNaglowek As String)
        Dim Text As String = ""
        Dim Text2 As String = ""
        Dim Num As Integer = 0
        'Dim MyDataView As DataView = CType(myDataGrid.DataSource, DataView)

        Dim N As Object = Nothing
        Dim T As String = ""
        Dim F As String = ""
        Dim R As Integer = 0
        Dim H As String = ""
        Dim W As Integer = 0
        Dim I As Integer = 0
        '---------------------------------
        Dim app As Application
        app = New Application()
        If app Is Nothing Then
            MessageBox.Show("Nie mogę uruchomić program EXCEL", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            app.Visible = False      'EXCEL WIDOCZNY NA EKRANIE

            Dim workbooks As Workbooks      'Getting the workbooks collection
            workbooks = app.Workbooks

            Dim workbook As _Workbook       'adding a new workbook
            workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)

            Dim sheets As Sheets            'Getting the worksheets collection
            sheets = workbook.Worksheets

            Dim worksheet As _Worksheet     'adding a new worksheet
            worksheet = sheets.Item(1)
            If worksheet Is Nothing Then
                MessageBox.Show("Nie mogę dodać nowego arkusza EXCEL", "Uwaga :", _
                                 System.Windows.Forms.MessageBoxButtons.OK, _
                                 MessageBoxIcon.Information, _
                                 MessageBoxDefaultButton.Button1)
            Else
                worksheet.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnRight
                worksheet.Columns.ColumnWidth = 100

                Dim ir As Integer = 0
                Dim ic As Integer = 0
                Dim dr As DataRowView

                'TYTUŁ TABELI - wiersz nr 1
                WpiszTytulDoArkusza(worksheet, "A1", myNaglowek)

                'WIERSZ NAGŁÓWKOWY - wiersz nr 3
                Dim ie As Integer = 0
                For ic = 0 To myDataView.Table.Columns.Count - 1
                    If myDataGrid.Columns(ic).Visible Then
                        'N = myDataGrid.Columns(ic).DataPropertyName
                        N = myDataGrid.Columns(ic).HeaderText
                        'N = Trim(myDataView.Table.Columns.Item(ic).Caption)

                        'R = CType(myDataGrid.DataSource, DataView).Table.Columns.Item(ic).MaxLength
                        R = myDataGrid.Columns(ic).Width

                        W = HorizontalAlignment.Center

                        'Select Case myDataView.Table.Columns.Item(ic).DataType
                        '    Case System.Type.GetType("System.String"), _
                        '         System.Type.GetType("System.TimeSpan"), _
                        '         System.Type.GetType("System.DateTime")
                        '        W = HorizontalAlignment.Left

                        '    Case System.Type.GetType("System.Boolean"), _
                        '         System.Type.GetType("System.Byte"), _
                        '         System.Type.GetType("System.SByte"), _
                        '         System.Type.GetType("System.Char")
                        '        W = HorizontalAlignment.Center

                        '    Case Else
                        '        W = HorizontalAlignment.Right

                        'End Select

                        WpiszNaglowekDoArkusza(worksheet, AdrColDlaExcela(ie) & "3", N, pts2mm(R), W)
                        ie += 1
                    End If
                Next

                'WIERSZE DANYCH

                If myDataView.Count > 0 Then
                    For ir = 0 To myDataView.Count - 1
                        dr = myDataView.Item(ir)
                        ie = 0
                        For ic = 0 To myDataView.Table.Columns.Count - 1
                            If myDataGrid.Columns(ic).Visible Then
                                R = myDataGrid.Columns(ic).Width

                                If myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.String") Or _
                                    myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.TimeSpan") Or _
                                    myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.DateTime") Then
                                    W = HorizontalAlignment.Left
                                    F = "@"
                                    If IsDBNull(myDataView.Item(ir).Item(ic)) Then
                                        N = ""
                                    Else
                                        N = Trim(myDataView.Item(ir).Item(ic))
                                    End If

                                ElseIf myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.Boolean") Or _
                                    myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.Byte") Or _
                                    myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.SByte") Or _
                                    myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.Char") Then
                                    W = HorizontalAlignment.Center
                                    F = "@"
                                    If IsDBNull(myDataView.Item(ir).Item(ic)) Then
                                        N = ""
                                    Else
                                        N = Trim(myDataView.Item(ir).Item(ic))
                                    End If

                                Else
                                    W = HorizontalAlignment.Right
                                    F = "########0.000"
                                    If IsDBNull(myDataView.Item(ir).Item(ic)) Then
                                        N = 0.0
                                    Else
                                        N = myDataView.Item(ir).Item(ic)
                                    End If

                                End If

                                WpiszDoArkusza(worksheet, AdrColDlaExcela(ie) & Trim(Str(ir + 4)), N, pts2mm(R), W, F)
                                ie += 1
                            End If
                        Next
                    Next
                End If

                'wybierz góre arkusza...
                Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
                range = worksheet.Range("A1", Missing.Value)
                range.Select()
                app.Visible = True      'EXCEL WIDOCZNY NA EKRANIE



                'Try
                '    ' If user interacted with Excel it will not close when the app object is destroyed, so we close it explicitely
                '    workbook.Saved = True
                '    app.UserControl = False
                '    app.Quit()
                'Catch Outer As COMException
                '    'Console.WriteLine("User closed Excel manually, so we don't have to do that")
                'End Try
            End If
        End If
    End Sub








    Public Sub ExportDataGrid2Excel(ByRef myDataGrid As DataGrid, _
                                ByRef myDataView As DataView, _
                                ByVal myNaglowek As String)
        Dim Text As String = ""
        Dim Text2 As String = ""
        Dim Num As Integer = 0
        'Dim MyDataView As DataView = CType(myDataGrid.DataSource, DataView)

        Dim N As Object = Nothing
        Dim T As String = ""
        Dim F As String = ""
        Dim R As Integer = 0
        Dim H As String = ""
        Dim W As Integer = 0
        Dim I As Integer = 0
        '---------------------------------
        Dim app As Application
        app = New Application()
        If app Is Nothing Then
            MessageBox.Show("Nie mogę uruchomić program EXCEL", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            app.Visible = False      'EXCEL WIDOCZNY NA EKRANIE

            Dim workbooks As Workbooks      'Getting the workbooks collection
            workbooks = app.Workbooks

            Dim workbook As _Workbook       'adding a new workbook
            workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)

            Dim sheets As Sheets            'Getting the worksheets collection
            sheets = workbook.Worksheets

            Dim worksheet As _Worksheet     'adding a new worksheet
            worksheet = sheets.Item(1)
            If worksheet Is Nothing Then
                MessageBox.Show("Nie mogę dodać nowego arkusza EXCEL", "Uwaga :", _
                                 System.Windows.Forms.MessageBoxButtons.OK, _
                                 MessageBoxIcon.Information, _
                                 MessageBoxDefaultButton.Button1)
            Else
                worksheet.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnRight
                worksheet.Columns.ColumnWidth = 100

                Dim ir As Integer = 0
                Dim ic As Integer = 0
                Dim dr As DataRowView

                'TYTUŁ TABELI - wiersz nr 1
                WpiszTytulDoArkusza(worksheet, "A1", myNaglowek)

                'WIERSZ NAGŁÓWKOWY - wiersz nr 3
                Dim ie As Integer = 0
                For ic = 0 To myDataView.Table.Columns.Count - 1
                    'If myDataGrid.Columns(ic).Visible Then
                    If myDataGrid.TableStyles(0).GridColumnStyles(ic).Width > 0 Then

                        'N = myDataGrid.Columns(ic).DataPropertyName
                        'N = myDataGrid.Columns(ic).HeaderText
                        'N = Trim(myDataView.Table.Columns.Item(ic).Caption)
                        N = myDataGrid.TableStyles(0).GridColumnStyles(ic).HeaderText


                        'R = CType(myDataGrid.DataSource, DataView).Table.Columns.Item(ic).MaxLength
                        'R = myDataGrid.Columns(ic).Width
                        R = myDataGrid.TableStyles(0).GridColumnStyles(ic).Width

                        W = HorizontalAlignment.Center

                        'Select Case myDataView.Table.Columns.Item(ic).DataType
                        '    Case System.Type.GetType("System.String"), _
                        '         System.Type.GetType("System.TimeSpan"), _
                        '         System.Type.GetType("System.DateTime")
                        '        W = HorizontalAlignment.Left

                        '    Case System.Type.GetType("System.Boolean"), _
                        '         System.Type.GetType("System.Byte"), _
                        '         System.Type.GetType("System.SByte"), _
                        '         System.Type.GetType("System.Char")
                        '        W = HorizontalAlignment.Center

                        '    Case Else
                        '        W = HorizontalAlignment.Right

                        'End Select

                        WpiszNaglowekDoArkusza(worksheet, AdrColDlaExcela(ie) & "3", N, pts2mm(R), W)
                        ie += 1
                    End If
                Next

                'WIERSZE DANYCH

                If myDataView.Count > 0 Then
                    For ir = 0 To myDataView.Count - 1
                        dr = myDataView.Item(ir)
                        ie = 0
                        For ic = 0 To myDataView.Table.Columns.Count - 1
                            'If myDataGrid.Columns(ic).Visible Then
                            If myDataGrid.TableStyles(0).GridColumnStyles(ic).Width > 0 Then
                                'R = myDataGrid.Columns(ic).Width
                                R = myDataGrid.TableStyles(0).GridColumnStyles(ic).Width

                                If myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.String") Or _
                                    myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.TimeSpan") Or _
                                    myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.DateTime") Then
                                    W = HorizontalAlignment.Left
                                    F = "@"
                                    If IsDBNull(myDataView.Item(ir).Item(ic)) Then
                                        N = ""
                                    Else
                                        N = Trim(myDataView.Item(ir).Item(ic))
                                    End If

                                ElseIf myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.Boolean") Or _
                                    myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.Byte") Or _
                                    myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.SByte") Or _
                                    myDataView.Table.Columns.Item(ic).DataType Is System.Type.GetType("System.Char") Then
                                    W = HorizontalAlignment.Center
                                    F = "@"
                                    If IsDBNull(myDataView.Item(ir).Item(ic)) Then
                                        N = ""
                                    Else
                                        N = Trim(myDataView.Item(ir).Item(ic))
                                    End If

                                Else
                                    W = HorizontalAlignment.Right
                                    F = "########0.000"
                                    If IsDBNull(myDataView.Item(ir).Item(ic)) Then
                                        N = 0.0
                                    Else
                                        N = myDataView.Item(ir).Item(ic)
                                    End If

                                End If

                                WpiszDoArkusza(worksheet, AdrColDlaExcela(ie) & Trim(Str(ir + 4)), N, pts2mm(R), W, F)
                                ie += 1
                            End If
                        Next
                    Next
                End If

                'wybierz góre arkusza...
                Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
                range = worksheet.Range("A1", Missing.Value)
                range.Select()
                app.Visible = True      'EXCEL WIDOCZNY NA EKRANIE



                'Try
                '    ' If user interacted with Excel it will not close when the app object is destroyed, so we close it explicitely
                '    workbook.Saved = True
                '    app.UserControl = False
                '    app.Quit()
                'Catch Outer As COMException
                '    'Console.WriteLine("User closed Excel manually, so we don't have to do that")
                'End Try
            End If
        End If
    End Sub







    '===================================
    ' przepisujemy dane z DataView do EXCELa
    '===================================


    ''---------------------------------------
    ''zaloz strukture robocza do ktorej bedziemy zbierac dane do Raportu
    'Dim DataSetRaport As DataSet = New DataSet("Raport")
    ''zakladamy tablice
    'Dim TableRaport As System.Data.DataTable = DataSetRaport.Tables.Add("TabRaport")
    ''zakladamy kolumny
    '' SYMBOL | NAZWA | ILEZAMOWIONO | DOKWZ | ZAMKLIENTA
    'Dim KolumnaRaport As DataColumn = TableRaport.Columns.Add("Symbol towaru", Type.GetType("System.String"))
    '                    TableRaport.Columns.Item("Symbol towaru").MaxLength = 20
    '                    TableRaport.Columns.Add("Nazwa towaru", Type.GetType("System.String"))
    '                    TableRaport.Columns.Item("Nazwa towaru").MaxLength = 50
    '                    TableRaport.Columns.Add("Ile na Zleceniu", Type.GetType("System.String"))
    '                    TableRaport.Columns.Item("Ile na Zleceniu").MaxLength = 15
    '                    TableRaport.Columns.Add("Dok WZ", Type.GetType("System.String"))
    '                    TableRaport.Columns.Item("Dok WZ").MaxLength = 15
    '                    TableRaport.Columns.Add("Zam klienta", Type.GetType("System.String"))
    '                    TableRaport.Columns.Item("Zam klienta").MaxLength = 15
    '                    TableRaport.Columns.Add("Uwagi", Type.GetType("System.String"))
    '                    TableRaport.Columns.Item("Uwagi").MaxLength = 50
    ''wybierz klucz glowny...
    ''TableRaport.PrimaryKey = New DataColumn() {KolumnaRaport}
    'Dim DataViewRaport As New DataView(DataSetRaport.Tables(0))



    Public Sub ExportDataView2Excel(ByRef myDataView As DataView, _
                                    ByRef myDataViewTotal As DataView, _
                                    ByRef myFormat() As String, _
                                    ByRef myNaglowek1() As String, _
                                    ByRef myNaglowek2() As String, _
                                    ByRef myNaglowek As String)
        Dim Text As String = ""
        Dim Text2 As String = ""
        Dim Num As Integer = 0

        Dim N As Object = Nothing
        Dim T As System.Type = Nothing
        Dim R As Integer = 0
        Dim H As String = ""
        Dim F As String = ""
        Dim W As Integer = 0
        Dim I As Integer = 0
        '---------------------------------
        Dim app As Application
        app = New Application()
        If app Is Nothing Then
            MessageBox.Show("Nie mogę uruchomić program EXCEL", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
        Else
            app.Visible = False      'EXCEL WIDOCZNY NA EKRANIE

            Dim workbooks As Workbooks      'Getting the workbooks collection
            workbooks = app.Workbooks

            Dim workbook As _Workbook       'adding a new workbook
            workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)

            Dim sheets As Sheets            'Getting the worksheets collection
            sheets = workbook.Worksheets

            Dim worksheet As _Worksheet     'adding a new worksheet
            worksheet = sheets.Item(1)
            If worksheet Is Nothing Then
                MessageBox.Show("Nie mogę dodać nowego arkusza EXCEL", "Uwaga :", _
                                 System.Windows.Forms.MessageBoxButtons.OK, _
                                 MessageBoxIcon.Information, _
                                 MessageBoxDefaultButton.Button1)
            Else
                worksheet.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnRight
                worksheet.Columns.ColumnWidth = 100

                Dim ir As Integer = 0
                Dim ic As Integer = 0
                Dim dr As DataRowView

                'TYTUŁ TABELI - wiersz nr 1
                WpiszTytulDoArkusza(worksheet, "A1", myNaglowek)

                'WIERSZ NAGŁÓWKOWY - wiersz nr 3 i 4
                'For ic = 0 To myDataView.Table.Columns.Count - 1
                For ic = 0 To myNaglowek1.Length - 1
                    'F = myFormat(ic)      'myDataView.Table.Columns.Item(ic).Caption
                    N = myNaglowek1(ic)

                    If ic > myDataView.Table.Columns.Count - 1 Then
                        W = HorizontalAlignment.Center
                        R = 80      'myDataView.Table.Columns.Item(ic).MaxLength
                    Else
                        T = myDataView.Table.Columns.Item(ic).DataType
                        If T Is Type.GetType("System.String") Or _
                            T Is Type.GetType("System.TimeSpan") Or _
                            T Is Type.GetType("System.DateTime") Then
                            W = HorizontalAlignment.Center
                            R = 80      'myDataView.Table.Columns.Item(ic).MaxLength

                        ElseIf T Is Type.GetType("System.Boolean") Or _
                               T Is Type.GetType("System.Byte") Or _
                               T Is Type.GetType("System.SByte") Or _
                               T Is Type.GetType("System.Char") Then
                            W = HorizontalAlignment.Center
                            R = 20

                        Else
                            W = HorizontalAlignment.Center
                            R = 60

                        End If
                    End If

                    WpiszNaglowekDoArkusza(worksheet, AdrColDlaExcela(ic) & "3", N, pts2mm(R), W)
                Next

                'For ic = 0 To myDataView.Table.Columns.Count - 1
                For ic = 0 To myNaglowek2.Length - 1
                    'F = myFormat(ic)      'myDataView.Table.Columns.Item(ic).Caption
                    N = myNaglowek2(ic)

                    If ic > myDataView.Table.Columns.Count - 1 Then
                        W = HorizontalAlignment.Center
                        R = 80      'myDataView.Table.Columns.Item(ic).MaxLength
                    Else
                        T = myDataView.Table.Columns.Item(ic).DataType
                        If T Is Type.GetType("System.String") Or _
                            T Is Type.GetType("System.TimeSpan") Or _
                            T Is Type.GetType("System.DateTime") Then
                            W = HorizontalAlignment.Center
                            R = 80      'myDataView.Table.Columns.Item(ic).MaxLength

                        ElseIf T Is Type.GetType("System.Boolean") Or _
                               T Is Type.GetType("System.Byte") Or _
                               T Is Type.GetType("System.SByte") Or _
                               T Is Type.GetType("System.Char") Then
                            W = HorizontalAlignment.Center
                            R = 20

                        Else
                            W = HorizontalAlignment.Center
                            R = 60

                        End If
                    End If
                    WpiszNaglowekDoArkusza(worksheet, AdrColDlaExcela(ic) & "4", N, pts2mm(R), W)
                Next




                '-------------------
                'WIERSZE DANYCH - od wiersza nr 5
                If myDataView.Count > 0 Then
                    For ir = 0 To myDataView.Count - 1
                        dr = myDataView.Item(ir)
                        For ic = 0 To myDataView.Table.Columns.Count - 1
                            F = myFormat(ic)      'myDataView.Table.Columns.Item(ic).Caption

                            T = myDataView.Table.Columns.Item(ic).DataType
                            If T Is Type.GetType("System.String") Or _
                                T Is Type.GetType("System.TimeSpan") Or _
                                T Is Type.GetType("System.DateTime") Then
                                R = 80      'myDataView.Table.Columns.Item(ic).MaxLength
                                W = HorizontalAlignment.Left
                                If IsDBNull(myDataView.Item(ir).Item(ic)) Then
                                    N = ""
                                Else
                                    N = Trim(myDataView.Item(ir).Item(ic))
                                End If

                            ElseIf T Is Type.GetType("System.Boolean") Or _
                                   T Is Type.GetType("System.Byte") Or _
                                   T Is Type.GetType("System.SByte") Or _
                                   T Is Type.GetType("System.Char") Then
                                R = 20
                                W = HorizontalAlignment.Center
                                If IsDBNull(myDataView.Item(ir).Item(ic)) Then
                                    N = ""
                                Else
                                    N = Trim(myDataView.Item(ir).Item(ic))
                                End If

                            Else
                                R = 60
                                W = HorizontalAlignment.Right
                                If IsDBNull(myDataView.Item(ir).Item(ic)) Then
                                    N = 0
                                Else
                                    N = myDataView.Item(ir).Item(ic)
                                End If

                            End If

                            WpiszDoArkusza(worksheet, AdrColDlaExcela(ic) & Trim(Str(ir + 5)), N, pts2mm(R), W, F)
                        Next
                    Next
                End If

                If myDataViewTotal Is Nothing Then
                Else
                    If myDataViewTotal.Count > 0 Then
                        For ir = 0 To myDataViewTotal.Count - 1
                            dr = myDataViewTotal.Item(ir)
                            For ic = 0 To myDataViewTotal.Table.Columns.Count - 1
                                F = myFormat(ic)      'myDataViewTotal.Table.Columns.Item(ic).Caption

                                T = myDataViewTotal.Table.Columns.Item(ic).DataType
                                If T Is Type.GetType("System.String") Or _
                                    T Is Type.GetType("System.TimeSpan") Or _
                                    T Is Type.GetType("System.DateTime") Then
                                    R = 80      'myDataViewTotal.Table.Columns.Item(ic).MaxLength
                                    W = HorizontalAlignment.Left
                                    If IsDBNull(myDataViewTotal.Item(ir).Item(ic)) Then
                                        N = ""
                                    Else
                                        N = Trim(myDataViewTotal.Item(ir).Item(ic))
                                    End If

                                ElseIf T Is Type.GetType("System.Boolean") Or _
                                       T Is Type.GetType("System.Byte") Or _
                                       T Is Type.GetType("System.SByte") Or _
                                       T Is Type.GetType("System.Char") Then
                                    R = 20
                                    W = HorizontalAlignment.Center
                                    If IsDBNull(myDataViewTotal.Item(ir).Item(ic)) Then
                                        N = ""
                                    Else
                                        N = Trim(myDataViewTotal.Item(ir).Item(ic))
                                    End If

                                Else
                                    R = 60
                                    W = HorizontalAlignment.Right
                                    If IsDBNull(myDataViewTotal.Item(ir).Item(ic)) Then
                                        N = 0
                                    Else
                                        N = myDataViewTotal.Item(ir).Item(ic)
                                    End If

                                End If

                                WpiszDoArkusza(worksheet, AdrColDlaExcela(ic) & Trim(Str(myDataView.Count + ir + 5)), N, pts2mm(R), W, F)
                            Next
                        Next
                    End If
                End If





                'wybierz góre arkusza...
                Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
                range = worksheet.Range("A1", Missing.Value)
                range.Select()
                app.Visible = True      'EXCEL WIDOCZNY NA EKRANIE



                'Try
                '    ' If user interacted with Excel it will not close when the app object is destroyed, so we close it explicitely
                '    workbook.Saved = True
                '    app.UserControl = False
                '    app.Quit()
                'Catch Outer As COMException
                '    'Console.WriteLine("User closed Excel manually, so we don't have to do that")
                'End Try
            End If
        End If
    End Sub
































    '    '------------------------------------------
    '    'UWAGA :
    '    'Do referencji trzeba DODAC
    '    '   Microsoft Excel 11.0 Object Library
    '    '   Microsoft Office 11.0 Object Library
    '    '-------------------------------------------
    'Imports System
    'Imports System.Reflection ' For Missing.Value and BindingFlags
    'Imports System.Runtime.InteropServices ' For COMException
    'Imports Microsoft.Office.Interop.Excel

    'Module modExcel

    '        Private Function AdrColDlaExcela(ByVal colNo As Integer) As String
    '            Dim IdWierszy As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    '            Dim colId As String = ""
    '            If colNo < Len(IdWierszy) Then
    '                colId = Mid(IdWierszy, colNo + 1, 1)
    '            Else
    '                colId = Mid(IdWierszy, Int(colNo / Len(IdWierszy)), 1) & Mid(IdWierszy, (colNo Mod Len(IdWierszy)) + 1, 1)
    '            End If
    '            Return (colId)
    '        End Function


    '        Public Sub ExportDataGrid2Excel2(ByRef myDataGrid As DataGrid)
    '            'On Error GoTo er
    '            Try

    '                Dim cm As CurrencyManager = CType(myDataGrid.BindingContext(myDataGrid.DataSource), CurrencyManager)
    '                Dim myDataGridRowsNo As Integer = cm.Count
    '                'Dim myDataGridColumnsNo As Integer = CType(myDataGrid.DataSource, DataTable).Columns.Count

    '                Dim myDataGridColumnsNo As Integer = CType(myDataGrid.DataSource, DataView).ToTable.Columns.Count

    '                Dim oExcel As Object = CreateObject("Excel.Application")

    '                'Dim oWorkBook As Object
    '                Dim oWorkSheet As Object

    '                Dim i As Integer = 0
    '                Dim k As Integer = 0
    '                Dim lRow As Long = 0
    '                Dim LastRow As Long = 0
    '                Dim LastCol As Long = 0

    '                Dim ExcelRow As Integer = 0
    '                Dim ExcelCol As Integer = 0
    '                Dim ExcelRowOffset As Integer = 2
    '                Dim ExcelColOffset As Integer = 0

    '                oExcel.Visible = False
    '                oExcel.Workbooks.Open(oKatProgramu & "\Nirmal.xls")
    '                oWorkSheet = oExcel.Workbooks("Nirmal.xls").sheets("Batch")

    '                'zapamietaj lokalizacej kursora
    '                LastRow = myDataGrid.CurrentCell.RowNumber 'Save Current row
    '                LastCol = myDataGrid.CurrentCell.ColumnNumber 'and column
    '                '-------------------------------------
    '                Dim row As Integer
    '                Dim col As Integer
    '                Dim cellValue As Object = Nothing
    '                For row = 0 To myDataGridRowsNo - 1
    '                    ExcelRow = row + ExcelRowOffset
    '                    For col = 0 To myDataGridColumnsNo - 1
    '                        cellValue = myDataGrid(row, col)
    '                        ExcelCol = col + ExcelColOffset

    '                        oWorkSheet.Cells(ExcelRow, ExcelCol).Font.Bold = False
    '                        oWorkSheet.Cells(ExcelRow, ExcelCol).Font.Color = System.Drawing.Color.Navy
    '                        oWorkSheet.Cells(ExcelRow, ExcelCol).value = cellValue.ToString
    '                    Next col
    '                Next row
    '                '-------------------------------------
    '                myDataGrid.CurrentCell = New DataGridCell(LastRow, LastCol)




    '                oExcel.Workbooks("Nirmal.xls").Save()
    '                oExcel.Workbooks("Nirmal.xls").Close(savechanges:=True)
    '                oExcel.Quit()

    '            Catch ex As Exception
    '                MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & ex.Message, "Uwaga :", _
    '                    System.Windows.Forms.MessageBoxButtons.OK, _
    '                    MessageBoxIcon.Information, _
    '                    MessageBoxDefaultButton.Button1)
    '            End Try

    '            'If Err.Number = 1004 Then
    '            '    Exit Sub
    '            'End If
    '        End Sub


    '        '===================================
    '        ' przepisujemy dane adresowe do EXCELa
    '        '===================================

    '        '-------------odbc
    '        'Private EXCELConnection As String = "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & oPlikNaDysku & ";"
    '        Private EXCELConnectionODBC As String = _
    '                    "Driver={Microsoft Excel Driver (*.xls)};" & _
    '                    "DSN=Pliki programu Excel;" & _
    '                    "DBQ=" & oPlikNaDysku & ";" & _
    '                    "DriverId=790;" & _
    '                    "MaxBufferSize=2048;" & _
    '                    "PageTimeout=5;"

    '        '-------------oleDB
    '        Private EXCELConnectionOLeDB As String = "provider=Microsoft.Jet.OLEDB.4.0; " & _
    '                                        "data source=" & oPlikNaDysku & "; " & _
    '                                        "Extended Properties='Excel 8.0; HDR=YES; IMEX=1';"

    '        '=========================================================
    '        'Connection String: 
    '        'The connection string should be set to the OleDbConnection object.
    '        'This is very critical as Jet Engine might not give a proper error message
    '        'if the appropriate details are not given.
    '        '
    '        'Syntax: Provider=Microsoft.Jet.OLEDB.4.0;
    '        '                 Data Source=<Full Path of Excel File>;
    '        '                 Extended Properties="Excel 8.0; HDR=No; IMEX=1".
    '        '
    '        'Definition of Extended Properties: 
    '        '   Excel = <No> 
    '        '   One should specify the version of Excel Sheet here.
    '        '   For Excel 2000 and above, it is set it to Excel 8.0
    '        '   and for all others, it is Excel 5.0.
    '        '
    '        '   HDR= <Yes/No> 
    '        '   This property will be used to specify the definition of header
    '        '   for each column. If the value is ‘Yes’, the first row will be
    '        '   treated as heading. Otherwise, the heading will be generated
    '        '   by the system like F1, F2 and so on.
    '        '
    '        '   IMEX= <0/1/2> 
    '        '   IMEX refers to IMport EXport mode. This can take three possible values.
    '        '   IMEX=0 and IMEX=2 will result in ImportMixedTypes being ignored and the default value of  
    '        '       'Majority Types’ is used. In this case, it will take the first 8 rows
    '        '       and then the data type for each column will be decided. 
    '        '   IMEX=1 is the only way to set the value of ImportMixedTypes
    '        '       as Text. Here, everything will be treated as text. 
    '        '=========================================================

    '        Public Sub ExportDataGrid2Excel(ByRef myDataGrid As DataGridView, ByRef myDataView As DataView)
    '            Dim Text As String = ""
    '            Dim Text2 As String = ""
    '            Dim Num As Integer = 0
    '            'Dim MyDataView As DataView = CType(myDataGrid.DataSource, DataView)

    '            Dim N As String = ""
    '            Dim T As String = ""
    '            Dim R As Integer = 0
    '            Dim H As String = ""
    '            Dim W As Integer = 0
    '            Dim I As Integer = 0
    '            '---------------------------------
    '            Dim app As Application
    '            app = New Application()
    '            If app Is Nothing Then
    '                MessageBox.Show("Nie mogę uruchomić program EXCEL", "Uwaga :", _
    '                    System.Windows.Forms.MessageBoxButtons.OK, _
    '                    MessageBoxIcon.Information, _
    '                    MessageBoxDefaultButton.Button1)
    '            Else
    '                app.Visible = False      'EXCEL WIDOCZNY NA EKRANIE

    '                Dim workbooks As Workbooks      'Getting the workbooks collection
    '                workbooks = app.Workbooks

    '                Dim workbook As _Workbook       'adding a new workbook
    '                workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet)

    '                Dim sheets As Sheets            'Getting the worksheets collection
    '                sheets = workbook.Worksheets

    '                Dim worksheet As _Worksheet     'adding a new worksheet
    '                worksheet = sheets.Item(1)
    '                If worksheet Is Nothing Then
    '                    MessageBox.Show("Nie mogę dodać nowego arkusza EXCEL", "Uwaga :", _
    '                                     System.Windows.Forms.MessageBoxButtons.OK, _
    '                                     MessageBoxIcon.Information, _
    '                                     MessageBoxDefaultButton.Button1)
    '                Else
    '                    worksheet.Outline.SummaryColumn = XlSummaryColumn.xlSummaryOnRight
    '                    worksheet.Columns.ColumnWidth = 100

    '                    Dim ir As Integer = 0
    '                    Dim ic As Integer = 0
    '                    Dim dr As DataRowView

    '                    'WIERSZ NAGŁÓWKOWY
    '                    For ic = 0 To myDataView.Table.Columns.Count - 1
    '                        'N = myDataGrid.TableStyles(0).GridColumnStyles(ic).HeaderText
    '                        ''N = Trim(MyDataView.Table.Columns.Item(ic).Caption)
    '                        ''R = CType(myDataGrid.DataSource, DataView).Table.Columns.Item(ic).MaxLength
    '                        'R = myDataGrid.TableStyles(0).GridColumnStyles.Item(ic).Width
    '                        'W = DataGridViewContentAlignment.MiddleCenter

    '                        N = myDataGrid.Columns(ic).Name
    '                        R = myDataGrid.Columns(ic).Width
    '                        W = DataGridViewContentAlignment.MiddleCenter

    '                        'Select Case myDataView.Table.Columns.Item(ic).DataType
    '                        '    Case System.Type.GetType("System.String"), _
    '                        '         System.Type.GetType("System.TimeSpan"), _
    '                        '         System.Type.GetType("System.DateTime")
    '                        '        W = DataGridViewContentAlignment.MiddleLeft

    '                        '    Case System.Type.GetType("System.Boolean"), _
    '                        '         System.Type.GetType("System.Byte"), _
    '                        '         System.Type.GetType("System.SByte"), _
    '                        '         System.Type.GetType("System.Char")
    '                        '        W = DataGridViewContentAlignment.MiddleCenter

    '                        '    Case Else
    '                        '        W = HorizontalAlignment.Right

    '                        'End Select

    '                        WpiszNaglowekDoArkusza(worksheet, AdrColDlaExcela(ic) & "1", N, pts2mm(R), W)
    '                    Next

    '                    'WIERSZE DANYCH
    '                    If myDataView.Count > 0 Then
    '                        For ir = 0 To myDataView.Count - 1
    '                            dr = myDataView.Item(ir)
    '                            For ic = 0 To myDataView.Table.Columns.Count - 1
    '                                If IsDBNull(myDataView.Item(ir).Item(ic)) Then
    '                                    N = ""
    '                                Else
    '                                    N = Trim(myDataView.Item(ir).Item(ic))
    '                                End If
    '                                'N = IIf(IsDBNull(myDataView.Item(ir).Item(ic)), "", Trim(myDataView.Item(ir).Item(ic)))
    '                                'R = MYDataView.Table.Columns.Item(ic).MaxLength
    '                                'R = myDataGrid.TableStyles(0).GridColumnStyles.Item(ic).Width
    '                                R = myDataGrid.Columns(ic).Width

    '                                Select Case myDataView.Table.Columns.Item(ic).DataType.Name
    '                                    Case "String", _
    '                                         "TimeSpan", _
    '                                         "DateTime"
    '                                        W = DataGridViewContentAlignment.MiddleLeft

    '                                    Case "Boolean", _
    '                                         "Byte", _
    '                                         "SByte", _
    '                                         "Char"
    '                                        W = DataGridViewContentAlignment.MiddleCenter

    '                                    Case Else
    '                                        W = HorizontalAlignment.Right

    '                                End Select
    '                                WpiszDoArkusza(worksheet, AdrColDlaExcela(ic) & Trim(Str(ir + 2)), N, pts2mm(R), W)
    '                            Next
    '                        Next
    '                    End If

    '                    'wybierz góre arkusza...
    '                    Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
    '                    range = worksheet.Range("A1", Missing.Value)
    '                    range.Select()
    '                    app.Visible = True      'EXCEL WIDOCZNY NA EKRANIE



    '                    'Try
    '                    '    ' If user interacted with Excel it will not close when the app object is destroyed, so we close it explicitely
    '                    '    workbook.Saved = True
    '                    '    app.UserControl = False
    '                    '    app.Quit()
    '                    'Catch Outer As COMException
    '                    '    'Console.WriteLine("User closed Excel manually, so we don't have to do that")
    '                    'End Try
    '                End If
    '            End If
    '        End Sub





    '        Private Sub WpiszNaglowekDoArkusza(ByVal Arkusz As _Worksheet, _
    '                                   ByVal Adres As String, _
    '                                   ByVal Zawartosc As String, _
    '                                   ByVal Szerokosc As Integer, _
    '                                   ByVal Wyrownanie As Integer)
    '            Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
    '            'Dim ColRange As Microsoft.Office.Interop.Excel.Range
    '            'Dim RowRange As Microsoft.Office.Interop.Excel.Range

    '            range = Arkusz.Range(Adres, Missing.Value)
    '            If range Is Nothing Then
    '                MessageBox.Show("Nie mogę zdefiniować zakresu arkusza EXCEL", "Uwaga :", _
    '                    System.Windows.Forms.MessageBoxButtons.OK, _
    '                    MessageBoxIcon.Information, _
    '                    MessageBoxDefaultButton.Button1)
    '            Else
    '                range.Value2 = Zawartosc
    '                range.ColumnWidth = Szerokosc
    '                range.RowHeight = 15
    '                range.WrapText = True
    '                'range.HorizontalAlignment = Wyrownanie
    '                range.Font.Bold = True
    '                range.Font.Color = System.Drawing.Color.Black.ToArgb


    '                'dotyczy calej kolumny
    '                'ColRange = range.Columns
    '                'Select Case Wyrownanie
    '                '    Case "Do lewej"
    '                '        ColRange.HorizontalAlignment = DataGridViewContentAlignment.MiddleLeft
    '                '    Case "Do prawej"
    '                '        ColRange.HorizontalAlignment = HorizontalAlignment.Right
    '                '    Case Else
    '                '        ColRange.HorizontalAlignment = DataGridViewContentAlignment.MiddleCenter
    '                'End Select

    '                'RowRange = range.Rows
    '                'RowRange.VerticalAlignment

    '            End If

    '        End Sub





    '        Private Sub WpiszDoArkusza(ByVal Arkusz As _Worksheet, _
    '                                   ByVal Adres As String, _
    '                                   ByVal Zawartosc As String, _
    '                                   ByVal Szerokosc As Integer, _
    '                                   ByVal Wyrownanie As Integer)
    '            Dim range As Microsoft.Office.Interop.Excel.Range             'Setting the value for cell
    '            'Dim ColRange As Microsoft.Office.Interop.Excel.Range
    '            'Dim RowRange As Microsoft.Office.Interop.Excel.Range

    '            range = Arkusz.Range(Adres, Missing.Value)
    '            If range Is Nothing Then
    '                MessageBox.Show("Nie mogę zdefiniować zakresu arkusza EXCEL", "Uwaga :", _
    '                    System.Windows.Forms.MessageBoxButtons.OK, _
    '                    MessageBoxIcon.Information, _
    '                    MessageBoxDefaultButton.Button1)
    '            Else
    '                Try
    '                    range.Value2 = Zawartosc
    '                    range.ColumnWidth = Szerokosc
    '                    range.RowHeight = 15
    '                    range.WrapText = True
    '                    'range.HorizontalAlignment = Wyrownanie
    '                    range.Font.Bold = False
    '                    range.Font.Color = System.Drawing.Color.Navy.ToArgb

    '                    'dotyczy calej kolumny
    '                    'ColRange = range.Columns
    '                    'Select Case Wyrownanie
    '                    '    Case "Do lewej"
    '                    '        ColRange.HorizontalAlignment = DataGridViewContentAlignment.MiddleLeft
    '                    '    Case "Do prawej"
    '                    '        ColRange.HorizontalAlignment = HorizontalAlignment.Right
    '                    '    Case Else
    '                    '        ColRange.HorizontalAlignment = DataGridViewContentAlignment.MiddleCenter
    '                    'End Select

    '                    'RowRange = range.Rows
    '                    'RowRange.VerticalAlignment
    '                Catch ex As Exception
    '                    MessageBox.Show("W programie wystąpił błąd :" & vbCrLf & ex.Message, "Uwaga :", _
    '                        System.Windows.Forms.MessageBoxButtons.OK, _
    '                        MessageBoxIcon.Information, _
    '                        MessageBoxDefaultButton.Button1)
    '                End Try
    '            End If

    '        End Sub





End Module