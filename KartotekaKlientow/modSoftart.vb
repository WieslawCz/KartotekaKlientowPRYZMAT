Module modSoftart

    '**************************************************
    '* testuj architekture systemu
    '**************************************************
    'Public Bit64System As Boolean = System.Environment.Is64BitOperatingSystem
    'Public Bit64Process As Boolean = System.Environment.Is64BitProcess

    Public Function OSArchitectureIs64bit() As String

        'Environment Variable       32bit   64bit   WOW64
        '                           Native  Native 
        '----------------------------------------------------
        'PROCESSOR_ARCHITECTURE     x86     AMD64   x86 
        'PROCESSOR_ARCHITEW6432     undef   undef   MD64 

        Dim PROCESSOR_ARCHITECTURE As String = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE")
        Dim PROCESSOR_ARCHITEW6432 As String = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432")

        If PROCESSOR_ARCHITECTURE = "AMD64" Or PROCESSOR_ARCHITEW6432 = "AMD64" Then
            ' OS is 64bit
            Return True
        Else
            ' OS is 32bit
            Return False
        End If


        ''-----------------------------------
        'Dim _osVersion As String = ""
        'Dim _osServicePack As String = ""
        'Dim _osArchitecture As String = ""

        'Dim searcher As System.Management.ManagementObjectSearcher = New System.Management.ManagementObjectSearcher("select * from Win32_OperatingSystem")
        'Dim collect As System.Management.ManagementObjectCollection = searcher.Get()
        'Dim mbo As System.Management.ManagementObject
        'For Each mbo In collect
        '    _osVersion = mbo.GetPropertyValue("Caption").ToString()
        '    _osServicePack = String.Format("{0}.{1}", mbo.GetPropertyValue("ServicePackMajorVersion").ToString(), mbo.GetPropertyValue("ServicePackMinorVersion").ToString())
        '    Try
        '        _osArchitecture = mbo.GetPropertyValue("OSArchitecture").ToString()
        '    Catch
        '        ' OSArchitecture only supported on Win7/W2K8 
        '    End Try
        'Next mbo
        'Console.WriteLine("osVersion     : " + _osVersion)
        'Console.WriteLine("osServicePack : " + _osServicePack)
        'Console.WriteLine("osArchitecture: " + _osArchitecture)
        ''-----------------------------------D:\!__VB.vs2017\_Pryzmat\KartotekaKlientow_3\KartotekaKlientow\KatalogKartPayBack.vb
    End Function



    '**************************************************
    '* procedura od której startujemy
    '**************************************************

    Public Sub Main()
        Dim StartForm As New Zaczynamy
        StartForm.ShowDialog()
    End Sub




    'Public SoftartRysujGradientNColor As Color = System.Drawing.Color.FromArgb(CType(180, Byte),
    '                                                                       CType(180, Byte),
    '                                                                       CType(255, Byte))
    Public SoftartRysujGradientnColor As Color = System.Drawing.Color.Navy
    Public Sub RysujGradientN(ByVal meForm As Form)
        ' gradient w tle formy
        Dim G As Graphics
        G = meForm.CreateGraphics
        Dim R As New System.Drawing.RectangleF(0, 0, meForm.Width, meForm.Height)

        Dim startColor As Color = System.Drawing.SystemColors.Control
        Dim EndColor As Color = SoftartRysujGradientnColor

        Dim LGBrush As New System.Drawing.Drawing2D.LinearGradientBrush _
                (R, startColor, EndColor, System.Drawing.Drawing2D.LinearGradientMode.Horizontal)
        G.FillRectangle(LGBrush, New System.Drawing.Rectangle(0, 0, meForm.Width, meForm.Height))
    End Sub

    'Public SoftartRysujGradientZColor As Color = System.Drawing.Color.FromArgb(CType(120, Byte),
    '                                                                       CType(214, Byte),
    '                                                                       CType(132, Byte))
    Public SoftartRysujGradientzColor As Color = System.Drawing.Color.Green
    Public Sub RysujGradientZ(ByVal meForm As Form)
        ' gradient w tle formy
        Dim G As Graphics
        G = meForm.CreateGraphics
        Dim R As New System.Drawing.RectangleF(0, 0, meForm.Width, meForm.Height)

        Dim startColor As Color = System.Drawing.SystemColors.Control
        Dim EndColor As Color = SoftartRysujGradientzColor

        Dim LGBrush As New System.Drawing.Drawing2D.LinearGradientBrush _
                (R, startColor, EndColor, System.Drawing.Drawing2D.LinearGradientMode.Horizontal)
        G.FillRectangle(LGBrush, New System.Drawing.Rectangle(0, 0, meForm.Width, meForm.Height))
    End Sub

    'Public SoftartRysujGradientCColor As Color = System.Drawing.Color.FromArgb(CType(255, Byte),
    '                                                                       CType(150, Byte),
    '                                                                       CType(150, Byte))
    Public SoftartRysujGradientCColor As Color = System.Drawing.Color.Red
    Public Sub RysujGradientC(ByVal meForm As Form)
        Dim LeftOffset As Integer = (8 + 300)
        ' gradient w tle formy
        Dim G As Graphics
        G = meForm.CreateGraphics
        Dim R As New System.Drawing.RectangleF(LeftOffset, 0, meForm.Width - LeftOffset, meForm.Height)

        Dim startColor As Color = System.Drawing.SystemColors.Control
        Dim EndColor As Color = SoftartRysujGradientCColor

        Dim LGBrush As New System.Drawing.Drawing2D.LinearGradientBrush _
                (R, startColor, EndColor, System.Drawing.Drawing2D.LinearGradientMode.Horizontal)
        G.FillRectangle(LGBrush, New System.Drawing.Rectangle(LeftOffset + 5, 0, meForm.Width - LeftOffset, meForm.Height))
    End Sub




    Public Sub RysujGradient(ByVal meForm As Form)
        ' gradient w tle formy
        Dim G As Graphics
        G = meForm.CreateGraphics
        Dim R As New RectangleF(0, 0, meForm.Width, meForm.Height)
        Dim startColor As Color = System.Drawing.SystemColors.Control
        Dim EndColor As Color = System.Drawing.SystemColors.ControlDark

        Dim LGBrush As New System.Drawing.Drawing2D.LinearGradientBrush _
                (R, startColor, EndColor, System.Drawing.Drawing2D.LinearGradientMode.Horizontal)
        G.FillRectangle(LGBrush, New Rectangle(0, 0, meForm.Width, meForm.Height))
    End Sub

    Public Sub RysujGradientWPanelu(ByVal myPanel As Panel)
        ' gradient w tle formy
        Dim G As Graphics
        G = myPanel.CreateGraphics
        Dim R As New RectangleF(0, 0, myPanel.Width, myPanel.Height)
        Dim startColor As Color = System.Drawing.SystemColors.Control
        Dim EndColor As Color = System.Drawing.SystemColors.ControlDark

        Dim LGBrush As New System.Drawing.Drawing2D.LinearGradientBrush _
                (R, EndColor, startColor, System.Drawing.Drawing2D.LinearGradientMode.Horizontal)
        G.FillRectangle(LGBrush, New Rectangle(0, 0, myPanel.Width, myPanel.Height))
    End Sub

    '**************************************************
    '* definiowanie kolumny w tabeli
    '**************************************************

    Public Sub TxtColumn(ByVal TabStyle As DataGridTableStyle, _
                      ByVal MapName As String, _
                      ByVal HeadText As String, _
                      ByVal ColWidth As Integer, _
                      ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment)
        Dim TCol As New Softart.myDataGridTextBoxColumn
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub

    Public Sub TxtColoredColumn(ByVal TabStyle As DataGridTableStyle, _
                      ByVal MapName As String, _
                      ByVal HeadText As String, _
                      ByVal ColWidth As Integer, _
                      ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment)
        'Dim TCol As New Softart.myColoredDataGridTextBoxColumn
        Dim TCol As New Softart.myDataGridTextBoxColumn
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub

    Public Sub LogColumn(ByVal TabStyle As DataGridTableStyle, _
                      ByVal MapName As String, _
                      ByVal HeadText As String, _
                      ByVal ColWidth As Integer)
        Dim TCol As New Softart.myDataGridBoolColumn
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = HorizontalAlignment.Center
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub

    Public Sub NumColumn(ByVal TabStyle As DataGridTableStyle, _
                      ByVal MapName As String, _
                      ByVal HeadText As String, _
                      ByVal ColWidth As Integer, _
                      ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment, _
                      ByVal ColFormat As String)
        Dim TCol As New Softart.myDataGridNumColumn
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TCol.Format = ColFormat
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub

    '**************************************************
    '* kolrowany grid - BackColor
    '**************************************************

    Public Sub TxtColoredColumnB(ByVal deleg As delegateDefColor, _
                                      ByVal TabStyle As DataGridTableStyle, _
                                      ByVal MapName As String, _
                                      ByVal HeadText As String, _
                                      ByVal ColWidth As Integer, _
                                      ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnB(deleg)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub

    Public Sub NumColoredColumnB(ByVal deleg As delegateDefColor, _
                                      ByVal TabStyle As DataGridTableStyle, _
                                      ByVal MapName As String, _
                                      ByVal HeadText As String, _
                                      ByVal ColWidth As Integer, _
                                      ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment, _
                                      ByVal ColFormat As String)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnB(deleg)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TCol.Format = ColFormat
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub


    '**************************************************
    '* kolorowany grid - ForeColor
    '**************************************************

    Public Sub TxtColoredColumnF(ByVal deleg As delegateDefColor, _
                                      ByVal TabStyle As DataGridTableStyle, _
                                      ByVal MapName As String, _
                                      ByVal HeadText As String, _
                                      ByVal ColWidth As Integer, _
                                      ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnF(deleg)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub

    ' dlA potrzeb funkcji AnalizaZakupow
    Public Sub TxtColoredColumnF_Analiza0(ByVal deleg As delegateDefColor, _
                                  ByVal TabStyle As DataGridTableStyle, _
                                  ByVal MapName As String, _
                                  ByVal HeadText As String, _
                                  ByVal ColWidth As Integer, _
                                  ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnF_Analiza0(deleg)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TCol.NullText = "---"
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub
    Public Sub TxtColoredColumnF_Analiza1(ByVal deleg As delegateDefColor, _
                                  ByVal TabStyle As DataGridTableStyle, _
                                  ByVal MapName As String, _
                                  ByVal HeadText As String, _
                                  ByVal ColWidth As Integer, _
                                  ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnF_Analiza1(deleg)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TCol.NullText = "---"
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub
    Public Sub TxtColoredColumnF_Analiza2(ByVal deleg As delegateDefColor, _
                              ByVal TabStyle As DataGridTableStyle, _
                              ByVal MapName As String, _
                              ByVal HeadText As String, _
                              ByVal ColWidth As Integer, _
                              ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnF_Analiza2(deleg)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TCol.NullText = "---"
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub
    Public Sub TxtColoredColumnF_Analiza3(ByVal deleg As delegateDefColor, _
                              ByVal TabStyle As DataGridTableStyle, _
                              ByVal MapName As String, _
                              ByVal HeadText As String, _
                              ByVal ColWidth As Integer, _
                              ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnF_Analiza3(deleg)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TCol.NullText = "---"
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub









    Public Sub NumColoredColumnF(ByVal deleg As delegateDefColor, _
                                      ByVal TabStyle As DataGridTableStyle, _
                                      ByVal MapName As String, _
                                      ByVal HeadText As String, _
                                      ByVal ColWidth As Integer, _
                                      ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment, _
                                      ByVal ColFormat As String)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnF(deleg)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TCol.Format = ColFormat
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub

    ' dlA potrzeb funkcji AnalizaZakupow
    Public Sub NumColoredColumnF_Analiza0(ByVal deleg As delegateDefColor, _
                                  ByVal TabStyle As DataGridTableStyle, _
                                  ByVal MapName As String, _
                                  ByVal HeadText As String, _
                                  ByVal ColWidth As Integer, _
                                  ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment, _
                                  ByVal ColFormat As String)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnF_Analiza0(deleg)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TCol.Format = ColFormat
        TCol.NullText = "---"
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub
    Public Sub NumColoredColumnF_Analiza1(ByVal deleg As delegateDefColor, _
                                  ByVal TabStyle As DataGridTableStyle, _
                                  ByVal MapName As String, _
                                  ByVal HeadText As String, _
                                  ByVal ColWidth As Integer, _
                                  ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment, _
                                  ByVal ColFormat As String)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnF_Analiza1(deleg)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TCol.Format = ColFormat
        TCol.NullText = "---"
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub
    Public Sub NumColoredColumnF_Analiza2(ByVal deleg As delegateDefColor, _
                                  ByVal TabStyle As DataGridTableStyle, _
                                  ByVal MapName As String, _
                                  ByVal HeadText As String, _
                                  ByVal ColWidth As Integer, _
                                  ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment, _
                                  ByVal ColFormat As String)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnF_Analiza2(deleg)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TCol.Format = ColFormat
        TCol.NullText = "---"
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub
    Public Sub NumColoredColumnF_Analiza3(ByVal deleg As delegateDefColor, _
                                  ByVal TabStyle As DataGridTableStyle, _
                                  ByVal MapName As String, _
                                  ByVal HeadText As String, _
                                  ByVal ColWidth As Integer, _
                                  ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment, _
                                  ByVal ColFormat As String)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnF_Analiza3(deleg)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TCol.Format = ColFormat
        TCol.NullText = "---"
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub

    '**************************************************
    '* kolorowany grid - ForeColor & BackColor
    '**************************************************

    Public Sub TxtColoredColumnBF(ByVal delegB As delegateDefColor, _
                                  ByVal delegF As delegateDefColor, _
                                      ByVal TabStyle As DataGridTableStyle, _
                                      ByVal MapName As String, _
                                      ByVal HeadText As String, _
                                      ByVal ColWidth As Integer, _
                                      ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnBF(delegB, delegF)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub

    Public Sub NumColoredColumnBF(ByVal delegB As delegateDefColor, _
                                  ByVal delegF As delegateDefColor, _
                                      ByVal TabStyle As DataGridTableStyle, _
                                      ByVal MapName As String, _
                                      ByVal HeadText As String, _
                                      ByVal ColWidth As Integer, _
                                      ByVal ColAllignment As System.Windows.Forms.HorizontalAlignment, _
                                      ByVal ColFormat As String)
        Dim TCol As New Softart.myColoredDataGridTextBoxColumnBF(delegB, delegF)
        TCol.MappingName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.Alignment = ColAllignment
        TCol.Format = ColFormat
        TabStyle.GridColumnStyles.Add(TCol)
    End Sub


    '**************************************************
    '* pobieranie danych z pola rekordu
    '**************************************************

    'Public Function GetTxtField(ByVal dGrid As System.Windows.Forms.DataGrid, ByVal FieldNo As Integer) As String
    '    'Return (IIf(IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)), "", Trim(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo))))
    '    If IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)) Then
    '        Return ""
    '    Else
    '        Return dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)
    '    End If
    'End Function

    'Public Function GetLogField(ByVal dGrid As System.Windows.Forms.DataGrid, ByVal FieldNo As Integer) As Boolean
    '    'Return (IIf(IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)), False, dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)))
    '    If IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)) Then
    '        Return False
    '    Else
    '        Return dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)
    '    End If
    'End Function

    'Public Function GetIntegerField(ByVal dGrid As System.Windows.Forms.DataGrid, ByVal FieldNo As Integer) As Integer
    '    'Return (IIf(IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)), 0, dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)))
    '    If IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)) Then
    '        Return 0
    '    Else
    '        Return dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)
    '    End If
    'End Function

    'Public Function GetDblField(ByVal dGrid As System.Windows.Forms.DataGrid, ByVal FieldNo As Integer) As Double
    '    'Return (IIf(IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)), 0, dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)))
    '    If IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)) Then
    '        Return 0
    '    Else
    '        Return dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)
    '    End If
    'End Function






    'Public Function GetTxtDRField(ByVal Row As DataRow, ByVal FieldN As String) As String
    '    'Return (IIf(IsDBNull(Row.Item(FieldN)), "", Row.Item(FieldN)))
    '    If IsDBNull(Row.Item(FieldN)) Then
    '        Return ""
    '    Else
    '        Return Trim(Row.Item(FieldN))
    '    End If
    'End Function

    'Public Function GetLogDRField(ByVal Row As DataRow, ByVal FieldN As String) As Boolean
    '    'Return (IIf(IsDBNull(Row.Item(FieldN)), False, Row.Item(FieldN)))
    '    If IsDBNull(Row.Item(FieldN)) Then
    '        Return False
    '    Else
    '        Return Trim(Row.Item(FieldN))
    '    End If
    'End Function

    'Public Function GetIntDRField(ByVal Row As DataRow, ByVal FieldN As String) As Integer
    '    'Return (IIf(IsDBNull(Row.Item(FieldN)), 0, Row.Item(FieldN)))
    '    If IsDBNull(Row.Item(FieldN)) Then
    '        Return 0
    '    Else
    '        Return Trim(Row.Item(FieldN))
    '    End If
    'End Function

    'Public Function GetDblDRField(ByVal Row As DataRow, ByVal FieldN As String) As Double
    '    'Return (IIf(IsDBNull(Row.Item(FieldN)), 0, Row.Item(FieldN)))
    '    If IsDBNull(Row.Item(FieldN)) Then
    '        Return 0
    '    Else
    '        Return Trim(Row.Item(FieldN))
    '    End If
    'End Function



    'Public Function GetTxtDRVField(ByVal Row As DataRowView, ByVal FieldN As String) As String
    '    'Return (IIf(IsDBNull(Row.Item(FieldN)), "", Row.Item(FieldN)))
    '    If IsDBNull(Row.Item(FieldN)) Then
    '        Return ""
    '    Else
    '        Return Trim(Row.Item(FieldN))
    '    End If
    'End Function

    'Public Function GetLogDRVField(ByVal Row As DataRowView, ByVal FieldN As String) As Boolean
    '    'Return (IIf(IsDBNull(Row.Item(FieldN)), False, Row.Item(FieldN)))
    '    If IsDBNull(Row.Item(FieldN)) Then
    '        Return False
    '    Else
    '        Return Trim(Row.Item(FieldN))
    '    End If
    'End Function

    'Public Function GetIntDRVField(ByVal Row As DataRowView, ByVal FieldN As String) As Integer
    '    'Return (IIf(IsDBNull(Row.Item(FieldN)), 0, Row.Item(FieldN)))
    '    If IsDBNull(Row.Item(FieldN)) Then
    '        Return 0
    '    Else
    '        Return Trim(Row.Item(FieldN))
    '    End If
    'End Function

    'Public Function GetDblDRVField(ByVal Row As DataRowView, ByVal FieldN As String) As Double
    '    'Return (IIf(IsDBNull(Row.Item(FieldN)), 0, Row.Item(FieldN)))
    '    If IsDBNull(Row.Item(FieldN)) Then
    '        Return 0
    '    Else
    '        Return Trim(Row.Item(FieldN))
    '    End If
    'End Function




    '**************************************************
    '* procedury testowania poprawnosci
    '**************************************************
    Public Sub TestLen(ByVal txtBox As TextBox, ByVal Nazwa As String, ByVal Dlugosc As Long)
        If Len(txtBox.Text) > Dlugosc Then
            MessageBox.Show("Pole " & Nazwa & vbCrLf & "nie mo¿e byæ d³u¿sze ni¿ " & Str(Dlugosc) & " znaków...", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
            txtBox.Text = Mid(txtBox.Text, 1, Dlugosc)
        Else
        End If
    End Sub

    Public Function TestIntLen(ByVal txtBox As TextBox, ByVal Nazwa As String, ByVal Dlugosc As Long) As Boolean
        Dim komu As String = ""
        Dim ii As Integer = 0
        Dim ch As String = ""
        Dim newtxtbox As String = ""
        Dim TxSelFrom As Integer = txtBox.SelectionStart
        Dim TxSelLen As Integer = txtBox.SelectionLength

        If Len(txtBox.Text) > Dlugosc Then
            komu = "Pole " & Nazwa & vbCrLf & "nie mo¿e byæ d³u¿sze ni¿ " & Str(Dlugosc) & " znaków..." & vbCrLf
        Else
        End If

        newtxtbox = ""
        If Len(txtBox.Text) > 0 Then
            For ii = 1 To IIf(Len(txtBox.Text) < Dlugosc, Len(txtBox.Text), Dlugosc)
                ch = Mid(txtBox.Text, ii, 1)
                Select Case ch
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                        newtxtbox &= ch
                    Case Else
                        komu &= "Pole " & Nazwa & " powinno sk³adaæ siê z max " & Str(Dlugosc) & " cyfr..." & vbCrLf
                End Select
            Next
        End If

        If Len(komu) > 0 Then
            MessageBox.Show(komu, "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)

            txtBox.Text = newtxtbox
            txtBox.SelectionStart = TxSelFrom
            txtBox.SelectionLength = TxSelLen
            Return (False)
        Else
            Return (True)
        End If
    End Function

    Public Sub TestCbxLen(ByVal txtBox As ComboBox, ByVal Nazwa As String, ByVal Dlugosc As Long)
        If Len(txtBox.Text) > Dlugosc Then
            MessageBox.Show("Pole " & Nazwa & vbCrLf & "nie mo¿e byæ d³u¿sze ni¿ " & Str(Dlugosc) & " znaków...", "Uwaga :", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
            txtBox.Text = Mid(txtBox.Text, 1, Dlugosc)
        Else
        End If
    End Sub

    Public Function TestNum(ByVal txtBox As TextBox, ByVal Nazwa As String) As Boolean
        ''Dim idl As Long = Len(txtBox.Text)
        ''Dim sdl As String = txtBox.Text
        'If Len(Trim(txtBox.Text)) = 0 Or IsNumeric(txtBox.Text) Then
        '    Return (True)
        'Else
        '    MessageBox.Show("Pole " & Nazwa & vbCrLf & "musi byæ liczb¹...", "Uwaga :", _
        '        System.Windows.Forms.MessageBoxButtons.OK, _
        '        MessageBoxIcon.Information, _
        '        MessageBoxDefaultButton.Button1)
        '    Return (False)
        'End If



        Dim poz As Integer = 0
        If Len(Trim(txtBox.Text)) = 0 Or
           txtBox.Text = "-" Or
           txtBox.Text = "+" Then
            Return (True)
        ElseIf IsNumeric(txtBox.Text) Then
            Return (True)
        Else
            poz = txtBox.SelectionStart
            txtBox.Text = MakeNum(txtBox.Text)
            txtBox.SelectionStart = poz

            If IsNumeric(txtBox.Text) Then
                Return (True)
            Else
                MessageBox.Show("Pole " & Nazwa & vbCrLf & "musi byæ liczb¹...", "Uwaga :",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1)
                Return (False)
            End If
        End If

    End Function

    Public Function MakeNum(ByVal Text As String) As String
        'Dim OutText As String = ""
        'Dim InText As String = Trim(Text)
        'Dim ii As Integer
        'Dim ch As String
        'If Len(InText) > 0 Then
        '    For ii = 1 To Len(InText)
        '        ch = Mid(InText, ii, 1)
        '        Select Case ch
        '            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ","
        '                OutText += ch
        '        End Select
        '    Next
        'Else
        '    OutText = "0"
        'End If
        'Return (OutText)

        Dim OutText As String = ""
        Dim InText As String = Trim(Text)
        Dim ii As Integer
        Dim ch As String
        If Len(InText) > 0 Then
            For ii = 1 To Len(InText)
                ch = Mid(InText, ii, 1)
                Select Case ch
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ","
                        OutText += ch
                    Case "."
                        OutText += ","

                End Select
            Next
        Else
            OutText = "0"
        End If
        Return (OutText)
    End Function




    '---------------
    ' zmiana separatora decymalnego w liczxbach jako string
    '---------------------------------------
    'Dim decimalSeparator As String = Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator

    Public Function KropNaPrzec(ByVal Kwo As String) As String
        'zamienia przecinek dziesiêtny na kropkê
        Dim pos As Integer
        pos = InStr(Kwo, ".")
        Do While pos > 0
            If pos = 1 Then
                Kwo = "," & Mid(Kwo, pos + 1)
            Else
                Kwo = Mid(Kwo, 1, pos - 1) & "," & Mid(Kwo, pos + 1)
            End If
            pos = InStr(Kwo, ".")
        Loop
        Return Kwo
    End Function

    Public Function PrzecNaKrop(ByVal Kwo As String) As String
        'zamienia przecinek dziesiêtny na kropkê
        Dim pos As Integer
        pos = InStr(Kwo, ",")
        Do While pos > 0
            If pos = 1 Then
                Kwo = "." & Mid(Kwo, pos + 1)
            Else
                Kwo = Mid(Kwo, 1, pos - 1) & "." & Mid(Kwo, pos + 1)
            End If
            pos = InStr(Kwo, ",")
        Loop
        Return Kwo
    End Function









    '********************************************************
    '**  procedury obslugi daty
    '********************************************************

    Public Function SysData() As String
        SysData = Right("0000" & Trim(Now.Year.ToString()), 4) & "-" & _
                  Right("00" & Trim(Now.Month.ToString()), 2) & "-" & _
                  Right("00" & Trim(Now.Day.ToString()), 2)
    End Function

    Public Function SysGodz() As String
        Return (Right("00" & Trim(Now.Hour.ToString()), 2) & ":" & _
                Right("00" & Trim(Now.Minute.ToString()), 2) & ":" & _
                Right("00" & Trim(Now.Second.ToString()), 2))
    End Function



    Public Function SysTime() As String
        'czas w formacie YYYY-MM-DD  GG:MM:SS.mmm
        Return (Right("0000" & Trim(Now.Year.ToString()), 4) & "-" &
                Right("00" & Trim(Now.Month.ToString()), 2) & "-" &
                Right("00" & Trim(Now.Day.ToString()), 2) & "  " &
                Right("00" & Trim(Now.Hour.ToString()), 2) & ":" &
                Right("00" & Trim(Now.Minute.ToString()), 2) & ":" &
                Right("00" & Trim(Now.Second.ToString()), 2) & "." &
                Right("00" & Trim(Now.Millisecond.ToString()), 2))
    End Function








    Public Function IleSekundMinelo(ByVal ChwilaOd As Date, ByVal ChwilaDo As Date) As Long
        Dim SekOd As Long = (Hour(ChwilaOd) * 3600) + (Minute(ChwilaOd) * 60) + Second(ChwilaOd)
        Dim SekDo As Long = (Hour(ChwilaDo) * 3600) + (Minute(ChwilaDo) * 60) + Second(ChwilaDo)
        Return (SekDo - SekOd)
    End Function

    'termin zaplaty
    Public Function WyliczDate(ByVal Data1 As String, ByVal IleDni As Long) As String
        Dim RRRR As Integer = Val(Mid(Data1, 1, 4))
        Dim MM As Integer = Val(Mid(Data1, 6, 2))
        Dim DD As Integer = Val(Mid(Data1, 9, 2))
        Return (JakiToDzien(IleDniOd20000101(RRRR, MM, DD) + IleDni))
    End Function

    ' wylicza ile dni minelo do 2000.01.01
    Private Function IleDniOd20000101(ByVal YY As Integer, ByVal MM As Integer, ByVal DD As Integer) As Long
        Dim IleDni As Long = 0
        Dim ii As Integer

        If YY > 2000 Then
            For ii = 2000 To YY - 1
                If ii / 100 = Int(ii / 100) Then
                    IleDni += 366   'przestepny
                ElseIf ii / 4 = Int(ii / 4) Then
                    IleDni += 366   'przestepny
                Else
                    IleDni += 365
                End If
            Next ii
        End If

        If MM > 1 Then
            For ii = 1 To MM - 1
                Select Case ii
                    Case 1
                        IleDni += 31
                    Case 2
                        If YY / 100 = Int(YY / 100) Then
                            IleDni += 29   'przestepny
                        ElseIf YY / 4 = Int(YY / 4) Then
                            IleDni += 29   'przestepny
                        Else
                            IleDni += 28
                        End If
                    Case 3
                        IleDni += 31
                    Case 4
                        IleDni += 30
                    Case 5
                        IleDni += 31
                    Case 6
                        IleDni += 30
                    Case 7
                        IleDni += 31
                    Case 8
                        IleDni += 31
                    Case 9
                        IleDni += 30
                    Case 10
                        IleDni += 31
                    Case 11
                        IleDni += 30
                    Case 12
                        IleDni += 31
                End Select
            Next ii
        End If

        IleDni += DD
        Return (IleDni)
    End Function

    'wylicza date na podstawie ilosci dni od 2000.01.01
    Private Function JakiToDzien(ByVal IleDniOd As Long) As String
        Dim RRRR As String
        Dim MM As String
        Dim DD As String

        Dim rok As Integer
        Dim miesiac As Integer
        Dim dzien As Integer
        Dim iledni As Integer

        'wyliczamy rok...
        rok = 2000
        iledni = 366
        Do While IleDniOd > iledni
            IleDniOd -= iledni
            rok += 1
            If rok / 100 = Int(rok / 100) Then
                iledni = 366   'przestepny
            ElseIf rok / 4 = Int(rok / 4) Then
                iledni = 366   'przestepny
            Else
                iledni = 365
            End If
        Loop

        'wyliczamy miesiac...
        miesiac = 1
        iledni = 31
        Do While IleDniOd > iledni
            IleDniOd -= iledni
            miesiac += 1
            Select Case miesiac
                Case 1
                    iledni = 31
                Case 2
                    If rok / 100 = Int(rok / 100) Then
                        iledni = 29   'przestepny
                    ElseIf rok / 4 = Int(rok / 4) Then
                        iledni = 29   'przestepny
                    Else
                        iledni = 28
                    End If
                Case 3
                    iledni = 31
                Case 4
                    iledni = 30
                Case 5
                    iledni = 31
                Case 6
                    iledni = 30
                Case 7
                    iledni = 31
                Case 8
                    iledni = 31
                Case 9
                    iledni = 30
                Case 10
                    iledni = 31
                Case 11
                    iledni = 30
                Case 12
                    iledni = 31
            End Select
        Loop

        'wyliczam dzien...
        dzien = IleDniOd

        RRRR = "0000" & Trim(rok.ToString)
        MM = "00" & Trim(miesiac.ToString)
        DD = "00" & Trim(dzien.ToString)
        Return (Mid(RRRR, Len(RRRR) - 4 + 1, 4) & "-" & Mid(MM, Len(MM) - 2 + 1, 2) & "-" & Mid(DD, Len(DD) - 2 + 1, 2))
    End Function


    Public Function WyliczMiesiac(ByVal BiezMiesiac As String, _
                                  ByVal RoznicaMiesiecy As Integer) As String
        Dim RRRRnum As Integer = GetNumFromText(Mid(BiezMiesiac, 1, 4))
        Dim MMnum As Integer = GetNumFromText(Mid(BiezMiesiac, 6, 2))

        MMnum = MMnum + RoznicaMiesiecy
        Do While MMnum < 1
            MMnum += 12
            RRRRnum -= 1
        Loop
        Do While MMnum > 12
            MMnum -= 12
            RRRRnum += 1
        Loop
        'Miesiac w formacie YYYY-MM
        Return Right("0000" & Trim(RRRRnum.ToString()), 4) & "-" & _
               Right("00" & Trim(MMnum.ToString()), 2)
    End Function

    Public Function WyliczRok(ByVal BiezRok As String, _
                              ByVal RoznicaLat As Integer) As String
        Dim RRRRnum As Integer = GetNumFromText(Mid(BiezRok, 1, 4))
        RRRRnum += RoznicaLat
        'Miesiac w formacie YYYY
        Return Right("0000" & Trim(RRRRnum.ToString()), 4)
    End Function


    '********************************************************
    '**  procedury przeliczania
    '********************************************************

    Public Function cm2pts(ByVal cm As Single) As Single
        Return (CType(cm * (100 / 2.54), Single))
    End Function
    Public Function mm2pts(ByVal mm As Single) As Single
        Return (CType((mm * 10) / 2.54, Single))
    End Function

    Public Function pts2mm(ByVal pts As Single) As Single
        Return (CType((pts * 2.54) / 10, Single))
    End Function

    '******************************************
    '** procedury obslugi wyjatkow
    '******************************************

    Public Sub ViewInLog(ByVal e As Exception, ByVal ControlName As String, ByVal AdditionalInfo As String)
        Dim ExecutingApp As System.Reflection.Assembly
        ExecutingApp = System.Reflection.Assembly.GetExecutingAssembly()
        Dim Name As System.Reflection.AssemblyName
        Name = ExecutingApp.GetName()

        ' nazwa = MAGAZYNRRDD.log
        Dim m_LogFile As String
        m_LogFile = Trim(System.Windows.Forms.Application.StartupPath) & "\" & _
                    Trim(Name.Name) & Right("00" & Trim(Now.Year.ToString()), 2) & _
                    Right("00" & Trim(Now.Month.ToString()), 2) & ".LOG"

        ' Write the entry to the log file.
        Dim FileNum As Integer
        FileNum = FreeFile()    'kolejny nr otwartego zbioru
        FileOpen(FileNum, m_LogFile, OpenMode.Append)

        PrintLine(FileNum, Now.ToString() & " [" & ControlName & "]")
        PrintLine(FileNum, e.Message)
        If Not AdditionalInfo Is Nothing Then
            PrintLine(FileNum, AdditionalInfo)
        End If
        FileClose(FileNum)
        '---------------------------------------
        ViewOnScreen(e, AdditionalInfo)
    End Sub

    Public Sub ViewOnScreen(ByVal e As Exception, ByVal AdditionalInfo As String)
        MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & e.Message & vbCrLf & _
            IIf(Not AdditionalInfo Is Nothing, AdditionalInfo, ""), _
            "Uwaga :", _
            System.Windows.Forms.MessageBoxButtons.OK, _
            MessageBoxIcon.Information, _
            MessageBoxDefaultButton.Button1)
    End Sub

    '******************************************
    '** filtr zlozony
    '******************************************

    Public Function BudujFiltr(ByVal Globalny As String, ByVal Lokalny As String) As String
        If Len(Globalny) > 0 Then
            If Len(Lokalny) > 0 Then
                Return (Globalny & " And " & Lokalny)
            Else
                Return (Globalny)
            End If
        Else
            If Len(Lokalny) > 0 Then
                Return (Lokalny)
            Else
                Return ("")
            End If
        End If
    End Function

    '******************************************
    '** zaznaczenie edytowanego pola
    '******************************************

    Public Sub StartEdycjiTxt(ByVal Pole As TextBox)
        Pole.BackColor = SoftartCursorBackColor
        Pole.ForeColor = SoftartCursorForeColor
    End Sub

    Public Sub KoniecEdycjiTxt(ByVal Pole As TextBox)
        Pole.BackColor = SoftartEditableBackColor
        Pole.ForeColor = SoftartEditableForeColor
    End Sub

    Public Sub StartEdycjiCbx(ByVal Pole As ComboBox)
        Pole.BackColor = SoftartCursorBackColor
        Pole.ForeColor = SoftartCursorForeColor
    End Sub
    Public Sub KoniecEdycjiCbx(ByVal Pole As ComboBox)
        Pole.BackColor = SoftartEditableBackColor
        Pole.ForeColor = SoftartEditableForeColor
    End Sub

    Public Sub StartEdycjiChb(ByVal Pole As CheckBox)
        Pole.BackColor = SoftartCursorBackColor
        'Pole.ForeColor = SoftartCursorForeColor
    End Sub
    Public Sub KoniecEdycjiChb(ByVal Pole As CheckBox)
        Pole.BackColor = SoftartEditableBackColor
        'Pole.ForeColor = SoftartEditableForeColor
    End Sub

    Public Sub StartEdycjiLst(ByVal Pole As ListBox)
        Pole.BackColor = SoftartCursorBackColor
        Pole.ForeColor = SoftartCursorForeColor
    End Sub
    Public Sub KoniecEdycjiLst(ByVal Pole As ListBox)
        Pole.BackColor = SoftartEditableBackColor
        Pole.ForeColor = SoftartEditableForeColor
    End Sub

    Public Sub StartEdycjiRab(ByVal Pole As RadioButton)
        Pole.BackColor = SoftartCursorBackColor
        'Pole.ForeColor = SoftartCursorForeColor
    End Sub
    Public Sub KoniecEdycjiRab(ByVal Pole As RadioButton)
        Pole.BackColor = SoftartEditableBackColor
        'Pole.ForeColor = SoftartEditableForeColor
    End Sub

    '***********************
    ' testuje format liczby i daty w komputerze
    '***********************

    Public Function DecimalSeparator() As String
        'DecimalSeparator = Mid$(1 / 2, 2, 1)
        Return (Mid$(1 / 2, 2, 1))
    End Function

    Public Function GetRegionalOption() As String
        Dim Format As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CurrentCulture

        '' Displays the name of the CurrentCulture of the current thread.
        'Console.WriteLine("CurrentCulture is {0}.", CultureInfo.CurrentCulture.Name)
        '' Changes the CurrentCulture of the current thread to th-TH.
        'Thread.CurrentThread.CurrentCulture = New CultureInfo("th-TH", False)
        'Console.WriteLine("CurrentCulture is now {0}.", CultureInfo.CurrentCulture.Name)
        '' Displays the name of the CurrentUICulture of the current thread.
        'Console.WriteLine("CurrentUICulture is {0}.", CultureInfo.CurrentUICulture.Name)
        '' Changes the CurrentUICulture of the current thread to ja-JP.
        'Thread.CurrentThread.CurrentUICulture = New CultureInfo("ja-JP", False)
        'Console.WriteLine("CurrentUICulture is now {0}.", CultureInfo.CurrentUICulture.Name)

        Return Format.Name
    End Function






    '***********************
    ' dzien tygodnia na podstawie daty
    '***********************

    Public Function DzienTygodnia(ByVal pData As String) As String
        Dim AktData As Date = Nothing
        Dim DOW As Integer = 0
        Dim RetDay As String = ""

        If Len(pData) > 0 Then
            AktData = pData
            DOW = AktData.DayOfWeek
            Select Case DOW
                Case 0
                    RetDay = "Niedziela"
                Case 1
                    RetDay = "Poniedzia³ek"
                Case 2
                    RetDay = "Wtorek"
                Case 3
                    RetDay = "Œroda"
                Case 4
                    RetDay = "Czwartek"
                Case 5
                    RetDay = "Pi¹tek"
                Case 6
                    RetDay = "Sobota"
            End Select
        End If
        Return RetDay
    End Function



    Public Function InnyMiesiac(ByVal Ktory As Integer) As String
        Dim AktData As String = ""
        Dim Rok As Integer = 0
        Dim Mies As Integer = 0
        Dim i As Integer = 0

        If Len(Trim(Par_OstOkresAnalizy)) = 0 Then
            AktData = SysData()
        Else
            AktData = NastMiesiac(Trim(Par_OstOkresAnalizy) & "-01")
        End If
        Rok = CInt(Mid(AktData, 1, 4))
        Mies = CInt(Mid(AktData, 6, 2))

        If Ktory > 0 Then
            For i = 1 To Ktory Step 1
                Mies += 1
                If Mies = 13 Then
                    Mies = 1
                    Rok += 1
                End If
            Next
        ElseIf Ktory = 0 Then
        ElseIf Ktory < 0 Then
            For i = -1 To Ktory Step -1
                Mies -= 1
                If Mies = 0 Then
                    Mies = 12
                    Rok -= 1
                End If
            Next
        End If
        Return Microsoft.VisualBasic.Right(Trim(Str(100 + Mies)), 2) & "." & Microsoft.VisualBasic.Right(Trim(Str(10000 + Rok)), 4)
    End Function

    Public Function InnyMiesiac(ByVal MiesBaz As String, _
                                ByVal Ktory As Integer) As String
        Dim AktData As String = MiesBaz
        Dim Rok As Integer = 0
        Dim Mies As Integer = 0
        Dim i As Integer = 0

        If OnlyDigits(Mid(AktData, 1, 4)) And OnlyDigits(Mid(AktData, 6, 2)) Then
            Rok = CInt(Mid(AktData, 1, 4))
            Mies = CInt(Mid(AktData, 6, 2))
        Else
            'bledny format parametru
            Rok = CInt(Mid(SysData, 1, 4))
            Mies = CInt(Mid(SysData, 6, 2))
        End If

        If Ktory > 0 Then
            For i = 1 To Ktory Step 1
                Mies += 1
                If Mies = 13 Then
                    Mies = 1
                    Rok += 1
                End If
            Next
        ElseIf Ktory = 0 Then
        ElseIf Ktory < 0 Then
            For i = -1 To Ktory Step -1
                Mies -= 1
                If Mies = 0 Then
                    Mies = 12
                    Rok -= 1
                End If
            Next
        End If
        Return Microsoft.VisualBasic.Right(Trim(Str(100 + Mies)), 2) & "." & Microsoft.VisualBasic.Right(Trim(Str(10000 + Rok)), 4)
    End Function



    Private Function OnlyDigits(ByVal SrcText As String) As Boolean
        Dim Ret As Boolean = True
        Dim ch As String = ""
        Dim i As Integer = 0
        If Len(SrcText) > 0 Then
            For i = 1 To Len(SrcText)
                ch = Mid(SrcText, i, 1)
                Select Case ch
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                        'OK - cyfry
                    Case Else
                        'to nie jest cyfra...
                        Ret = False
                        Exit For
                End Select
            Next
        End If
        Return Ret
    End Function


    '************************************
    ' miesiac do analizy
    '************************************

    Public Function MiesDoAnalizy(ByVal Ktory As Integer) As String
        Dim AktData As String = SysData()
        Dim Rok As Integer = CInt(Mid(AktData, 1, 4))
        Dim Mies As Integer = CInt(Mid(AktData, 6, 2))
        Dim i As Integer = 0

        If Ktory > 0 Then
            For i = 1 To Ktory Step 1
                Mies += 1
                If Mies = 13 Then
                    Mies = 1
                    Rok += 1
                End If
            Next
        Else
            For i = -1 To Ktory Step -1
                Mies -= 1
                If Mies = 0 Then
                    Mies = 12
                    Rok -= 1
                End If
            Next
        End If
        Return Microsoft.VisualBasic.Right(Trim(Str(10000 + Rok)), 4) & "-" & Microsoft.VisualBasic.Right(Trim(Str(100 + Mies)), 2)
    End Function

    Public Function MiesDoAnalizy(ByVal OdMiesiaca As String, _
                                  ByVal Ktory As Integer) As String
        Dim Rok As Integer = CInt(Mid(OdMiesiaca, 1, 4))
        Dim Mies As Integer = CInt(Mid(OdMiesiaca, 6, 2))
        Dim i As Integer = 0

        If Ktory > 0 Then
            For i = 1 To Ktory Step 1
                Mies += 1
                If Mies = 13 Then
                    Mies = 1
                    Rok += 1
                End If
            Next
        Else
            For i = -1 To Ktory Step -1
                Mies -= 1
                If Mies = 0 Then
                    Mies = 12
                    Rok -= 1
                End If
            Next
        End If
        Return Microsoft.VisualBasic.Right(Trim(Str(10000 + Rok)), 4) & "-" & Microsoft.VisualBasic.Right(Trim(Str(100 + Mies)), 2)
    End Function


    Public Function PopMiesiac(ByVal Data As String) As String
        Dim AktData As String = Data
        Dim Rok As Integer = CInt(Mid(AktData, 1, 4))
        Dim Mies As Integer = CInt(Mid(AktData, 6, 2))
        Dim Dzien As Integer = CInt(Mid(AktData, 9, 2))
        Dim i As Integer = 0

        Mies -= 1
        If Mies = 0 Then
            Mies = 12
            Rok -= 1
        End If

        Return Microsoft.VisualBasic.Right(Trim(Str(10000 + Rok)), 4) & "-" & Microsoft.VisualBasic.Right(Trim(Str(100 + Mies)), 2) & "-" & Microsoft.VisualBasic.Right(Trim(Str(100 + Dzien)), 2)
    End Function

    Public Function NastMiesiac(ByVal Data As String) As String
        Dim AktData As String = Data
        Dim Rok As Integer = CInt(Mid(AktData, 1, 4))
        Dim Mies As Integer = CInt(Mid(AktData, 6, 2))
        Dim Dzien As Integer = CInt(Mid(AktData, 9, 2))
        Dim i As Integer = 0

        Mies += 1
        If Mies = 13 Then
            Mies = 1
            Rok += 1
        End If

        Return Microsoft.VisualBasic.Right(Trim(Str(10000 + Rok)), 4) & "-" & Microsoft.VisualBasic.Right(Trim(Str(100 + Mies)), 2) & "-" & Microsoft.VisualBasic.Right(Trim(Str(100 + Dzien)), 2)
    End Function


    Public Function PopRok(ByVal Data As String) As String
        Dim AktData As String = Data
        Dim Rok As Integer = CInt(Mid(AktData, 1, 4))
        Dim Mies As Integer = CInt(Mid(AktData, 6, 2))
        Dim Dzien As Integer = CInt(Mid(AktData, 9, 2))
        Dim i As Integer = 0
        Rok -= 1
        Return Microsoft.VisualBasic.Right(Trim(Str(10000 + Rok)), 4) & "-" & Microsoft.VisualBasic.Right(Trim(Str(100 + Mies)), 2) & "-" & Microsoft.VisualBasic.Right(Trim(Str(100 + Dzien)), 2)
    End Function

    Public Function NastRok(ByVal Data As String) As String
        Dim AktData As String = Data
        Dim Rok As Integer = CInt(Mid(AktData, 1, 4))
        Dim Mies As Integer = CInt(Mid(AktData, 6, 2))
        Dim Dzien As Integer = CInt(Mid(AktData, 9, 2))
        Dim i As Integer = 0
        Rok += 1
        Return Microsoft.VisualBasic.Right(Trim(Str(10000 + Rok)), 4) & "-" & Microsoft.VisualBasic.Right(Trim(Str(100 + Mies)), 2) & "-" & Microsoft.VisualBasic.Right(Trim(Str(100 + Dzien)), 2)
    End Function


    '********************************************
    ' pozwala ustawic atrybuty tak, ¿eby kontrolke nie mozna bylo edytowac
    '********************************************

    Public Sub NieEdytowac(ByVal Cont As Control)
        'uniemo¿liwiæ edycje pól
        Dim txb As TextBox = Nothing
        Dim cbx As ComboBox = Nothing
        Dim rbt As RadioButton = Nothing
        Dim chk As CheckBox = Nothing
        Dim btn As Button = Nothing
        Dim mnu As ContextMenuStrip = Nothing

        Dim ctrl As Control
        For Each ctrl In Cont.Controls
            If TypeOf (ctrl) Is TextBox Then
                txb = ctrl
                txb.ReadOnly = True
                txb.Enabled = False
            ElseIf TypeOf (ctrl) Is ComboBox Then
                cbx = ctrl
                cbx.Enabled = False
            ElseIf TypeOf (ctrl) Is RadioButton Then
                rbt = ctrl
                rbt.Enabled = False
            ElseIf TypeOf (ctrl) Is CheckBox Then
                chk = ctrl
                chk.Enabled = False
            ElseIf TypeOf (ctrl) Is Button Then
                If ctrl.Name = "cmdStop" Then
                    btn = ctrl
                    btn.Enabled = True
                Else
                    btn = ctrl
                    btn.Enabled = False
                End If

            ElseIf TypeOf (ctrl) Is System.Windows.Forms.ContextMenuStrip Then
                mnu = ctrl
                mnu.Enabled = False

            ElseIf TypeOf (ctrl) Is System.Windows.Forms.Panel Then
                'wywolujemy rekursyjnie...
                NieEdytowac(ctrl)
            ElseIf TypeOf (ctrl) Is System.Windows.Forms.SplitContainer Then
                'wywolujemy rekursyjnie...
                NieEdytowac(ctrl)
            End If
        Next
    End Sub



End Module
