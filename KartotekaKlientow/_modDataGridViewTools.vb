Module _modDataGridViewTools

    'With DataGridView1
    '    ' Set DataGridView Combo Column for CarID field
    '    Dim ColumnCar As New DataGridViewComboColumn
    '    ' DataGridView Combo ValueMember field has name "CarID"
    '    ' DataGridView Combo DisplayMember field has name "Car"
    '    With ColumnCar
    '        .DataPropertyName =3D "CarID"
    '        .HeaderText =3D "Car Name"
    '        .Width =3D 80
    '        ' Bind ColumnCar to Cars table
    '        .box.DataSource =3D ds.Tables("Cars")
    '        .box.ValueMember =3D "CarID"
    '        .box.DisplayMember =3D "Car"
    '    End With
    '    .Columns.Add(ColumnCar)
    'End With




    '**************************************************
    '* definiowanie kolumny EDYTOWALNE w tabeli
    '**************************************************

    Public Sub TxtColumnViewEdi(ByVal DGrid As DataGridView,
                      ByVal MapName As String,
                      ByVal HeadText As String,
                      ByVal ColWidth As Integer,
                      ByVal ColAllignment As System.Windows.Forms.DataGridViewContentAlignment,
                      ByVal ColVisible As Boolean,
                      ByVal MaxInLength As Integer)
        Dim TCol As New System.Windows.Forms.DataGridViewTextBoxColumn
        TCol.DataPropertyName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.DefaultCellStyle.Alignment = ColAllignment
        TCol.DefaultCellStyle.NullValue = ""
        TCol.ReadOnly = False
        TCol.MaxInputLength = MaxInLength
        'TCol.Tag = DGrid.Columns.Count
        TCol.Tag = ColWidth
        '-------------------------------
        DGrid.Columns.Add(TCol)
        DGrid.Columns(DGrid.Columns.Count() - 1).Visible = ColVisible
        DGrid.Columns(DGrid.Columns.Count() - 1).ValueType = System.Type.GetType("STRING")
        'DGrid.Columns(MapName).Visible = ColVisible
        '-------------------------------
    End Sub



    '**************************************************
    '* definiowanie kolumny w tabeli - NOWE KOLUMNY
    '**************************************************

    Public Sub TxtColumnView(ByVal DGrid As DataGridView,
                      ByVal MapName As String,
                      ByVal HeadText As String,
                      ByVal ColWidth As Integer,
                      ByVal ColAllignment As System.Windows.Forms.DataGridViewContentAlignment,
                      ByVal ColVisible As Boolean)
        Dim TCol As New System.Windows.Forms.DataGridViewTextBoxColumn
        TCol.DataPropertyName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.DefaultCellStyle.Alignment = ColAllignment
        TCol.DefaultCellStyle.NullValue = ""
        TCol.ReadOnly = True
        'TCol.Tag = DGrid.Columns.Count
        TCol.Tag = ColWidth
        '-------------------------------
        DGrid.Columns.Add(TCol)
        DGrid.Columns(DGrid.Columns.Count() - 1).Visible = ColVisible
        DGrid.Columns(DGrid.Columns.Count() - 1).ValueType = System.Type.GetType("STRING")
        'DGrid.Columns(MapName).Visible = ColVisible
        '-------------------------------
        'DGrid.Columns.Add(MapName, HeadText)
        'DGrid.Columns(DGrid.Columns.Count() - 1).Width = ColWidth
        'DGrid.Columns(DGrid.Columns.Count() - 1).Visible = ColVisible
        'DGrid.Columns(DGrid.Columns.Count() - 1).DataPropertyName = MapName
        'DGrid.Columns(DGrid.Columns.Count() - 1).Name = HeadText
        'DGrid.Columns(DGrid.Columns.Count() - 1).DefaultCellStyle.Alignment = ColAllignment
        'DGrid.Columns(DGrid.Columns.Count() - 1).DefaultCellStyle.NullValue = ""
        DGrid.Columns(DGrid.Columns.Count() - 1).DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName, Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
    End Sub

    Public Sub BigTxtColumnView(ByVal DGrid As DataGridView,
                  ByVal MapName As String,
                  ByVal HeadText As String,
                  ByVal ColWidth As Integer,
                  ByVal ColAllignment As System.Windows.Forms.DataGridViewContentAlignment,
                  ByVal ColVisible As Boolean)
        Dim TCol As New System.Windows.Forms.DataGridViewTextBoxColumn
        TCol.DataPropertyName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.DefaultCellStyle.Alignment = ColAllignment
        TCol.DefaultCellStyle.NullValue = ""
        TCol.ReadOnly = True
        'TCol.Tag = DGrid.Columns.Count
        TCol.Tag = ColWidth
        '-------------------------------
        DGrid.Columns.Add(TCol)
        DGrid.Columns(DGrid.Columns.Count() - 1).Visible = ColVisible
        DGrid.Columns(DGrid.Columns.Count() - 1).ValueType = System.Type.GetType("STRING")
        'DGrid.Columns(MapName).Visible = ColVisible
        '-------------------------------
        'DGrid.Columns.Add(MapName, HeadText)
        'DGrid.Columns(DGrid.Columns.Count() - 1).Width = ColWidth
        'DGrid.Columns(DGrid.Columns.Count() - 1).Visible = ColVisible
        'DGrid.Columns(DGrid.Columns.Count() - 1).DataPropertyName = MapName
        'DGrid.Columns(DGrid.Columns.Count() - 1).Name = HeadText
        'DGrid.Columns(DGrid.Columns.Count() - 1).DefaultCellStyle.Alignment = ColAllignment
        'DGrid.Columns(DGrid.Columns.Count() - 1).DefaultCellStyle.NullValue = ""
        DGrid.Columns(DGrid.Columns.Count() - 1).DefaultCellStyle.Font = New System.Drawing.Font(Softart_FormFontName,
                                                                                                 1.3 * Softart_FormFontSize,
                                                             System.Drawing.FontStyle.Regular,
                                                             System.Drawing.GraphicsUnit.Point,
                                                             CType(238, Byte))
    End Sub

    Public Sub NumColumnViewStanMag(ByVal DGrid As DataGridView,
              ByVal MapName As String,
              ByVal HeadText As String,
              ByVal ColWidth As Integer,
              ByVal ColAllignment As System.Windows.Forms.DataGridViewContentAlignment,
              ByVal ColFormat As String,
              ByVal ColVisible As Boolean)
        Dim TCol As New System.Windows.Forms.DataGridViewTextBoxColumn
        TCol.DataPropertyName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.DefaultCellStyle.Alignment = ColAllignment
        TCol.DefaultCellStyle.Format = ColFormat
        TCol.DefaultCellStyle.NullValue = "---"
        TCol.ReadOnly = True
        'TCol.Tag = DGrid.Columns.Count
        TCol.Tag = ColWidth
        '-------------------------------
        DGrid.Columns.Add(TCol)
        DGrid.Columns(DGrid.Columns.Count() - 1).Visible = ColVisible
        'DGrid.Columns(MapName).Visible = ColVisible
    End Sub

    Public Sub NumColumnView(ByVal DGrid As DataGridView,
          ByVal MapName As String,
          ByVal HeadText As String,
          ByVal ColWidth As Integer,
          ByVal ColAllignment As System.Windows.Forms.DataGridViewContentAlignment,
          ByVal ColFormat As String,
          ByVal ColVisible As Boolean)
        Dim TCol As New System.Windows.Forms.DataGridViewTextBoxColumn
        TCol.DataPropertyName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.DefaultCellStyle.Alignment = ColAllignment
        TCol.DefaultCellStyle.Format = ColFormat
        TCol.DefaultCellStyle.NullValue = "0"
        'TCol.ValueType = GetType(Double)
        TCol.ReadOnly = True
        'TCol.Tag = DGrid.Columns.Count
        TCol.Tag = ColWidth
        '-------------------------------
        DGrid.Columns.Add(TCol)
        DGrid.Columns(DGrid.Columns.Count() - 1).Visible = ColVisible
        DGrid.Columns(DGrid.Columns.Count() - 1).ValueType = System.Type.GetType("STRING")
        'DGrid.Columns(MapName).Visible = ColVisible
    End Sub


    Public Sub LogColumnView(ByVal DGrid As DataGridView,
                  ByVal MapName As String,
                  ByVal HeadText As String,
                  ByVal ColWidth As Integer,
                  ByVal ColVisible As Boolean)
        Dim TCol As New DataGridViewCheckBoxColumn
        TCol.DataPropertyName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        TCol.DefaultCellStyle.NullValue = False
        TCol.ReadOnly = True
        'TCol.Tag = DGrid.Columns.Count
        TCol.Tag = ColWidth
        '-------------------------------
        DGrid.Columns.Add(TCol)
        DGrid.Columns(DGrid.Columns.Count() - 1).Visible = ColVisible
        DGrid.Columns(DGrid.Columns.Count() - 1).ValueType = System.Type.GetType("CheckBox")
        'DGrid.Columns(MapName).Visible = ColVisible
    End Sub






    '**************************************************
    '* definiowanie kolumny w tabeli - NOWE KOLUMNY gdy kolumny DATASOURCE niechronologiczne
    '**************************************************

    Public Sub TxtColumnView(ByVal DGrid As DataGridView,
                      ByVal MapName As String,
                      ByVal HeadText As String,
                      ByVal ColWidth As Integer,
                      ByVal ColAllignment As System.Windows.Forms.DataGridViewContentAlignment,
                      ByVal ColVisible As Boolean,
                      ByVal ColDataSource As Integer)
        Dim TCol As New System.Windows.Forms.DataGridViewTextBoxColumn
        TCol.DataPropertyName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.DefaultCellStyle.Alignment = ColAllignment
        TCol.DefaultCellStyle.NullValue = ""
        TCol.ReadOnly = True
        TCol.Tag = ColDataSource
        '-------------------------------
        DGrid.Columns.Add(TCol)
        DGrid.Columns(DGrid.Columns.Count() - 1).Visible = ColVisible
        DGrid.Columns(DGrid.Columns.Count() - 1).ValueType = System.Type.GetType("STRING")
        'DGrid.Columns(MapName).Visible = ColVisible
        '-------------------------------
        'DGrid.Columns.Add(MapName, HeadText)
        'DGrid.Columns(DGrid.Columns.Count() - 1).Width = ColWidth
        'DGrid.Columns(DGrid.Columns.Count() - 1).Visible = ColVisible
        'DGrid.Columns(DGrid.Columns.Count() - 1).DataPropertyName = MapName
        'DGrid.Columns(DGrid.Columns.Count() - 1).Name = HeadText
        'DGrid.Columns(DGrid.Columns.Count() - 1).DefaultCellStyle.Alignment = ColAllignment
        'DGrid.Columns(DGrid.Columns.Count() - 1).DefaultCellStyle.NullValue = ""
    End Sub

    Public Sub NumColumnViewStanMag(ByVal DGrid As DataGridView,
              ByVal MapName As String,
              ByVal HeadText As String,
              ByVal ColWidth As Integer,
              ByVal ColAllignment As System.Windows.Forms.DataGridViewContentAlignment,
              ByVal ColFormat As String,
                      ByVal ColVisible As Boolean,
                      ByVal ColDataSource As Integer)
        Dim TCol As New System.Windows.Forms.DataGridViewTextBoxColumn
        TCol.DataPropertyName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.DefaultCellStyle.Alignment = ColAllignment
        TCol.DefaultCellStyle.Format = ColFormat
        TCol.DefaultCellStyle.NullValue = "---"
        TCol.ReadOnly = True
        TCol.Tag = ColDataSource
        '-------------------------------
        DGrid.Columns.Add(TCol)
        DGrid.Columns(DGrid.Columns.Count() - 1).Visible = ColVisible
        'DGrid.Columns(MapName).Visible = ColVisible
    End Sub

    Public Sub NumColumnView(ByVal DGrid As DataGridView,
          ByVal MapName As String,
          ByVal HeadText As String,
          ByVal ColWidth As Integer,
          ByVal ColAllignment As System.Windows.Forms.DataGridViewContentAlignment,
          ByVal ColFormat As String,
                      ByVal ColVisible As Boolean,
                      ByVal ColDataSource As Integer)
        Dim TCol As New System.Windows.Forms.DataGridViewTextBoxColumn
        TCol.DataPropertyName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.DefaultCellStyle.Alignment = ColAllignment
        TCol.DefaultCellStyle.Format = ColFormat
        TCol.DefaultCellStyle.NullValue = "0"
        'TCol.ValueType = GetType(Double)
        TCol.ReadOnly = True
        TCol.Tag = ColDataSource
        '-------------------------------
        DGrid.Columns.Add(TCol)
        DGrid.Columns(DGrid.Columns.Count() - 1).Visible = ColVisible
        DGrid.Columns(DGrid.Columns.Count() - 1).ValueType = System.Type.GetType("STRING")
        'DGrid.Columns(MapName).Visible = ColVisible
    End Sub


    Public Sub LogColumnView(ByVal DGrid As DataGridView,
                  ByVal MapName As String,
                  ByVal HeadText As String,
                  ByVal ColWidth As Integer,
                      ByVal ColVisible As Boolean,
                      ByVal ColDataSource As Integer)
        Dim TCol As New DataGridViewCheckBoxColumn
        TCol.DataPropertyName = MapName
        TCol.HeaderText = HeadText
        TCol.Width = ColWidth
        TCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        TCol.DefaultCellStyle.NullValue = False
        TCol.ReadOnly = True
        TCol.Tag = ColDataSource
        '-------------------------------
        DGrid.Columns.Add(TCol)
        DGrid.Columns(DGrid.Columns.Count() - 1).Visible = ColVisible
        DGrid.Columns(DGrid.Columns.Count() - 1).ValueType = System.Type.GetType("CheckBox")
        'DGrid.Columns(MapName).Visible = ColVisible
    End Sub







    ''**************************************************
    ''* definiowanie kolumny w tabeli - KOLUMNY JU¯ ISTNIEJ¥
    ''**************************************************

    'Public Sub TxtColumnView(ByVal ColNumber As Int16, _
    '                         ByVal DGrid As DataGridView, _
    '                         ByVal MapName As String, _
    '                         ByVal HeadText As String, _
    '                         ByVal ColWidth As Integer, _
    '                         ByVal ColAllignment As System.Windows.Forms.DataGridViewContentAlignment, _
    '                         ByVal ColVisible As Boolean)
    '    DGrid.Columns(ColNumber).Width = ColWidth
    '    DGrid.Columns(ColNumber).ValueType = System.Type.GetType("STRING")
    '    DGrid.Columns(ColNumber).Visible = ColVisible
    '    DGrid.Columns(ColNumber).DataPropertyName = MapName
    '    DGrid.Columns(ColNumber).Name = HeadText
    '    DGrid.Columns(ColNumber).DefaultCellStyle.Alignment = ColAllignment
    '    DGrid.Columns(ColNumber).DefaultCellStyle.NullValue = ""
    '    DGrid.Columns(ColNumber).ReadOnly = True
    'End Sub

    'Public Sub numColumnView(ByVal ColNumber As Int16, _
    '                     ByVal DGrid As DataGridView, _
    '                     ByVal MapName As String, _
    '                     ByVal HeadText As String, _
    '                     ByVal ColWidth As Integer, _
    '                     ByVal ColAllignment As System.Windows.Forms.DataGridViewContentAlignment, _
    '                     ByVal ColFormat As String, _
    '                     ByVal ColVisible As Boolean)
    '    DGrid.Columns(ColNumber).Width = ColWidth
    '    DGrid.Columns(ColNumber).ValueType = System.Type.GetType("STRING")
    '    DGrid.Columns(ColNumber).Visible = ColVisible
    '    DGrid.Columns(ColNumber).DataPropertyName = MapName
    '    DGrid.Columns(ColNumber).Name = HeadText
    '    DGrid.Columns(ColNumber).DefaultCellStyle.Alignment = ColAllignment
    '    DGrid.Columns(ColNumber).DefaultCellStyle.NullValue = "0"
    '    DGrid.Columns(ColNumber).ReadOnly = True
    'End Sub


    'Public Sub logColumnView(ByVal ColNumber As Int16, _
    '                         ByVal DGrid As DataGridView, _
    '                         ByVal MapName As String, _
    '                         ByVal HeadText As String, _
    '                         ByVal ColWidth As Integer, _
    '                         ByVal ColVisible As Boolean)
    '    DGrid.Columns(ColNumber).Width = ColWidth
    '    DGrid.Columns(ColNumber).ValueType = System.Type.GetType("CheckBox")
    '    DGrid.Columns(ColNumber).Visible = ColVisible
    '    DGrid.Columns(ColNumber).DataPropertyName = MapName
    '    DGrid.Columns(ColNumber).Name = HeadText
    '    DGrid.Columns(ColNumber).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    '    DGrid.Columns(ColNumber).DefaultCellStyle.NullValue = False
    '    DGrid.Columns(ColNumber).ReadOnly = True
    'End Sub











    '**************************************************
    '* pobieranie danych z pola rekordu DataGridView - biezacy wiersz + nr kolumny
    '**************************************************

    Public Function GetGuidField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal ColNo As Integer) As String
        If dGridView.CurrentCell Is Nothing Then
            Return ""       'Guid.Empty
        Else
            If IsDBNull(dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value) Then
                Return ""       'Guid.Empty
            Else
                Return dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value
            End If
        End If
    End Function

    Public Function GetBinField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal ColNo As Integer) As Byte()
        If dGridView.CurrentCell Is Nothing Then
            Return Nothing
        Else
            If IsDBNull(dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value) Then
                Return Nothing
            Else
                Return dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value
            End If
        End If
    End Function

    Public Function GetTxtField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal ColNo As Integer) As String
        If dGridView.CurrentCell Is Nothing Then
            Return ""
        Else
            If IsDBNull(dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value) Then
                Return ""
            Else
                'Return Trim(dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value)
                Return dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value
            End If
        End If
    End Function

    Public Function GetByteField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal ColNo As Integer) As Byte()
        If dGridView.CurrentCell Is Nothing Then
            Return Nothing
        Else
            If IsDBNull(dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value) Then
                Return Nothing
            Else
                'Return Trim(dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value)
                Return dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value
            End If
        End If
    End Function

    Public Function GetDateField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal ColNo As Integer) As String
        If dGridView.CurrentCell Is Nothing Then
            Return SysData()
        Else
            If IsDBNull(dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value) Then
                Return SysData()
            Else
                Return Trim(dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value)
            End If
        End If
    End Function

    Public Function GetLogField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal ColNo As Integer) As Boolean
        If dGridView.CurrentCell Is Nothing Then
            Return False
        Else
            If IsDBNull(dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value) Then
                Return False
            Else
                Return dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value
            End If
        End If
    End Function

    Public Function GetNumField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal ColNo As Integer) As Double        'Return (IIf(IsDBNull(dGridView.Item(dGridView.CurrentCell.RowIndex, ColNo)), 0, dGridView.Item(dGridView.CurrentCell.RowIndex, ColNo)))
        If dGridView.CurrentCell Is Nothing Then
            Return 0
        Else
            If IsDBNull(dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value) Then
                Return 0
            Else
                Return dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value
            End If
        End If
    End Function

    Public Function GetDblField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal ColNo As Integer) As Double        'Return (IIf(IsDBNull(dGridView.Item(dGridView.CurrentCell.RowIndex, ColNo)), 0, dGridView.Item(dGridView.CurrentCell.RowIndex, ColNo)))
        If dGridView.CurrentCell Is Nothing Then
            Return 0
        Else
            If IsDBNull(dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value) Then
                Return 0
            Else
                Return dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value
            End If
        End If
    End Function

    Public Function GetIntField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal ColNo As Integer) As Integer
        If dGridView.CurrentCell Is Nothing Then
            Return 0
        Else
            If IsDBNull(dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value) Then
                Return 0
            Else
                Return dGridView.Rows(dGridView.CurrentCell.RowIndex).Cells(ColNo).Value
            End If
        End If
    End Function






    '**************************************************
    '* pobieranie danych z pola rekordu DataGridView - nr wiersza + nr kolumny
    '**************************************************

    Public Function GetGuidField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal RowNo As Integer, ByVal ColNo As Integer) As Guid
        If IsDBNull(dGridView.Rows(RowNo).Cells(ColNo).Value) Then
            Return Guid.Empty
        Else
            Return dGridView.Rows(RowNo).Cells(ColNo).Value
        End If
    End Function

    Public Function GetBinField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal RowNo As Integer, ByVal ColNo As Integer) As Byte()
        If IsDBNull(dGridView.Rows(RowNo).Cells(ColNo)) Then
            Return Nothing
        Else
            Return dGridView.Rows(RowNo).Cells(ColNo).Value
        End If
    End Function

    Public Function GetTxtField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal RowNo As Integer, ByVal ColNo As Integer) As String
        If IsDBNull(dGridView.Rows(RowNo).Cells(ColNo).Value) Then
            Return ""
        Else
            Return Trim(dGridView.Rows(RowNo).Cells(ColNo).Value)
        End If
    End Function

    Public Function GetLogField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal RowNo As Integer, ByVal ColNo As Integer) As Boolean
        If IsDBNull(dGridView.Rows(RowNo).Cells(ColNo).Value) Then
            Return False
        Else
            Return dGridView.Rows(RowNo).Cells(ColNo).Value
        End If
    End Function

    Public Function GetNumField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal RowNo As Integer, ByVal ColNo As Integer) As Double
        If IsDBNull(dGridView.Rows(RowNo).Cells(ColNo).Value) Then
            Return 0
        Else
            Return dGridView.Rows(RowNo).Cells(ColNo).Value
        End If
    End Function

    Public Function GetDblField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal RowNo As Integer, ByVal ColNo As Integer) As Double
        If IsDBNull(dGridView.Rows(RowNo).Cells(ColNo).Value) Then
            Return 0
        Else
            Return dGridView.Rows(RowNo).Cells(ColNo).Value
        End If
    End Function

    Public Function GetIntField(ByVal dGridView As System.Windows.Forms.DataGridView, ByVal RowNo As Integer, ByVal ColNo As Integer) As Integer
        If IsDBNull(dGridView.Rows(RowNo).Cells(ColNo).Value) Then
            Return 0
        Else
            Return dGridView.Rows(RowNo).Cells(ColNo).Value
        End If
    End Function



    '**************************************************
    '* pobieranie danych ze zmiennej tekstowej
    '**************************************************


    Public Function GetLogFromText(ByVal dText As String) As Boolean
        If Len(Trim(dText)) = 0 Then
            Return False
        Else
            Select Case dText
                Case "T", "t"
                    Return True
                Case Else
                    Return False
            End Select
        End If
    End Function

    Public Function GetNumFromText(ByVal dText As String) As Double
        If Len(Trim(dText)) = 0 Then
            Return 0
        Else
            If Not IsNumeric(dText) Then
                Return 0
            Else
                Return CDbl(dText)
            End If
        End If
    End Function

    Public Function GetDblFromText(ByVal dText As String) As Double
        If Len(Trim(dText)) = 0 Then
            Return 0
        Else
            If Not IsNumeric(dText) Then
                Return 0
            Else
                Return CDbl(dText)
            End If
        End If
    End Function

    Public Function GetIntFromText(ByVal dText As String) As Int32
        If Len(Trim(dText)) = 0 Then
            Return 0
        Else
            If Not IsNumeric(dText) Then
                Return 0
            Else
                Return CInt(dText)
            End If
        End If
    End Function









    '**************************************************
    '* pobieranie danych z pola rekordu datagrid - biezacy wiersz + nr kolumny
    '**************************************************

    Public Function GetGuidField(ByVal dGrid As System.Windows.Forms.DataGrid, ByVal FieldNo As Integer) As String
        If IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)) Then
            Return ""
        Else
            Return dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)
        End If
    End Function

    Public Function GetBinField(ByVal dGrid As System.Windows.Forms.DataGrid, ByVal fieldNo As Integer) As Byte()
        If IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, fieldNo)) Then
            Return Nothing
        Else
            Return dGrid.Item(dGrid.CurrentCell.RowNumber, fieldNo)
        End If
    End Function

    Public Function GetTxtField(ByVal dGrid As System.Windows.Forms.DataGrid, ByVal FieldNo As Integer) As String
        If IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)) Then
            Return ""
        Else
            Return dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)
        End If
    End Function

    Public Function GetByteField(ByVal dGrid As System.Windows.Forms.DataGrid, ByVal FieldNo As Integer) As Byte()
        If IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)) Then
            Return Nothing
        Else
            Return dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)
        End If
    End Function

    Public Function GetDateField(ByVal dGrid As System.Windows.Forms.DataGrid, ByVal FieldNo As Integer) As String
        If IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)) Then
            Return SysData()
        Else
            Return dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)
        End If
    End Function

    Public Function GetLogField(ByVal dGrid As System.Windows.Forms.DataGrid, ByVal FieldNo As Integer) As Boolean
        'Return (IIf(IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)), False, dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)))
        If IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)) Then
            Return False
        Else
            Return dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)
        End If
    End Function

    Public Function GetNumField(ByVal dGrid As System.Windows.Forms.DataGrid, ByVal fieldNo As Integer) As Double        'Return (IIf(IsDBNull(dGridView.Item(dGridView.CurrentCell.RowIndex, ColNo)), 0, dGridView.Item(dGridView.CurrentCell.RowIndex, ColNo)))
        If IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, fieldNo)) Then
            Return 0
        Else
            Return dGrid.Item(dGrid.CurrentCell.RowNumber, fieldNo)
        End If
    End Function

    Public Function GetDblField(ByVal dGrid As System.Windows.Forms.DataGrid, ByVal FieldNo As Integer) As Double        'Return (IIf(IsDBNull(dGridView.Item(dGridView.CurrentCell.RowIndex, ColNo)), 0, dGridView.Item(dGridView.CurrentCell.RowIndex, ColNo)))
        If IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)) Then
            Return 0
        Else
            Return dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)
        End If
    End Function

    Public Function GetIntField(ByVal dGrid As System.Windows.Forms.DataGrid, ByVal FieldNo As Integer) As Integer
        'Return (IIf(IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)), 0, dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)))
        If IsDBNull(dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)) Then
            Return 0
        Else
            Return dGrid.Item(dGrid.CurrentCell.RowNumber, FieldNo)
        End If
    End Function




End Module
