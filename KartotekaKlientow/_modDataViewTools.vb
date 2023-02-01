Module _modDataViewTools


    '**************************************************
    '* pobieranie danych z pola rekordu DataView
    '**************************************************

    Public Function GetGuidDRVField(ByVal Row As DataRowView, ByVal FieldN As String) As Guid
        'Return (IIf(IsDBNull(Row.Item(FieldN)), Guid.Empty, Row.Item(FieldN)))
        If IsDBNull(Row.Item(FieldN)) Then
            Return Guid.Empty
        Else
            Return Row.Item(FieldN)
        End If
    End Function

    Public Function GetTxtDRVField(ByVal Row As DataRowView, ByVal FieldN As String) As String
        'Return (Trim(IIf(IsDBNull(Row.Item(FieldN)), "", Row.Item(FieldN))))
        If IsDBNull(Row.Item(FieldN)) Then
            Return ""
        Else
            Return Row.Item(FieldN)
        End If
    End Function

    Public Function GetTxtDRVFieldTrim(ByVal Row As DataRowView, ByVal FieldN As String) As String
        'Return (Trim(IIf(IsDBNull(Row.Item(FieldN)), "", Row.Item(FieldN))))
        If IsDBNull(Row.Item(FieldN)) Then
            Return ""
        Else
            Return Trim(Row.Item(FieldN))
        End If
    End Function

    Public Function GetDateDRVField(ByVal Row As DataRowView, ByVal FieldN As String) As String
        'Return (Trim(IIf(IsDBNull(Row.Item(FieldN)), SysData(), Row.Item(FieldN))))
        If IsDBNull(Row.Item(FieldN)) Then
            Return SysData()
        Else
            Return Row.Item(FieldN)
        End If
    End Function

    Public Function GetLogDRVField(ByVal Row As DataRowView, ByVal FieldN As String) As Boolean
        'Return (IIf(IsDBNull(Row.Item(FieldN)), False, Row.Item(FieldN)))
        If IsDBNull(Row.Item(FieldN)) Then
            Return False
        Else
            Return Row.Item(FieldN)
        End If
    End Function

    Public Function GetIntDRVField(ByVal Row As DataRowView, ByVal FieldN As String) As Integer
        'Return (IIf(IsDBNull(Row.Item(FieldN)), 0, Row.Item(FieldN)))
        If IsDBNull(Row.Item(FieldN)) Then
            Return 0
        Else
            Return Row.Item(FieldN)
        End If
    End Function

    Public Function GetNumDRVField(ByVal Row As DataRowView, ByVal FieldN As String) As Double
        'Return (IIf(IsDBNull(Row.Item(FieldN)), 0, Row.Item(FieldN)))
        If IsDBNull(Row.Item(FieldN)) Then
            Return 0
        Else
            Return Row.Item(FieldN)
        End If
    End Function

    Public Function GetDblDRVField(ByVal Row As DataRowView, ByVal FieldN As String) As Double
        'Return (IIf(IsDBNull(Row.Item(FieldN)), 0, Row.Item(FieldN)))
        If IsDBNull(Row.Item(FieldN)) Then
            Return 0
        Else
            Return Row.Item(FieldN)
        End If
    End Function







    Public Function GetTxt(ByVal Field As String) As String
        'Return (Trim(IIf(IsDBNull(Field), "", Field)))
        If IsDBNull(Field) Then
            Return 0
        Else
            Return Field
        End If
    End Function

    Public Function GetTxtTrim(ByVal Field As String) As String
        'Return (Trim(IIf(IsDBNull(Field), "", Field)))
        If IsDBNull(Field) Then
            Return 0
        Else
            Return Trim(Field)
        End If
    End Function

End Module
