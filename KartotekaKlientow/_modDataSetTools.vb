Module _modDataSetTools


    '**************************************************
    '* pobieranie danych z pola rekordu DataSet
    '**************************************************

    Public Function GetGuidDRField(ByVal Row As DataRow, ByVal FieldN As String) As Guid
        'Return (IIf(IsDBNull(Row.Item(FieldN)), Guid.Empty, Row.Item(FieldN)))
        If IsDBNull(Row.Item(FieldN)) Then
            Return Guid.Empty
        Else
            Return Row.Item(FieldN)
        End If
    End Function

    Public Function GetTxtDRField(ByVal Row As DataRow, ByVal FieldN As String) As String
        'Return (Trim(IIf(IsDBNull(Row.Item(FieldN)), "", Row.Item(FieldN))))
        If IsDBNull(Row.Item(FieldN)) Then
            Return ""
        Else
            Return Row.Item(FieldN)
        End If
    End Function

    Public Function GetDateDRField(ByVal Row As DataRow, ByVal FieldN As String) As String
        'Return (Trim(IIf(IsDBNull(Row.Item(FieldN)), SysData(), Row.Item(FieldN))))
        If IsDBNull(Row.Item(FieldN)) Then
            Return SysData()
        Else
            Return Row.Item(FieldN)
        End If
    End Function

    Public Function GetLogDRField(ByVal Row As DataRow, ByVal FieldN As String) As Boolean
        'Return (IIf(IsDBNull(Row.Item(FieldN)), False, Row.Item(FieldN)))
        If IsDBNull(Row.Item(FieldN)) Then
            Return False
        Else
            Return Row.Item(FieldN)
        End If
    End Function

    Public Function GetIntDRField(ByVal Row As DataRow, ByVal FieldN As String) As Integer
        'Return (IIf(IsDBNull(Row.Item(FieldN)), 0, Row.Item(FieldN)))
        If IsDBNull(Row.Item(FieldN)) Then
            Return 0
        Else
            Return Row.Item(FieldN)
        End If
    End Function

    Public Function GetNumDRField(ByVal Row As DataRow, ByVal FieldN As String) As Double
        'Return (IIf(IsDBNull(Row.Item(FieldN)), 0, Row.Item(FieldN)))
        If IsDBNull(Row.Item(FieldN)) Then
            Return 0
        Else
            Return Row.Item(FieldN)
        End If
    End Function

    Public Function GetDblDRField(ByVal Row As DataRow, ByVal FieldN As String) As Double
        'Return (IIf(IsDBNull(Row.Item(FieldN)), 0, Row.Item(FieldN)))
        If IsDBNull(Row.Item(FieldN)) Then
            Return 0
        Else
            Return Row.Item(FieldN)
        End If
    End Function





End Module
