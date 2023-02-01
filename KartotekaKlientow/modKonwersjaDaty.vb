Module modKonwersjaDaty

    Public Function TestDate(ByVal strdate As String) As Boolean
        Dim RetBool As Boolean = True
        'zamien date w formacie STRING yyyy-MM-dd na DateTime w formacie regionalnym
        'nast zamien na string w formacie regionalnym
        If System.Globalization.CultureInfo.CurrentUICulture.DateTimeFormat.ShortDatePattern = "yyyy-MM-dd" Then
            RetBool = True
        Else
            Dim RetData As String = ""
            Try
                RetData = DateTime.ParseExact(strdate, "yyyy-MM-dd", System.Globalization.CultureInfo.CurrentUICulture.DateTimeFormat).ToShortDateString()
            Catch ex As Exception
                MessageBox.Show("Nieprawidłowy zapis w polu DATA." & vbCrLf & strdate, _
                    "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Exclamation, _
                    MessageBoxDefaultButton.Button1)
                RetBool = False
            End Try
        End If
        Return RetBool
    End Function


    'format datu w programie INTERNAL : yyyy-MM-dd
    'format REGIONALNY - zalezny od ustawień systemu
    Public Function ConvertInternalDateToRegional(ByVal strdate As String) As String
        'zamien date w formacie STRING yyyy-MM-dd na DateTime w formacie regionalnym
        'nast zamien na string w formacie regionalnym
        If System.Globalization.CultureInfo.CurrentUICulture.DateTimeFormat.ShortDatePattern = "yyyy-MM-dd" Then
            Return strdate
        Else
            Dim RetData As String = ""
            Try
                RetData = DateTime.ParseExact(strdate, "yyyy-MM-dd", System.Globalization.CultureInfo.CurrentUICulture.DateTimeFormat).ToShortDateString()
            Catch ex As Exception
                RetData = SysData()
                MessageBox.Show("Nieprawidłowy zapis w polu DATA.", _
                    "Uwaga :", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Exclamation, _
                    MessageBoxDefaultButton.Button1)
            End Try
            Return RetData
        End If
    End Function

    Public Function ConvertRegionalDateToInternal(ByVal strdate As String) As String
        'zamien date w formacie STRING REGIONALNY na DateTime w formacie regionalnym
        'nast zamien na string w formacie yyyy-MM-dd

        'MsgBox("ConvertRegionalDateToInternal  " & strdate & vbCrLf & _
        '       System.Globalization.CultureInfo.CurrentUICulture.DateTimeFormat.ShortDatePattern & vbCrLf & _
        '       DateTime.Parse(strdate).ToShortDateString())

        If System.Globalization.CultureInfo.CurrentUICulture.DateTimeFormat.ShortDatePattern = "yyyy-MM-dd" Then
            Return strdate
        Else
            Dim DTFormat As String = System.Globalization.CultureInfo.CurrentUICulture.DateTimeFormat.ShortDatePattern
            'Dim Ro As Integer = DateTime.ParseExact(strdate, DTFormat, Nothing).Year
            'Dim Mi As Integer = DateTime.ParseExact(strdate, DTFormat, Nothing).Month
            'Dim Dz As Integer = DateTime.ParseExact(strdate, DTFormat, Nothing).Day

            'Dim Ro As Integer = DateTime.ParseExact(DateTime.Parse(strdate).ToShortDateString(), DTFormat, Nothing).Year
            'Dim Mi As Integer = DateTime.ParseExact(DateTime.Parse(strdate).ToShortDateString(), DTFormat, Nothing).Month
            'Dim Dz As Integer = DateTime.ParseExact(DateTime.Parse(strdate).ToShortDateString(), DTFormat, Nothing).Day

            Dim Ro As Integer = DateTime.Parse(strdate).Year
            Dim Mi As Integer = DateTime.Parse(strdate).Month
            Dim Dz As Integer = DateTime.Parse(strdate).Day

            If Ro < 100 Then
                'format roku=rr
                Ro += 2000
            Else
                'format roku=rrrr
            End If

            Return Right("0000" & Trim(Ro.ToString()), 4) & "-" & _
                   Right("00" & Trim(Mi.ToString()), 2) & "-" & _
                   Right("00" & Trim(Dz.ToString()), 2)
        End If
    End Function


End Module
