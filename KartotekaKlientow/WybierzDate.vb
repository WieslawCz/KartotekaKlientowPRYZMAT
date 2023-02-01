Public Class WybierzDate
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    Friend WithEvents Kalendarz As System.Windows.Forms.MonthCalendar
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbxRok As System.Windows.Forms.ComboBox
    Friend WithEvents cbxMiesiac As System.Windows.Forms.ComboBox
    Friend WithEvents cbxDzien As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(WybierzDate))
        Me.Kalendarz = New System.Windows.Forms.MonthCalendar()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cbxDzien = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbxMiesiac = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbxRok = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Kalendarz
        '
        Me.Kalendarz.BackColor = System.Drawing.SystemColors.Control
        Me.Kalendarz.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Kalendarz.Location = New System.Drawing.Point(8, 104)
        Me.Kalendarz.Name = "Kalendarz"
        Me.Kalendarz.ShowWeekNumbers = True
        Me.Kalendarz.TabIndex = 0
        Me.Kalendarz.TitleBackColor = System.Drawing.Color.Gray
        Me.Kalendarz.TitleForeColor = System.Drawing.Color.Yellow
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPowrot.Image = CType(resources.GetObject("cmdPowrot.Image"), System.Drawing.Image)
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(216, 278)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(80, 22)
        Me.cmdPowrot.TabIndex = 2
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cbxDzien)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.cbxMiesiac)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.cbxRok)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(291, 88)
        Me.Panel1.TabIndex = 13
        '
        'cbxDzien
        '
        Me.cbxDzien.Location = New System.Drawing.Point(56, 56)
        Me.cbxDzien.Name = "cbxDzien"
        Me.cbxDzien.Size = New System.Drawing.Size(96, 21)
        Me.cbxDzien.TabIndex = 17
        Me.cbxDzien.Text = "01"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Navy
        Me.Label3.Location = New System.Drawing.Point(8, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 16)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Dzieñ . . . . . . . . . . ."
        '
        'cbxMiesiac
        '
        Me.cbxMiesiac.Location = New System.Drawing.Point(56, 32)
        Me.cbxMiesiac.Name = "cbxMiesiac"
        Me.cbxMiesiac.Size = New System.Drawing.Size(96, 21)
        Me.cbxMiesiac.TabIndex = 16
        Me.cbxMiesiac.Text = "PaŸdziernik"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Navy
        Me.Label2.Location = New System.Drawing.Point(8, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Miesi¹c . . . . . . . . . . ."
        '
        'cbxRok
        '
        Me.cbxRok.Location = New System.Drawing.Point(56, 8)
        Me.cbxRok.Name = "cbxRok"
        Me.cbxRok.Size = New System.Drawing.Size(96, 21)
        Me.cbxRok.TabIndex = 15
        Me.cbxRok.Text = "2005"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Rok . . . . . . . . . . ."
        '
        'WybierzDate
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(308, 310)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.Kalendarz)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "WybierzDate"
        Me.ShowInTaskbar = False
        Me.Text = "Wybierz datê"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim RobData As String
    Dim StartFormy As Boolean = True

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        oData = ConvertRegionalDateToInternal(Kalendarz.SelectionRange.Start.ToShortDateString())

        Dim DayOfYear As Integer = Kalendarz.SelectionRange.Start.DayOfYear
        Dim DayOfWeek As Integer = Kalendarz.SelectionRange.Start.DayOfWeek
        Dim FirstDayOfWeek As Integer

        DayOfWeek = IIf(DayOfWeek = 0, 7, DayOfWeek)

        If (DayOfYear < DayOfWeek) Then
            FirstDayOfWeek = (DayOfWeek - DayOfYear + 1)
        Else
            FirstDayOfWeek = (DayOfYear - DayOfWeek) Mod 7
        End If
        FirstDayOfWeek = IIf(FirstDayOfWeek = 0, 7, FirstDayOfWeek)

        oWeek = Int((DayOfYear - DayOfWeek + (FirstDayOfWeek - 1)) / 7) + 1
        'zmiana 05.09.2015 - tydzien przelomowy roku = zawsze ostatni tydzien starego roku
        If oWeek = 1 Then
            oWeek = 53
            'oData = TerminZaplaty(oData, -DayOfWeek)
            oData = TerminZaplaty(oData, -DayOfYear)
        End If
        Me.Close()
    End Sub


    Private Sub Kalendarz_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles Kalendarz.DateSelected
        ' Show the start and end dates in the text box.
        'MessageBox.Show("Date Selected: Start = " + _
        '        e.Start.ToShortDateString() + " : End = " + e.End.ToShortDateString())
        oData = ConvertRegionalDateToInternal(e.Start.ToShortDateString())

        Dim DayOfYear As Integer = e.Start.DayOfYear
        Dim DayOfWeek As Integer = e.Start.DayOfWeek
        Dim FirstDayOfWeek As Integer

        DayOfWeek = IIf(DayOfWeek = 0, 7, DayOfWeek)

        If (DayOfYear < DayOfWeek) Then
            FirstDayOfWeek = (DayOfWeek - DayOfYear + 1)
        Else
            FirstDayOfWeek = (DayOfYear - DayOfWeek) Mod 7
        End If
        FirstDayOfWeek = IIf(FirstDayOfWeek = 0, 7, FirstDayOfWeek)

        oWeek = Int((DayOfYear - DayOfWeek + (FirstDayOfWeek - 1)) / 7) + 1
        ''zmiana 05.09.2015 - tydzien przelomowy roku = zawsze ostatni tydzien starego roku
        'If oWeek = 1 Then
        '    oWeek = 53
        '    'oData = TerminZaplaty(oData, -DayOfWeek)
        '    oData = TerminZaplaty(oData, -DayOfYear)
        'End If
        ''zmiana 25.04.2016 - MA BYÆ DOBRZE...


        Me.Close()
    End Sub


    Private Sub WybierzDate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.TopMost = True
        Me.Icon = My.Resources.MrJones
        '-------------------------------
        'Lista lat
        Dim Ir As Integer
        For Ir = 1900 To 2100
            cbxRok.Items.Add(Trim(Str(Ir)))
        Next

        'Lista miesiêcy
        cbxMiesiac.Items.Add("Styczeñ")
        cbxMiesiac.Items.Add("Luty")
        cbxMiesiac.Items.Add("Marzec")
        cbxMiesiac.Items.Add("Kwiecieñ")
        cbxMiesiac.Items.Add("Maj")
        cbxMiesiac.Items.Add("Czerwiec")
        cbxMiesiac.Items.Add("Lipiec")
        cbxMiesiac.Items.Add("Sierpieñ")
        cbxMiesiac.Items.Add("Wrzesieñ")
        cbxMiesiac.Items.Add("PaŸdziernik")
        cbxMiesiac.Items.Add("Listopad")
        cbxMiesiac.Items.Add("Grudzieñ")

        'Lista dni
        For Ir = 1 To 31
            cbxDzien.Items.Add(Trim(Str(Ir)))
        Next

        If Len(oData) > 0 Then
            If TestDate(oData) Then
                'OK
            Else
                oData = SysData()
            End If
        Else
            oData = SysData()


            ''wylicz date na podstawie tygodnia w roku
            'Dim DayOfYear As Integer = 7 * oWeek    '- FirstDayOfWeekOfTheYear(Year(Now))

            'If DayOfYear = 0 Then
            '    oData = Year(Now).ToString & "-01-01"
            'ElseIf DayOfYear <= 31 Then
            '    oData = Year(Now).ToString & "-01-" & Mid(Trim(Str(100 + DayOfYear)), 2, 2)
            'ElseIf DayOfYear <= 31 + DaysOfFebruary() Then
            '    oData = Year(Now).ToString & "-02-" & Mid(Trim(Str(100 + DayOfYear - 31)), 2, 2)
            'ElseIf DayOfYear <= 31 + DaysOfFebruary() + 31 Then
            '    oData = Year(Now).ToString & "-03-" & Mid(Trim(Str(100 + DayOfYear - 31 - DaysOfFebruary())), 2, 2)
            'ElseIf DayOfYear <= 31 + DaysOfFebruary() + 31 + 30 Then
            '    oData = Year(Now).ToString & "-04-" & Mid(Trim(Str(100 + DayOfYear - 31 - DaysOfFebruary() - 31)), 2, 2)
            'ElseIf DayOfYear <= 31 + DaysOfFebruary() + 31 + 30 + 31 Then
            '    oData = Year(Now).ToString & "-05-" & Mid(Trim(Str(100 + DayOfYear - 31 - DaysOfFebruary() - 31 - 30)), 2, 2)
            'ElseIf DayOfYear <= 31 + DaysOfFebruary() + 31 + 30 + 31 + 30 Then
            '    oData = Year(Now).ToString & "-06-" & Mid(Trim(Str(100 + DayOfYear - 31 - DaysOfFebruary() - 31 - 30 - 31)), 2, 2)
            'ElseIf DayOfYear <= 31 + DaysOfFebruary() + 31 + 30 + 31 + 30 + 31 Then
            '    oData = Year(Now).ToString & "-07-" & Mid(Trim(Str(100 + DayOfYear - 31 - DaysOfFebruary() - 31 - 30 - 31 - 30)), 2, 2)
            'ElseIf DayOfYear <= 31 + DaysOfFebruary() + 31 + 30 + 31 + 30 + 31 + 31 Then
            '    oData = Year(Now).ToString & "-08-" & Mid(Trim(Str(100 + DayOfYear - 31 - DaysOfFebruary() - 31 - 30 - 31 - 30 - 31)), 2, 2)
            'ElseIf DayOfYear <= 31 + DaysOfFebruary() + 31 + 30 + 31 + 30 + 31 + 31 + 30 Then
            '    oData = Year(Now).ToString & "-09-" & Mid(Trim(Str(100 + DayOfYear - 31 - DaysOfFebruary() - 31 - 30 - 31 - 30 - 31 - 31)), 2, 2)
            'ElseIf DayOfYear <= 31 + DaysOfFebruary() + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 Then
            '    oData = Year(Now).ToString & "-10-" & Mid(Trim(Str(100 + DayOfYear - 31 - DaysOfFebruary() - 31 - 30 - 31 - 30 - 31 - 31 - 30)), 2, 2)
            'ElseIf DayOfYear <= 31 + DaysOfFebruary() + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + 30 Then
            '    oData = Year(Now).ToString & "-11-" & Mid(Trim(Str(100 + DayOfYear - 31 - DaysOfFebruary() - 31 - 30 - 31 - 30 - 31 - 31 - 30 - 31)), 2, 2)
            'ElseIf DayOfYear <= 31 + DaysOfFebruary() + 31 + 30 + 31 + 30 + 31 + 31 + 30 + 31 + 30 + 31 Then
            '    oData = Year(Now).ToString & "-12-" & Mid(Trim(Str(100 + DayOfYear - 31 - DaysOfFebruary() - 31 - 30 - 31 - 30 - 31 - 31 - 30 - 31 - 30)), 2, 2)
            'Else
            '    oData = (Year(Now) + 1).ToString & "-01-01"
            'End If
        End If

        Kalendarz.SelectionStart = ConvertInternalDateToRegional(oData)
        Kalendarz.SelectionEnd = ConvertInternalDateToRegional(oData)

        cbxRok.Text = Mid(oData, 1, 4)

        Select Case Mid(oData, 6, 2)
            Case "01"
                cbxMiesiac.Text = "Styczeñ"
            Case "02"
                cbxMiesiac.Text = "Luty"
            Case "03"
                cbxMiesiac.Text = "Marzec"
            Case "04"
                cbxMiesiac.Text = "Kwiecieñ"
            Case "05"
                cbxMiesiac.Text = "Maj"
            Case "06"
                cbxMiesiac.Text = "Czerwiec"
            Case "07"
                cbxMiesiac.Text = "Lipiec"
            Case "08"
                cbxMiesiac.Text = "Sierpieñ"
            Case "09"
                cbxMiesiac.Text = "Wrzesieñ"
            Case "10"
                cbxMiesiac.Text = "PaŸdziernik"
            Case "11"
                cbxMiesiac.Text = "Listopad"
            Case "12"
                cbxMiesiac.Text = "Grudzieñ"
        End Select

        cbxDzien.Text = Mid(oData, 9, 2)

        StartFormy = False
    End Sub


    Private Sub WybierzDate_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub





    Private Sub cbxRok_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxRok.TextChanged
        If StartFormy Then
        Else
            RobData = PobierzWpisanaDate()
            Kalendarz.SelectionStart = ConvertInternalDateToRegional(RobData)
            Kalendarz.SelectionEnd = ConvertInternalDateToRegional(RobData)
        End If
    End Sub

    Private Sub cbxMiesiac_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxMiesiac.TextChanged
        If StartFormy Then
        Else
            RobData = PobierzWpisanaDate()
            Kalendarz.SelectionStart = ConvertInternalDateToRegional(RobData)
            Kalendarz.SelectionEnd = ConvertInternalDateToRegional(RobData)
        End If
    End Sub

    Private Sub cbxDzien_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxDzien.TextChanged
        If StartFormy Then
        Else
            RobData = PobierzWpisanaDate()
            Kalendarz.SelectionStart = ConvertInternalDateToRegional(RobData)
            Kalendarz.SelectionEnd = ConvertInternalDateToRegional(RobData)
        End If
    End Sub


    Private Function PobierzWpisanaDate() As String
        Dim dRok As String = ""
        Dim dMiesiac As String = ""
        Dim dDzien As String = ""

        Dim iDzien As Integer = 0
        Dim iRok As Integer = 0

        dRok = cbxRok.Text
        dDzien = cbxDzien.Text

        iDzien = CInt(dDzien)
        iRok = CInt(dRok)

        Select Case cbxMiesiac.Text
            Case "Styczeñ"
                dMiesiac = "01"
            Case "Luty"
                dMiesiac = "02"
                If (Int(iRok / 4) = (iRok / 4)) And (Int(iRok / 100) <> (iRok / 100)) Then
                    'rok przestepny
                    If iDzien > 29 Then
                        dDzien = "29"
                    End If
                Else
                    'rok normalny
                    If iDzien > 28 Then
                        dDzien = "28"
                    End If
                End If
            Case "Marzec"
                dMiesiac = "03"
            Case "Kwiecieñ"
                dMiesiac = "04"
                If iDzien > 30 Then
                    dDzien = "30"
                End If
            Case "Maj"
                dMiesiac = "05"
            Case "Czerwiec"
                dMiesiac = "06"
                If iDzien > 30 Then
                    dDzien = "30"
                End If
            Case "Lipiec"
                dMiesiac = "07"
            Case "Sierpieñ"
                dMiesiac = "08"
            Case "Wrzesieñ"
                dMiesiac = "09"
                If iDzien > 30 Then
                    dDzien = "30"
                End If
            Case "PaŸdziernik"
                dMiesiac = "10"
            Case "Listopad"
                dMiesiac = "11"
                If iDzien > 30 Then
                    dDzien = "30"
                End If
            Case "Grudzieñ"
                dMiesiac = "12"
        End Select
        cbxDzien.Text = dDzien

        Return (dRok & "-" & dMiesiac & "-" & dDzien)
    End Function





    Private Sub Kalendarz_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles Kalendarz.DateChanged
        Dim Ro As Integer = e.Start.Year
        Dim Mi As Integer = e.Start.Month
        Dim Dz As Integer = e.Start.Day

        'SysData = Right("0000" & Trim(Now.Year.ToString()), 4) & "-" & _
        '  Right("00" & Trim(Now.Month.ToString()), 2) & "-" & _
        '  Right("00" & Trim(Now.Day.ToString()), 2)

        If Ro < 100 Then
            'format roku=rr
            cbxRok.Text = Microsoft.VisualBasic.Right("0000" & Trim((2000 + Ro).ToString()), 4)
        Else
            'format roku=rrrr
            cbxRok.Text = Microsoft.VisualBasic.Right("0000" & Trim(Ro.ToString()), 4)
        End If

        Select Case Microsoft.VisualBasic.Right("00" & Trim(Mi.ToString()), 2)
            Case "01"
                cbxMiesiac.Text = "Styczeñ"
            Case "02"
                cbxMiesiac.Text = "Luty"
            Case "03"
                cbxMiesiac.Text = "Marzec"
            Case "04"
                cbxMiesiac.Text = "Kwiecieñ"
            Case "05"
                cbxMiesiac.Text = "Maj"
            Case "06"
                cbxMiesiac.Text = "Czerwiec"
            Case "07"
                cbxMiesiac.Text = "Lipiec"
            Case "08"
                cbxMiesiac.Text = "Sierpieñ"
            Case "09"
                cbxMiesiac.Text = "Wrzesieñ"
            Case "10"
                cbxMiesiac.Text = "PaŸdziernik"
            Case "11"
                cbxMiesiac.Text = "Listopad"
            Case "12"
                cbxMiesiac.Text = "Grudzieñ"
        End Select

        cbxDzien.Text = Microsoft.VisualBasic.Right("00" & Trim(Dz.ToString()), 2)
    End Sub

    '*****************************************
    '** edycja pol
    '*****************************************

    Private Sub cbxRok_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxRok.GotFocus
        StartEdycjiCbx(cbxRok)
    End Sub
    Private Sub cbxRok_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxRok.LostFocus
        KoniecEdycjiCbx(cbxRok)
    End Sub

    Private Sub cbxMiesiac_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxMiesiac.GotFocus
        StartEdycjiCbx(cbxMiesiac)
    End Sub
    Private Sub cbxMiesiac_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxMiesiac.LostFocus
        KoniecEdycjiCbx(cbxMiesiac)
    End Sub

    Private Sub cbxDzien_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxDzien.GotFocus
        StartEdycjiCbx(cbxDzien)
    End Sub
    Private Sub cbxDzien_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxDzien.LostFocus
        KoniecEdycjiCbx(cbxDzien)
    End Sub





    Private Function DaysOfFebruary() As Integer
        Dim CurYear As Integer = Year(Now)
        'przestepny=podzielny przez 4 i niepodzielny przez 100
        If (CurYear = 4 * Int(CurYear / 4)) And (CurYear <> 100 * Int(CurYear / 100)) Then
            Return (29)
        Else
            Return (28)
        End If
    End Function

    Private Function FirstDayOfWeekOfTheYear(ByVal Rok As Integer) As Integer
        '01-01-2004 wypad³ w czwartek = 04
        Dim ileDni As Integer
        Dim DayOfWeek As Integer

        If Rok > 2004 Then
            ileDni = (Rok - 2004) * 365 + Int((Rok - 1 - 2004) / 4) + 1
            DayOfWeek = (ileDni + 4) Mod 7
        Else
            ileDni = (2004 - Rok) * 365 + Int((2004 - Rok) / 4)
            DayOfWeek = (7 + 4 - (ileDni Mod 7)) Mod 7
        End If
        Return (IIf(DayOfWeek = 0, 7, DayOfWeek))
    End Function

End Class
