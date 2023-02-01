Option Strict Off
Option Explicit On

Imports System
Imports System.Reflection

Namespace Softart
    Public Class myDataGridRowHeightSetter

        'spsosb zastosowania
        'DEFINICJA OBIEKTU DOSTÊPNEGO W MODULE
        '   Dim RowHeightSetter As Softart.myDataGridRowHeightSetter
        '   RowHeightSetter = New Softart.myDataGridRowHeightSetter(Me.dagKontakty, 0, 6)
        'PROCEDURY POSZERZANIA I POMNIEJSZANIA
        '   Private Sub mnuMultiRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMultiRow.Click
        '        RowHeightSetter.MultiHightsView()
        '   End Sub
        '   Private Sub mnuSingleRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSingleRow.Click
        '        RowHeightSetter.SingleHightsView()
        '   End Sub
        'gdy dodano lub usunieto zapisy w tablicy
        '   RowHeightSetter.Dispose()
        '   RowHeightSetter = New Softart.myDataGridRowHeightSetter(Me.dagKontakty, 0, 6)



        Private dg As DataGrid
        Private dv As DataView

        Private rowObjects As ArrayList

        Private _SingleHiRow As Integer = 0
        Private _MultiHiRow As Integer = 0
        Private _SingleHi As Integer = 17

        Public Sub New(ByVal dg As DataGrid, _
                       ByVal SingleHiRow As Integer, _
                       ByVal MultiHiRow As Integer)
            Me.dg = dg
            Me.dv = dg.DataSource
            _SingleHiRow = SingleHiRow
            _MultiHiRow = MultiHiRow

            InitHeights()
            'DefSingleHi()  'nie definiujemy, bo trzeba by wrociæ do SingleLone
        End Sub 'New

        Public Sub Dispose()
            rowObjects = Nothing
        End Sub

        Protected Overrides Sub Finalize()
            Dispose()
            MyBase.Finalize()
        End Sub





        Private Sub InitHeights()
            Dim mi As MethodInfo = dg.GetType().GetMethod("get_DataGridRows", BindingFlags.FlattenHierarchy Or BindingFlags.IgnoreCase Or BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Static)
            Dim dgra As System.Array = CType(mi.Invoke(Me.dg, Nothing), System.Array)

            rowObjects = New ArrayList
            Dim dgrr As Object
            For Each dgrr In dgra
                If dgrr.ToString().EndsWith("DataGridRelationshipRow") = True Then
                    rowObjects.Add(dgrr)
                End If
            Next dgrr
        End Sub 'InitHeights 

        Private Sub DefSingleHi()
            Try
                Dim pi As PropertyInfo = rowObjects(_SingleHiRow).GetType().GetProperty("Height")
                _SingleHi = Fix(pi.GetValue(rowObjects(_SingleHiRow), Nothing))
            Catch
                'Throw New ArgumentException("invalid row index")
            End Try
        End Sub

        '================================

        Default Public Property Row(ByVal Xrow As Integer) As Integer
            Get
                Try
                    Dim pi As PropertyInfo = rowObjects(Xrow).GetType().GetProperty("Height")
                    Return Fix(pi.GetValue(rowObjects(Xrow), Nothing))
                Catch
                    Throw New ArgumentException("invalid row index")
                End Try
            End Get
            Set(ByVal Value As Integer)
                Try
                    Dim pi As PropertyInfo = rowObjects(Xrow).GetType().GetProperty("Height")
                    pi.SetValue(rowObjects(Xrow), Value, Nothing)
                Catch
                    Throw New ArgumentException("invalid row index")
                End Try
            End Set
        End Property

        Public Sub MultiHightsView()
            If dv.Count > 0 Then
                Dim i As Integer
                Dim ili As Integer
                Dim val As Integer
                Dim txt As String
                Dim CrLfPos As Integer
                For i = 0 To dv.Count - 1
                    txt = IIf(IsDBNull(dg.Item(i, _MultiHiRow)), "", dg.Item(i, _MultiHiRow))
                    ili = 0
                    'sprawdz ile linii ma tekst
                    If Len(txt) > 0 Then
                        ili = 1
                        CrLfPos = InStr(txt, vbCrLf)
                        Do While CrLfPos > 0
                            ili += 1
                            txt = Mid(txt, CrLfPos + 2)
                            CrLfPos = InStr(txt, vbCrLf)
                        Loop
                    End If

                    If ili > 1 Then
                        val = (ili) * _SingleHi
                    Else
                        val = _SingleHi
                    End If

                    Try
                        Dim pi As PropertyInfo = rowObjects(i).GetType().GetProperty("Height")
                        pi.SetValue(rowObjects(i), val, Nothing)
                    Catch
                        'Throw New ArgumentException("invalid row index")
                    End Try
                Next
            End If
        End Sub

        Public Sub SingleHightsView()
            If dv.Count > 0 Then
                Dim i As Integer
                For i = 0 To dv.Count - 1
                    Try
                        Dim pi As PropertyInfo = rowObjects(i).GetType().GetProperty("Height")
                        pi.SetValue(rowObjects(i), _SingleHi, Nothing)
                    Catch
                        'Throw New ArgumentException("invalid row index")
                    End Try
                Next
            End If
        End Sub
    End Class 'myDataGridRowHeightSetter 

End Namespace
