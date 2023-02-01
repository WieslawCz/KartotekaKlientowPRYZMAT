Option Strict Off
Option Explicit On 

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace Softart
    Public Class MyDataGrid
        Inherits DataGrid

        Private Const OFFSET As Integer = 5
        Private oldSelectedRow As Integer

        'Fields
        'Constructors
        'Events
        'Methods
        Public Sub New()
            'Warning: Implementation not found
        End Sub

        'nadpisujemy OnCurrentCellChanged
        Protected Overloads Overrides Sub OnCurrentCellChanged(ByVal e As EventArgs)
            MyBase.OnCurrentCellChanged(e)
            If (Control.MouseButtons <> Windows.Forms.MouseButtons.Left) Then
                Dim rect As Rectangle
                rect = Me.GetCellBounds(Me.CurrentCell)
                Dim e1 As MouseEventArgs
                e1 = New MouseEventArgs(Windows.Forms.MouseButtons.Left, 0, OFFSET, (rect.Y + OFFSET), 0)
                OnMouseDown(e1)
            End If

        End Sub

        'nadpisujemy metode OnMouseMove
        Protected Overloads Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)
            'nie pozwol na zmiane wymiarow wiersza
            'Dim hti As DataGrid.HitTestInfo
            'hti = Me.HitTest(New Point(e.X, e.Y))
            'If (hti.Type = DataGrid.HitTestType.RowResize) Then
            'Return                'no baseclass call
            'End If
            '-----------------------------------------
            'don't call the base class if left mouse down
            If (e.Button <> Windows.Forms.MouseButtons.Left) Then
                MyBase.OnMouseMove(e)
            End If
        End Sub

        'nadpisujemu metode OnMouseDown
        Protected Overloads Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
            'don't call the base class if in header
            Dim hti As DataGrid.HitTestInfo
            hti = Me.HitTest(New Point(e.X, e.Y))

            Select Case hti.Type
                Case DataGrid.HitTestType.Cell
                    'sprawdz czy pojedynczy select..
                    If (oldSelectedRow > -(1)) And (oldSelectedRow < Me.DataSource.count) Then
                        Me.UnSelect(oldSelectedRow)
                    End If
                    oldSelectedRow = -(1)
                    'oldSelectedRow = hti.Row

                    'MyBase.OnMouseDown(e)
                    'nie trzeba edytowac...
                    Dim e1 As MouseEventArgs
                    e1 = New MouseEventArgs(e.Button, e.Clicks, OFFSET, e.Y, e.Delta)
                    MyBase.OnMouseDown(e1)
                Case DataGrid.HitTestType.RowHeader
                    If (oldSelectedRow > -(1)) And (oldSelectedRow < Me.DataSource.count) Then
                        Me.UnSelect(oldSelectedRow)
                    End If
                    If ((Control.ModifierKeys And Keys.Shift) = 0) Then
                        MyBase.OnMouseDown(e)
                    Else
                        Me.CurrentCell = New DataGridCell(hti.Row, hti.Column)
                    End If
                    Me.Select(hti.Row)
                    oldSelectedRow = hti.Row
                Case DataGrid.HitTestType.RowResize
                    'sprawdz czy pojedynczy select..
                    If (oldSelectedRow > -(1)) And (oldSelectedRow < Me.DataSource.count) Then
                        Me.UnSelect(oldSelectedRow)
                    End If
                    oldSelectedRow = -(1)
                    'oldSelectedRow = hti.Row
                Case DataGrid.HitTestType.ColumnResize
                    'sprawdz czy pojedynczy select..
                    If (oldSelectedRow > -(1)) And (oldSelectedRow < Me.DataSource.count) Then
                        Me.UnSelect(oldSelectedRow)
                    End If
                    oldSelectedRow = -(1)
                    'oldSelectedRow = hti.Row
                Case Else
                    'pozostale...
                    MyBase.OnMouseDown(e)
            End Select
        End Sub

        '***************************************************
        '** przewin DataGrid do konkretnego wiersza Row
        '*****************************************************

        Sub ScrollToRow(ByVal row As Integer)
            If Not Me.DataSource Is Nothing Then
                Me.GridVScrolled(Me, New ScrollEventArgs(ScrollEventType.LargeIncrement, row))
            End If
        End Sub

        Sub ScrollToColumn(ByVal col As Integer)
            If Not Me.DataSource Is Nothing Then
                Me.GridHScrolled(Me, New ScrollEventArgs(ScrollEventType.LargeIncrement, col))
            End If
        End Sub
    End Class
End Namespace
