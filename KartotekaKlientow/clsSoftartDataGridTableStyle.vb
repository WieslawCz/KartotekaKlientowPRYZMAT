Option Strict Off
Option Explicit On 

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Namespace Softart
    Public Class myDataGridTextBoxColumn
        Inherits DataGridTextBoxColumn

        Private SelectedRow As Integer
        Private SelectedColumn As Integer ' column where this columnstyle is located...
        'Fields
        'Constructors
        'Events
        'Methods
        Public Sub New()
            SelectedColumn = -2
        End Sub

        '*******************************************
        '*** Wyró¿nienie kursora
        '*******************************************

        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, ByVal bounds As Rectangle, ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal backBrush As Brush, ByVal foreBrush As Brush, ByVal alignToRight As Boolean)
            Try
                Dim grid As DataGrid = Me.DataGridTableStyle.DataGrid
                'first time set the column properly
                If SelectedColumn = -2 Then
                    Dim i As Integer
                    i = Me.DataGridTableStyle.GridColumnStyles.IndexOf(Me)
                    If i > -1 Then
                        SelectedColumn = i
                    End If
                End If

                If grid.CurrentRowIndex = rowNum And grid.CurrentCell.ColumnNumber = SelectedColumn Then
                    'backBrush = New LinearGradientBrush(bounds, Color.FromArgb(255, 200, 200), Color.FromArgb(128, 20, 20), LinearGradientMode.BackwardDiagonal)
                    'foreBrush = New SolidBrush(Color.White)
                    backBrush = New LinearGradientBrush(bounds, Color.FromArgb(200, 255, 200), Color.FromArgb(20, 64, 20), LinearGradientMode.BackwardDiagonal)
                    foreBrush = New SolidBrush(Color.White)
                End If
            Catch ex As Exception
                ' empty catch 
            Finally
                ' make sure the base class gets called to do the drawing with
                ' the possibly changed brushes
                MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
            End Try

        End Sub

        '*******************************************
        '*** wy³¹czamy mo¿liwoœæ edycji komórki
        '*******************************************

        Protected Overloads Overrides Sub Edit(ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal bounds As Rectangle, ByVal [readOnly] As Boolean, ByVal instantText As String, ByVal cellIsVisible As Boolean)
            'make sure selectedrow is valid
            'jesli AllowAdd=true to o jeden wiersz wiecej...
            'If (SelectedRow > -1) And (SelectedRow < source.List.Count + 1) Then
            'jesli AllowAdd=false - tylko tyle wierszy ile w tabeli...
            If (SelectedRow > -1) And (SelectedRow < source.List.Count) Then
                Me.DataGridTableStyle.DataGrid.UnSelect(SelectedRow)
            End If
            SelectedRow = rowNum
            Me.DataGridTableStyle.DataGrid.Select(SelectedRow)
        End Sub
    End Class



    Public Class myDataGridNumColumn
        Inherits DataGridTextBoxColumn

        Private SelectedRow As Integer
        Private SelectedColumn As Integer ' column where this columnstyle is located...
        'Fields
        'Constructors
        'Events
        'Methods
        Public Sub New()
            SelectedColumn = -2
        End Sub

        '*******************************************
        '*** Wyró¿nienie kursora
        '*******************************************

        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, ByVal bounds As Rectangle, ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal backBrush As Brush, ByVal foreBrush As Brush, ByVal alignToRight As Boolean)
            Try
                Dim grid As DataGrid = Me.DataGridTableStyle.DataGrid
                'first time set the column properly
                If SelectedColumn = -2 Then
                    Dim i As Integer
                    i = Me.DataGridTableStyle.GridColumnStyles.IndexOf(Me)
                    If i > -1 Then
                        SelectedColumn = i
                    End If
                End If

                If grid.CurrentRowIndex = rowNum And grid.CurrentCell.ColumnNumber = SelectedColumn Then
                    'backBrush = New LinearGradientBrush(bounds, Color.FromArgb(255, 200, 200), Color.FromArgb(128, 20, 20), LinearGradientMode.BackwardDiagonal)
                    'foreBrush = New SolidBrush(Color.White)
                    backBrush = New LinearGradientBrush(bounds, Color.FromArgb(200, 255, 200), Color.FromArgb(20, 64, 20), LinearGradientMode.BackwardDiagonal)
                    foreBrush = New SolidBrush(Color.White)
                End If
            Catch ex As Exception
                ' empty catch 
            Finally
                ' make sure the base class gets called to do the drawing with
                ' the possibly changed brushes
                MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
            End Try

        End Sub

        '*******************************************
        '*** wy³¹czamy mo¿liwoœæ edycji komórki
        '*******************************************

        Protected Overloads Overrides Sub Edit(ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal bounds As Rectangle, ByVal [readOnly] As Boolean, ByVal instantText As String, ByVal cellIsVisible As Boolean)
            'make sure selectedrow is valid
            'jesli AllowAdd=true to o jeden wiersz wiecej...
            'If (SelectedRow > -1) And (SelectedRow < source.List.Count + 1) Then
            'jesli AllowAdd=false - tylko tyle wierszy ile w tabeli...
            If (SelectedRow > -1) And (SelectedRow < source.List.Count) Then
                Me.DataGridTableStyle.DataGrid.UnSelect(SelectedRow)
            End If
            SelectedRow = rowNum
            Me.DataGridTableStyle.DataGrid.Select(SelectedRow)
        End Sub
    End Class




    Public Class myDataGridBoolColumn
        Inherits DataGridBoolColumn

        Private SelectedRow As Integer
        Private SelectedColumn As Integer ' column where this columnstyle is located...
        'Fields
        'Constructors
        'Events
        'Methods
        Public Sub New()
            SelectedColumn = -2
        End Sub

        '*******************************************
        '*** Wyró¿nienie kursora
        '*******************************************

        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, ByVal bounds As Rectangle, ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal backBrush As Brush, ByVal foreBrush As Brush, ByVal alignToRight As Boolean)
            Try
                Dim grid As DataGrid = Me.DataGridTableStyle.DataGrid
                'first time set the column properly
                If SelectedColumn = -2 Then
                    Dim i As Integer
                    i = Me.DataGridTableStyle.GridColumnStyles.IndexOf(Me)
                    If i > -1 Then
                        SelectedColumn = i
                    End If
                End If

                If grid.CurrentRowIndex = rowNum And grid.CurrentCell.ColumnNumber = SelectedColumn Then
                    'backBrush = New LinearGradientBrush(bounds, Color.FromArgb(255, 200, 200), Color.FromArgb(128, 20, 20), LinearGradientMode.BackwardDiagonal)
                    'foreBrush = New SolidBrush(Color.White)
                    backBrush = New LinearGradientBrush(bounds, Color.FromArgb(200, 255, 200), Color.FromArgb(20, 64, 20), LinearGradientMode.BackwardDiagonal)
                    foreBrush = New SolidBrush(Color.White)
                End If
            Catch ex As Exception
                ' empty catch 
            Finally
                ' make sure the base class gets called to do the drawing with
                ' the possibly changed brushes
                MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
            End Try

        End Sub

        '*******************************************
        '*** wy³¹czamy mo¿liwoœæ edycji komórki
        '*******************************************

        Protected Overloads Overrides Sub Edit(ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal bounds As Rectangle, ByVal [readOnly] As Boolean, ByVal instantText As String, ByVal cellIsVisible As Boolean)
            'make sure selectedrow is valid
            If (SelectedRow > -1) And (SelectedRow < source.List.Count) Then
                Me.DataGridTableStyle.DataGrid.UnSelect(SelectedRow)
            End If
            SelectedRow = rowNum
            Me.DataGridTableStyle.DataGrid.Select(SelectedRow)
        End Sub
    End Class
End Namespace
