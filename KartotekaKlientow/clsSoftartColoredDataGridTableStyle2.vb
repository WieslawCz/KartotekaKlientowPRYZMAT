Option Strict Off
Option Explicit On 

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

'delegate do procedury definiowania koloru w datagrid
'Procedura musi miec nast sygnature
Public Delegate Function delegateDefColor(ByVal row As Long) As Integer


Namespace Softart
    Public Class myColoredDataGridTextBoxColumnB
        Inherits myDataGridTextBoxColumn

        Dim _DefiniujBackColor As delegateDefColor

        Public Sub New(ByVal deleg As delegateDefColor)
            'zapamietaj adres delegata
            _DefiniujBackColor = deleg
        End Sub


        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, _
                                                ByVal bounds As Rectangle, _
                                                ByVal source As CurrencyManager, _
                                                ByVal rowNum As Integer, _
                                                ByVal backBrush As Brush, _
                                                ByVal foreBrush As Brush, _
                                                ByVal alignToRight As Boolean)

            Try
                backBrush = New SolidBrush(Color.FromArgb(_DefiniujBackColor(rowNum)))
                'foreBrush = New SolidBrush(Color.FromArgb(_DefiniujForeColor(rowNum)))

            Catch ex As Exception
                ' empty catch 
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace, _
                '    "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ' make sure the base class gets called to do the drawing with 
                ' the possibly changed brushes 
                MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
            End Try

            '--------------------------------------------------
            ''pobierz wartoœæ wyswietlanej komórki
            'Dim bdel As Boolean = GetColumnValueAtRow(source, rowNum)
            ''w zaleznosci od wartosci zastosuj kolor
            'If bdel = True Then
            '    backBrush = New SolidBrush(Color.FromArgb(_DefiniujBackColor(rowNum)))
            'Else
            '    backBrush = New SolidBrush(Color.FromArgb(_DefiniujBackColor(rowNum)))
            'End If
            'g.FillRectangle(backBrush, bounds.X, bounds.Y, bounds.Width, bounds.Height)

            ''okresl font i pisz wartosc
            'Dim font As System.Drawing.Font = New Font(System.Drawing.FontFamily.GenericSansSerif, 8.25)
            'g.DrawString(bdel.ToString(), Font, Brushes.Black, bounds.X, bounds.Y)
            '--------------------------------------------------
        End Sub
    End Class






    Public Class myColoredDataGridTextBoxColumnF
        Inherits myDataGridTextBoxColumn

        Dim _DefiniujForeColor As delegateDefColor

        Public Sub New(ByVal deleg As delegateDefColor)
            'zapamietaj adres delegata
            _DefiniujForeColor = deleg
        End Sub


        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, _
                                                ByVal bounds As Rectangle, _
                                                ByVal source As CurrencyManager, _
                                                ByVal rowNum As Integer, _
                                                ByVal backBrush As Brush, _
                                                ByVal foreBrush As Brush, _
                                                ByVal alignToRight As Boolean)

            Try
                foreBrush = New SolidBrush(Color.FromArgb(_DefiniujForeColor(rowNum)))

            Catch ex As Exception
                ' empty catch 
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace, _
                '    "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ' make sure the base class gets called to do the drawing with 
                ' the possibly changed brushes 
                MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
            End Try
        End Sub
    End Class




    Public Class myColoredDataGridTextBoxColumnBF
        Inherits myDataGridTextBoxColumn

        Dim _DefiniujForeColor As delegateDefColor
        Dim _DefiniujBackColor As delegateDefColor

        Public Sub New(ByVal delegB As delegateDefColor, ByVal delegF As delegateDefColor)
            'zapamietaj adres delegata
            _DefiniujForeColor = delegF
            _DefiniujBackColor = delegB
        End Sub


        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, _
                                                ByVal bounds As Rectangle, _
                                                ByVal source As CurrencyManager, _
                                                ByVal rowNum As Integer, _
                                                ByVal backBrush As Brush, _
                                                ByVal foreBrush As Brush, _
                                                ByVal alignToRight As Boolean)

            Try
                foreBrush = New SolidBrush(Color.FromArgb(_DefiniujForeColor(rowNum)))
                backBrush = New SolidBrush(Color.FromArgb(_DefiniujBackColor(rowNum)))

            Catch ex As Exception
                ' empty catch 
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace, _
                '    "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ' make sure the base class gets called to do the drawing with 
                ' the possibly changed brushes 
                MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
            End Try
        End Sub
    End Class

    '====================================================

    Public Class myColoredDataGridTextBoxColumnF_Analiza3
        Inherits myDataGridTextBoxColumn

        Dim _DefiniujForeColor As delegateDefColor

        Public Sub New(ByVal deleg As delegateDefColor)
            'zapamietaj adres delegata
            _DefiniujForeColor = deleg
        End Sub


        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, _
                                                ByVal bounds As Rectangle, _
                                                ByVal source As CurrencyManager, _
                                                ByVal rowNum As Integer, _
                                                ByVal backBrush As Brush, _
                                                ByVal foreBrush As Brush, _
                                                ByVal alignToRight As Boolean)

            Try
                foreBrush = New SolidBrush(Color.FromArgb(_DefiniujForeColor(rowNum)))
                backBrush = New SolidBrush(System.Drawing.Color.Azure)
                'backBrush = New SolidBrush(Color.FromArgb(RGB(255, 192, 192)))

            Catch ex As Exception
                ' empty catch 
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace, _
                '    "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ' make sure the base class gets called to do the drawing with 
                ' the possibly changed brushes 
                MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
            End Try
        End Sub
    End Class

    Public Class myColoredDataGridTextBoxColumnF_Analiza2
        Inherits myDataGridTextBoxColumn

        Dim _DefiniujForeColor As delegateDefColor

        Public Sub New(ByVal deleg As delegateDefColor)
            'zapamietaj adres delegata
            _DefiniujForeColor = deleg
        End Sub


        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, _
                                                ByVal bounds As Rectangle, _
                                                ByVal source As CurrencyManager, _
                                                ByVal rowNum As Integer, _
                                                ByVal backBrush As Brush, _
                                                ByVal foreBrush As Brush, _
                                                ByVal alignToRight As Boolean)

            Try
                foreBrush = New SolidBrush(Color.FromArgb(_DefiniujForeColor(rowNum)))
                backBrush = New SolidBrush(System.Drawing.Color.LightCyan)
                'backBrush = New SolidBrush(Color.FromArgb(RGB(192, 255, 192)))

            Catch ex As Exception
                ' empty catch 
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace, _
                '    "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ' make sure the base class gets called to do the drawing with 
                ' the possibly changed brushes 
                MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
            End Try
        End Sub
    End Class

    Public Class myColoredDataGridTextBoxColumnF_Analiza1
        Inherits myDataGridTextBoxColumn

        Dim _DefiniujForeColor As delegateDefColor

        Public Sub New(ByVal deleg As delegateDefColor)
            'zapamietaj adres delegata
            _DefiniujForeColor = deleg
        End Sub


        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, _
                                                ByVal bounds As Rectangle, _
                                                ByVal source As CurrencyManager, _
                                                ByVal rowNum As Integer, _
                                                ByVal backBrush As Brush, _
                                                ByVal foreBrush As Brush, _
                                                ByVal alignToRight As Boolean)

            Try
                foreBrush = New SolidBrush(Color.FromArgb(_DefiniujForeColor(rowNum)))
                backBrush = New SolidBrush(System.Drawing.Color.PaleTurquoise)
                'backBrush = New SolidBrush(Color.FromArgb(RGB(192, 192, 255)))

            Catch ex As Exception
                ' empty catch 
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace, _
                '    "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ' make sure the base class gets called to do the drawing with 
                ' the possibly changed brushes 
                MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
            End Try
        End Sub
    End Class


    Public Class myColoredDataGridTextBoxColumnF_Analiza0
        Inherits myDataGridTextBoxColumn

        Dim _DefiniujForeColor As delegateDefColor

        Public Sub New(ByVal deleg As delegateDefColor)
            'zapamietaj adres delegata
            _DefiniujForeColor = deleg
        End Sub


        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, _
                                                ByVal bounds As Rectangle, _
                                                ByVal source As CurrencyManager, _
                                                ByVal rowNum As Integer, _
                                                ByVal backBrush As Brush, _
                                                ByVal foreBrush As Brush, _
                                                ByVal alignToRight As Boolean)

            Try
                foreBrush = New SolidBrush(Color.FromArgb(_DefiniujForeColor(rowNum)))
                backBrush = New SolidBrush(System.Drawing.Color.WhiteSmoke)
                'backBrush = New SolidBrush(Color.FromArgb(RGB(192, 192, 255)))

            Catch ex As Exception
                ' empty catch 
                'MessageBox.Show("W programie wyst¹pi³ b³¹d :" & vbCrLf & ex.Message & vbCrLf & ex.StackTrace, _
                '    "Uwaga :", _
                '    System.Windows.Forms.MessageBoxButtons.OK, _
                '    MessageBoxIcon.Information, _
                '    MessageBoxDefaultButton.Button1)
            Finally
                ' make sure the base class gets called to do the drawing with 
                ' the possibly changed brushes 
                MyBase.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight)
            End Try
        End Sub
    End Class


End Namespace
