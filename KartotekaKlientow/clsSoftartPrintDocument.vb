Option Strict Off
Option Explicit On 

Imports System.Drawing.Printing

Namespace Softart
    Public Class myPrintDocument
        Inherits PrintDocument

        Private _PageNumber As Integer = 0
        Private _LineNumber As Integer = 0
        Private _NextRowNumber As Integer = 0

        'Fields
        'Constructors
        'Events
        'Methods
        Public Sub New()
            'Warning: Implementation not found
        End Sub

        '******************************************
        '** obsluga licznikow
        '******************************************

        Public Property PageNumber() As Long
            Get
                Return _PageNumber
            End Get
            Set(ByVal Value As Long)
                _PageNumber = Value
            End Set
        End Property

        Public Property LineNumber() As Long
            Get
                Return _LineNumber
            End Get
            Set(ByVal Value As Long)
                _LineNumber = Value
            End Set
        End Property

        Public Property NextRowNumber() As Long
            Get
                Return _NextRowNumber
            End Get
            Set(ByVal Value As Long)
                _NextRowNumber = Value
            End Set
        End Property
    End Class
End Namespace
