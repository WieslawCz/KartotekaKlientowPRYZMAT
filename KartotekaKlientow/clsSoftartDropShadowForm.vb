Option Strict Off
Option Explicit On

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Namespace Softart
    Public Class myDropShadowForm
        Inherits System.Windows.Forms.Form

        '====================================================
        'zmienne do œledzenia JAK ZAMKNIETO FORME
        Private _closeClick As Boolean

        'Private components As System.ComponentModel.Container
        'kody sposobu zamkniecia formy
        Public Const SC_CLOSE As Integer = 61536
        Public Const WM_SYSCOMMAND As Integer = 274

        'Fields
        'Constructors
        'Events
        'Methods

        Public Sub New()

            MyBase.New()

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call
            _closeClick = False

        End Sub

        'ta metoda pozwoli na sprawdzenie sposobu zamkniecia formy
        Protected Overloads Overrides Sub WndProc(ByRef m As Message)
            If m.Msg = WM_SYSCOMMAND AndAlso m.WParam.ToInt32 = SC_CLOSE Then
                'zamkniecie przez COntrolBox (krzyzyk w prawym gornym rogu...)
                Me._closeClick = True
            Else
                'Me._closeClick = False
            End If
            MyBase.WndProc(m)
        End Sub

        'wlasciwosc sluzaca do sprawdzenia sposobu zamkniecia formy
        Public Property ClosedByControlBox() As Boolean
            Get
                Return _closeClick
            End Get
            Set(ByVal value As Boolean)
                _closeClick = value
            End Set
        End Property



        ''====================================================
        ''== w Formie treba dodaæ kod testowania zamkniêcia
        ''====================================================
        'Private Sub EdytujOferteOps_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        '    If Me.ClosedByControlBox() Then
        '        e.Cancel = True     'nie pozwalaj zamkn¹c forme
        '        MessageBox.Show("Proszê zamkn¹æ formê klawiszami ZAPISZ lub WYCOFAJ SIÊ", _
        '            "Uwaga :", _
        '            System.Windows.Forms.MessageBoxButtons.OK, _
        '            MessageBoxIcon.Information, _
        '            MessageBoxDefaultButton.Button1)
        '        Me.ClosedByControlBox() = False
        '    Else
        '        'MsgBox("ZAPISZ lub WYCOFAJ SIÊ")
        '    End If
        'End Sub
        '====================================================






        ''**********************************************
        ''** obs³uga cieniowania formy
        ''**********************************************

        'Private dropShadow As Boolean = True

        'Protected Overrides ReadOnly Property CreateParams() As CreateParams
        '    Get
        '        Dim params1 As CreateParams = MyBase.CreateParams
        '        If (Me.dropShadow AndAlso DropShadowSupported) Then
        '            params1.ClassStyle = (params1.ClassStyle Or 131072)
        '        End If
        '        Return params1
        '    End Get
        'End Property
        'Public Shared ReadOnly Property DropShadowSupported() As Boolean
        '    Get
        '        Return IsWindowsXPOrAbove
        '    End Get
        'End Property
        'Public Shared ReadOnly Property IsWindowsXPOrAbove() As Boolean
        '    Get
        '        Dim system1 As OperatingSystem = Environment.OSVersion
        '        Return ((system1.Platform = PlatformID.Win32NT) AndAlso (system1.Version.CompareTo(New Version(5, 1, 0, 0)) >= 0))
        '    End Get
        'End Property

        ''**********************************************

        Private Sub InitializeComponent()
            Me.SuspendLayout()
            '
            'myDropShadowForm
            '
            Me.ClientSize = New System.Drawing.Size(292, 266)
            Me.Name = "myDropShadowForm"
            Me.ResumeLayout(False)

        End Sub

        Private Sub myDropShadowForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        End Sub
    End Class
End Namespace
