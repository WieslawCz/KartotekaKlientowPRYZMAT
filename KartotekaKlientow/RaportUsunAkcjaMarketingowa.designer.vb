<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RaportUsunAkcjaMarketingowa
    'Inherits System.Windows.Forms.Form
    Inherits Softart.myDropShadowForm

    Dim _Ident As String
    Dim _dvKlienci As DataView
    Dim _dsAkcjeSpec As DataSet
    Dim _AktualizujAkcjeSpec As delegateUsunZ

    Public Sub New(ByVal IdAkcji As String, _
                   ByVal dvKlienci As DataView, _
                   ByVal dsAkcjeSpec As DataSet, _
                   ByVal Aktualizuj As delegateUsunZ)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        _Ident = IdAkcji
        _dvKlienci = dvKlienci
        _dsAkcjeSpec = dsAkcjeSpec
        _AktualizujAkcjeSpec = Aktualizuj
    End Sub


    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Progres = New System.Windows.Forms.ProgressBar()
        Me.lblOperacja = New System.Windows.Forms.Label()
        Me.lblFunkcja = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Progres)
        Me.Panel1.Controls.Add(Me.lblOperacja)
        Me.Panel1.Controls.Add(Me.lblFunkcja)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(347, 66)
        Me.Panel1.TabIndex = 0
        '
        'Progres
        '
        Me.Progres.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Progres.ForeColor = System.Drawing.Color.Navy
        Me.Progres.Location = New System.Drawing.Point(11, 45)
        Me.Progres.Name = "Progres"
        Me.Progres.Size = New System.Drawing.Size(324, 10)
        Me.Progres.Step = 1
        Me.Progres.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.Progres.TabIndex = 2
        '
        'lblOperacja
        '
        Me.lblOperacja.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblOperacja.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblOperacja.ForeColor = System.Drawing.Color.Navy
        Me.lblOperacja.Location = New System.Drawing.Point(8, 25)
        Me.lblOperacja.Name = "lblOperacja"
        Me.lblOperacja.Size = New System.Drawing.Size(327, 13)
        Me.lblOperacja.TabIndex = 1
        Me.lblOperacja.Text = "Label1"
        Me.lblOperacja.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblFunkcja
        '
        Me.lblFunkcja.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFunkcja.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblFunkcja.ForeColor = System.Drawing.Color.Navy
        Me.lblFunkcja.Location = New System.Drawing.Point(8, 8)
        Me.lblFunkcja.Name = "lblFunkcja"
        Me.lblFunkcja.Size = New System.Drawing.Size(327, 13)
        Me.lblFunkcja.TabIndex = 0
        Me.lblFunkcja.Text = "Label1"
        Me.lblFunkcja.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Timer1
        '
        '
        'RaportUsunAkcjaMarketingowa
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(367, 82)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.Name = "RaportUsunAkcjaMarketingowa"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblFunkcja As System.Windows.Forms.Label
    Friend WithEvents Progres As System.Windows.Forms.ProgressBar
    Friend WithEvents lblOperacja As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
