
Public Class _frmOAutorach
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents picIkona As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cmdEMail As System.Windows.Forms.Button
    Friend WithEvents cmdWWW As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cmdPowrot As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(_frmOAutorach))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.picIkona = New System.Windows.Forms.PictureBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdEMail = New System.Windows.Forms.Button()
        Me.cmdWWW = New System.Windows.Forms.Button()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        CType(Me.picIkona, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Navy
        Me.Label1.Location = New System.Drawing.Point(152, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(208, 88)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Obs³uga Kartoteki Klientów firmy PRYZMAT"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'picIkona
        '
        Me.picIkona.Location = New System.Drawing.Point(16, 12)
        Me.picIkona.Name = "picIkona"
        Me.picIkona.Size = New System.Drawing.Size(129, 116)
        Me.picIkona.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picIkona.TabIndex = 1
        Me.picIkona.TabStop = False
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.Purple
        Me.Label2.Location = New System.Drawing.Point(153, 112)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(204, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Tag = "2wsx`t5tv9ik,"
        Me.Label2.Text = "Wersja 1.00.00"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(0, 136)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(440, 1)
        Me.Label3.TabIndex = 3
        '
        'cmdEMail
        '
        Me.cmdEMail.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEMail.Location = New System.Drawing.Point(245, 175)
        Me.cmdEMail.Name = "cmdEMail"
        Me.cmdEMail.Size = New System.Drawing.Size(181, 23)
        Me.cmdEMail.TabIndex = 4
        Me.cmdEMail.Text = "opole@softart.com.pl"
        '
        'cmdWWW
        '
        Me.cmdWWW.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdWWW.Location = New System.Drawing.Point(245, 207)
        Me.cmdWWW.Name = "cmdWWW"
        Me.cmdWWW.Size = New System.Drawing.Size(181, 23)
        Me.cmdWWW.TabIndex = 5
        Me.cmdWWW.Text = "www.softart.com.pl"
        '
        'cmdPowrot
        '
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(245, 239)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(181, 24)
        Me.cmdPowrot.TabIndex = 6
        Me.cmdPowrot.Text = "Powrót"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Navy
        Me.Label4.Location = New System.Drawing.Point(8, 144)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(224, 16)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Autorem programu jest"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Navy
        Me.Label5.Location = New System.Drawing.Point(8, 165)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(224, 16)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "EiP Softart spó³ka z o.o."
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'picLogo
        '
        Me.picLogo.Image = CType(resources.GetObject("picLogo.Image"), System.Drawing.Image)
        Me.picLogo.Location = New System.Drawing.Point(11, 184)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(221, 31)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 9
        Me.picLogo.TabStop = False
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Navy
        Me.Label6.Location = New System.Drawing.Point(8, 223)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(224, 16)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "PL45-222 Opole, ul. Oleska 97 H"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(8, 255)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(224, 16)
        Me.Label8.TabIndex = 12
        Me.Label8.Text = "www.softart.com.pl"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Navy
        Me.Label9.Location = New System.Drawing.Point(8, 239)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(224, 16)
        Me.Label9.TabIndex = 13
        Me.Label9.Text = "eMail: opole@softart.com.pl"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Navy
        Me.Label12.Location = New System.Drawing.Point(8, 271)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(224, 16)
        Me.Label12.TabIndex = 16
        Me.Label12.Text = "NIP: 754-033-76-69"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '_frmOAutorach
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(438, 294)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.cmdEMail)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.cmdWWW)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.picIkona)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "_frmOAutorach"
        Me.ShowInTaskbar = False
        Me.Text = "O autorach"
        CType(Me.picIkona, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub cmdWWW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWWW.Click
        System.Diagnostics.Process.Start( _
           "C:\Program Files\Internet Explorer\IExplore.exe", "http://www.softart.com.pl")
    End Sub

    Private Sub cmdEMail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEMail.Click
        If TestujParametryeMail() Then
            Select Case ParL_eMailService
                Case "Wbudowany"
                    Dim Form As New SendeMail(SoftarteMail, "Program " & oNazwaProgramu, "", "")
                    Form.Show()
                    Form = Nothing
                Case "MS Outlook"
                    SendByOutlook(SoftarteMail, "Program " & oNazwaProgramu, "", "", "Edytuj")
                Case "Mozilla Thunderbird"
                    SendByThunderbird(SoftarteMail, "", "", "Program " & oNazwaProgramu, "", "")
                Case Else
            End Select
        End If
    End Sub

    Private Sub OAutorach_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.MrJones
        Me.picLogo.Image = My.Resources.logomini
        Me.cmdPowrot.Image = My.Resources._EXIT.ToBitmap
        Me.cmdWWW.Image = My.Resources.Globe.ToBitmap
        Me.cmdEMail.Image = My.Resources.KOPERTA.ToBitmap
        '----------------------------
        'Numer wersji
        Label2.Text = "Wersja " & Trim(Str(System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMajorPart)) & _
                         "." & Trim(Str(System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMinorPart)) & _
                         "." & Trim(Str(System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileBuildPart))
    End Sub
End Class
