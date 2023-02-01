Public Class _frmPolaczenieSieciowe
    Inherits System.Windows.Forms.Form

    ' Declare the API function.
    Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Int32, ByVal dwReserved As Int32) As Int32

    ' Define the possible types of connections.
    Private Enum ConnectStates
        LAN = &H2
        Modem = &H1
        Proxy = &H4
        Offline = &H20
        Configured = &H40
        RasInstalled = &H10
    End Enum

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
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents lblHead1 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblOpis As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblTenKomputer As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblSposobPolaczenia As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblAdresIP As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblAdresIPv6 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblNazwaHosta As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdPowrot = New System.Windows.Forms.Button()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblAdresIPv6 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblNazwaHosta = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblAdresIP = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblSposobPolaczenia = New System.Windows.Forms.Label()
        Me.lblTenKomputer = New System.Windows.Forms.Label()
        Me.lblOpis = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblHead1 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdPowrot
        '
        Me.cmdPowrot.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPowrot.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPowrot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPowrot.Location = New System.Drawing.Point(460, 278)
        Me.cmdPowrot.Name = "cmdPowrot"
        Me.cmdPowrot.Size = New System.Drawing.Size(88, 23)
        Me.cmdPowrot.TabIndex = 2
        Me.cmdPowrot.Text = "Powrót"
        Me.cmdPowrot.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'picLogo
        '
        Me.picLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picLogo.Location = New System.Drawing.Point(8, 278)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(104, 24)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 11
        Me.picLogo.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lblAdresIPv6)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.lblNazwaHosta)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.lblAdresIP)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.lblSposobPolaczenia)
        Me.Panel1.Controls.Add(Me.lblTenKomputer)
        Me.Panel1.Controls.Add(Me.lblOpis)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.lblHead1)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.ForeColor = System.Drawing.Color.Purple
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(546, 264)
        Me.Panel1.TabIndex = 12
        '
        'lblAdresIPv6
        '
        Me.lblAdresIPv6.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAdresIPv6.BackColor = System.Drawing.SystemColors.ControlLight
        Me.lblAdresIPv6.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.lblAdresIPv6.ForeColor = System.Drawing.Color.Purple
        Me.lblAdresIPv6.Location = New System.Drawing.Point(112, 115)
        Me.lblAdresIPv6.Name = "lblAdresIPv6"
        Me.lblAdresIPv6.Size = New System.Drawing.Size(423, 131)
        Me.lblAdresIPv6.TabIndex = 12
        Me.lblAdresIPv6.Text = "  Ten komputer"
        '
        'Label6
        '
        Me.Label6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label6.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Label6.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label6.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label6.Location = New System.Drawing.Point(0, 115)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(120, 131)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "  Adres IP v.6"
        '
        'lblNazwaHosta
        '
        Me.lblNazwaHosta.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNazwaHosta.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.lblNazwaHosta.ForeColor = System.Drawing.Color.Purple
        Me.lblNazwaHosta.Location = New System.Drawing.Point(112, 99)
        Me.lblNazwaHosta.Name = "lblNazwaHosta"
        Me.lblNazwaHosta.Size = New System.Drawing.Size(423, 16)
        Me.lblNazwaHosta.TabIndex = 10
        Me.lblNazwaHosta.Text = "  Ten komputer"
        '
        'Label4
        '
        Me.Label4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label4.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label4.Location = New System.Drawing.Point(3, 99)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(120, 16)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "  Nazwa Hosta"
        '
        'lblAdresIP
        '
        Me.lblAdresIP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAdresIP.BackColor = System.Drawing.SystemColors.ControlLight
        Me.lblAdresIP.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.lblAdresIP.ForeColor = System.Drawing.Color.Purple
        Me.lblAdresIP.Location = New System.Drawing.Point(112, 48)
        Me.lblAdresIP.Name = "lblAdresIP"
        Me.lblAdresIP.Size = New System.Drawing.Size(423, 51)
        Me.lblAdresIP.TabIndex = 8
        Me.lblAdresIP.Text = "  Ten komputer"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.ControlLight
        Me.Label3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label3.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label3.Location = New System.Drawing.Point(0, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 51)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "  Adres IP po³¹czenia"
        '
        'lblSposobPolaczenia
        '
        Me.lblSposobPolaczenia.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSposobPolaczenia.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.lblSposobPolaczenia.ForeColor = System.Drawing.Color.Purple
        Me.lblSposobPolaczenia.Location = New System.Drawing.Point(112, 32)
        Me.lblSposobPolaczenia.Name = "lblSposobPolaczenia"
        Me.lblSposobPolaczenia.Size = New System.Drawing.Size(423, 16)
        Me.lblSposobPolaczenia.TabIndex = 6
        Me.lblSposobPolaczenia.Text = "  Ten komputer"
        '
        'lblTenKomputer
        '
        Me.lblTenKomputer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTenKomputer.BackColor = System.Drawing.SystemColors.ControlLight
        Me.lblTenKomputer.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.lblTenKomputer.ForeColor = System.Drawing.Color.Purple
        Me.lblTenKomputer.Location = New System.Drawing.Point(112, 16)
        Me.lblTenKomputer.Name = "lblTenKomputer"
        Me.lblTenKomputer.Size = New System.Drawing.Size(423, 16)
        Me.lblTenKomputer.TabIndex = 4
        Me.lblTenKomputer.Text = "  Ten komputer"
        '
        'lblOpis
        '
        Me.lblOpis.BackColor = System.Drawing.SystemColors.ControlLight
        Me.lblOpis.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.lblOpis.ForeColor = System.Drawing.SystemColors.Desktop
        Me.lblOpis.Location = New System.Drawing.Point(0, 16)
        Me.lblOpis.Name = "lblOpis"
        Me.lblOpis.Size = New System.Drawing.Size(112, 16)
        Me.lblOpis.TabIndex = 3
        Me.lblOpis.Text = "  Ten komputer"
        '
        'Label2
        '
        Me.Label2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label2.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Location = New System.Drawing.Point(0, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(120, 16)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "  Sposób po³¹czenia"
        '
        'lblHead1
        '
        Me.lblHead1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblHead1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.lblHead1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.lblHead1.Location = New System.Drawing.Point(0, 0)
        Me.lblHead1.Name = "lblHead1"
        Me.lblHead1.Size = New System.Drawing.Size(112, 16)
        Me.lblHead1.TabIndex = 1
        Me.lblHead1.Text = "  Parametr"
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Location = New System.Drawing.Point(112, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(423, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "  Wartoœæ"
        '
        '_frmPolaczenieSieciowe
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(560, 308)
        Me.Controls.Add(Me.picLogo)
        Me.Controls.Add(Me.cmdPowrot)
        Me.Controls.Add(Me.Panel1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "_frmPolaczenieSieciowe"
        Me.Text = "Po³¹czenie sieciowe"
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub _frmPolaczenieSieciowe_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.MrJones
        Me.picLogo.Image = My.Resources.logomini
        Me.cmdPowrot.Image = My.Resources._EXIT.ToBitmap

        lblTenKomputer.Text = ""
        lblSposobPolaczenia.Text = ""
        lblAdresIP.Text = ""
        lblNazwaHosta.Text = ""
        lblAdresIPv6.Text = ""

        '-------------------------------
        ' Get the connected status.
        Dim dwFlags As Int32
        Dim Connected As Boolean = CBool(InternetGetConnectedState(dwFlags, 0&))
        If Connected Then
            lblTenKomputer.Text = "Po³¹czony z Internetem"

            ' Display all connection flags.
            Dim ConnectionType As ConnectStates
            For Each ConnectionType In System.Enum.GetValues(GetType(ConnectStates))
                If (ConnectionType And dwFlags) = ConnectionType Then
                    lblSposobPolaczenia.Text = ConnectionType.ToString()
                    Exit For
                End If
            Next

            'Get Host Name
            Dim HostName As String = System.Net.Dns.GetHostName()
            Dim IPAddress As String = ""
            lblNazwaHosta.Text = HostName

            'Get IP v6 Address
            'IPAddress = System.Net.Dns.GetHostEntry(HostName).AddressList(0).ToString()
            IPAddress = ""
            Dim ii As Integer = 0
            For ii = 0 To System.Net.Dns.GetHostEntry(HostName).AddressList.Length - 1
                If System.Net.Dns.GetHostEntry(HostName).AddressList(ii).IsIPv6LinkLocal Or _
                                    System.Net.Dns.GetHostEntry(HostName).AddressList(ii).IsIPv6Multicast Or _
                                    System.Net.Dns.GetHostEntry(HostName).AddressList(ii).IsIPv6SiteLocal Or _
                                    System.Net.Dns.GetHostEntry(HostName).AddressList(ii).IsIPv6Teredo Then
                    IPAddress &= System.Net.Dns.GetHostEntry(HostName).AddressList(ii).ToString() & vbCrLf
                End If
            Next
            lblAdresIPv6.Text = IPAddress

            'Get IP v4 Address
            IPAddress = ""
            For ii = 0 To System.Net.Dns.GetHostEntry(HostName).AddressList.Length - 1
                If System.Net.Dns.GetHostEntry(HostName).AddressList(ii).IsIPv6LinkLocal Or _
                                    System.Net.Dns.GetHostEntry(HostName).AddressList(ii).IsIPv6Multicast Or _
                                    System.Net.Dns.GetHostEntry(HostName).AddressList(ii).IsIPv6SiteLocal Or _
                                    System.Net.Dns.GetHostEntry(HostName).AddressList(ii).IsIPv6Teredo Then
                Else
                    IPAddress &= System.Net.Dns.GetHostEntry(HostName).AddressList(ii).ToString() & vbCrLf
                End If
            Next
            lblAdresIP.Text = IPAddress


        Else
            lblTenKomputer.Text = "Nie jest po³¹czony z Internetem"
        End If
    End Sub

    Private Sub cmdPowrot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowrot.Click
        Me.Close()
    End Sub

    Private Sub _frmPolaczenieSieciowe_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        RysujGradient(Me)
    End Sub
End Class
