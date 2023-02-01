Module modThunderBird


    '------------------------------------------
    'MOZILLA THUNDERBIRD
    '------------------------------------------

    'https://developer.mozilla.org/en-US/docs/Mozilla/Command_Line_Options#-compose_message_options
    '------------------------------------------
    'Command line options are used to specify various startup options for Mozilla applications. 
    'For example, you can use command line configuration options to bypass the Profile Manager 
    'and open a specific profile (if you have multiple profiles). You can also control how Mozilla 
    'applications open, which components open initially, and what the components do when they open. '
    'This page describes the commonly used options and how to use them.
    '
    'Syntax Rules
    'But first, let's describe the syntax rules that apply for all options.
    '   •Command parameters containing spaces must be enclosed in quotes; for example, "Joel User".
    '   •Command actions are not case sensitive.
    '   •Command parameters except profile names are not case sensitive.
    '   •Blank spaces ( ) separate commands and parameters.
    '   •Each message option follows the syntax field=value, for example: 
    '       ◦to=foo@nowhere.net
    '       ◦subject=cool page
    '       ◦attachment=www.mozilla.org
    '       ◦attachment='file:///c:/test.txt'
    '       ◦body=check this page
    '   •Multiple message options are separated by comma (,), for example: 
    '       "to=foo@nowhere.net,subject=cool page" . 
    '
    'Comma separators must not follow or precede spaces ( ). To assign multiple values to a field, 
    'enclose the values in single quotes ('), for example: 
    '       "to='foo@nowhere.net,foo@foo.de',subject=cool page" .
    '
    'Using command line options
    'Command line options are entered after the command to start the application. 
    'Some options have arguments. These are entered after the command line option. 
    'Some options have abbreviations. For example, the command line option "-editor" can be abbreviated as "-edit". 
    '(Where abbreviations are available, they are described in the text below.) 
    'In some cases option arguments must be enclosed in quotation marks. 
    '(This is noted in the option descriptions below.) Multiple command line options can be specified. 
    'In general, the syntax is as follows:
    '   application -option -option "argument" -option argument1
    '------------------------------------------
    '-mail
    'Start with the mail client. Thunderbird and SeaMonkey only.
    '-news news_URL
    'Start with the news client. If news_URL (optional) is given, open the specified newsgroup. Thunderbird and SeaMonkey only.
    '   thunderbird -news news://server/group
    '-compose message_options
    'Start with mail composer. See syntax rules. Thunderbird and SeaMonkey only.
    '   thunderbird -compose "to=foo@nowhere.net"
    '-addressbook
    'Start with address book. Thunderbird and SeaMonkey only.
    '-options
    'Open Options/Preferences window. Thunderbird only.
    '-offline
    'Start with the offline mode. Thunderbird and SeaMonkey only.
    '-setDefaultMail
    'Set the application as the default email client. Thunderbird only.
    '------------------------------------------
    'Thunderbird supports the following command line arguments:
    'All of the Mozilla command line arguments that aren't browser specific. However, you need to figure out what version of Gecko your version of Thunderbird uses to see if some of those command line arguments aren't available. Thunderbird 3.1 is based on Gecko 1.9.2. The -install-global-extension and -install-global-theme arguments have been removed from Gecko 1.9.2 and later versions.
    'Notice the syntax section at the bottom of that writeup. You can use -compose message_options to have it bring up the compose message window and fill in everything for you, but you still need to press the Send button to actually send the message.
    'Command line arguments for the extension manager. It's used for a number of troubleshooting or system administration tasks such as running in Safe Mode.
    'The -profile "path" command line argument to specify the location of the profile. It's used to run Thunderbird with the specified profile regardless of whether the Profile Manager knows about that profile's existence. It's described in more detail in the writeup on USB drive support but it does not require a USB drive. Its useful if you're a roaming user or Thunderbird somehow lost track of your profile (perhaps due to your system crashing) and you want to verify the profile is still good before trying to fix the problem. Example: "C:\Program Files\Mozilla Thunderbird\thunderbird.exe" -profile e:\my_profile will launch Thunderbird with the profile stored at e:\my_profile. A quick sanity check is that e:\my_profile should contain your prefs.js file. Note however that you normally cannot start Thunderbird with a second profile if Thunderbird is already running.
    'The -migration command line argument opens import wizard which help with import data from Outlook (Express) and Mozilla Suite/SeaMonkey.
    'The path of a .eml file. This will launch Thunderbird with a window displaying the contents of the .eml file. For example, "C:\Program Files\Thunderbird\thunderbird.exe" c:\test.eml
    'The -mail URL command line argument opens the message whose URL you specified. Select the message in Thunderbird, enter var hdr = top.opener.gFolderDisplay.selectedMessage; alert(hdr.folder.getUriForMsg(hdr)); in the error console and press the evaluate button to get a popup with the URL. You can't specify the URL for a folder, only for a specific message. [1] For example: thunderbird.exe -mail "mailbox-message://user%40gmail.com@pop.googlemail.com/Templates#12345"
    '[edit]
    'Compose new mail with command line

    'You have to use the command line "-compose" to launch Thunderbird and open a new compose window. The following arguments are available:
    '"to" : used to specify the email of the recipient
    '"cc" : used to specify the email of the recipient of a copy of the mail
    '"bcc" : used to specify the email of the recipient of a blind copy of the mail*
    '"newsgroups" : one or more news groups to submit the message to*
    '"subject" : subject of the mail
    '"format" : compose the message in HTML ("format=1") or plain text ("format=2")*
    '"preselectid" : an identifier for the "From" identity to choose from the menu*
    'note that you cannot directly specify an e-mail address but need to find the identity key
    'for example, "preselectid=id2" would select the identity #2
    'you can find the key by try-and-error or by searching for useremail in the Config Editor
    '"body" : body of the mail
    '"attachment' : specify the directory and the name of an attachment
    'the value should be a file:// url, properly encoded
    'with tb3+, you can alternatively use the absolute file name (unencoded)
    '*(in Thunderbird 3.1.9 and earlier, these options must be preceded by at least one other option and cannot be at the start of the argument list due to bug 627999.)
    'Watch out for the somewhat complex syntax of the "-compose" command-line option. The double-quotes enclose the full comma-separated list of arguments passed to "-compose", whereas single quotes are used to group items for the same argument. Example:
    'thunderbird -compose "to='john@example.com,kathy@example.com',cc='britney@example.com',subject='dinner',body='How about dinner tonight?',attachment='C:\temp\info.doc,C:\temp\food.doc'"
    '(use attachment="file:///C:/temp/food.doc" for Thunderbird 2.0)
    'For mailto: urls the "in-reply-to" header is also supported
    '"in-reply-to" : adds In-Reply-To with the provided reference, adds it to References
    '==========================================

    '   <path>thunderbird -compose "to=foo@nowhere.net,subject=<tytuł>,body=<tresc>,attachment='file:///c:/test.txt'"
    '   <path>=C:\Program Files (x86)\Mozilla Thunderbird\thunderbird.exe

    'Aby w treści umieścić informacje wielowierszową, należy wszystkie niewidoszne znaki i te które są zabronione w URL zamienić na sekwencje :
    'spacja - %20, CRLF na %0d%0a, ! na %21, 



    '****************************************************
    ' moduł integracji z programem pocztowym Mozilla Thunderbird
    '****************************************************

    Public Sub SendByThunderbird(ByVal pAdresOdbiorcy As String, _
                                 ByVal pCC As String, _
                                 ByVal pBCC As String, _
                                 ByVal pTemat As String, _
                                 ByVal pTresc As String, _
                                 ByVal pZalacznik As String)
        Dim rThunderbirdExe As String = "Thunderbird.exe"
        Dim rThunderbirdFullPath As String = ""
        Dim rKomenda As String = ""
        Dim i As Integer = 0
        Dim ch As String = ""
        Dim NewTresc As String = ""

        If Len(ParL_eMailThunderbirdPath) = 0 Then
            MessageBox.Show("Katalog z programem Mozilla Thunderbird nie istnieje" & vbCrLf & ParL_eMailThunderbirdPath, _
                            "Nie mogę wysłać wiadomości.", _
                System.Windows.Forms.MessageBoxButtons.OK, _
                MessageBoxIcon.Information, _
                MessageBoxDefaultButton.Button1)
            Return
        Else
            If Mid(ParL_eMailThunderbirdPath, Len(ParL_eMailThunderbirdPath), 1) = "\" Then
                rThunderbirdFullPath = ParL_eMailThunderbirdPath & rThunderbirdExe
            Else
                rThunderbirdFullPath = ParL_eMailThunderbirdPath & "\" & rThunderbirdExe
            End If

            If Not IO.File.Exists(rThunderbirdFullPath) Then
                MessageBox.Show("Wybrany katalog z programem Mozilla Thunderbird " & vbCrLf & _
                                ParL_eMailThunderbirdPath & vbCrLf & _
                                "nie zawiera programu Thunderbird.exe", _
                                "Nie mogę wysłać wiadomości.", _
                    System.Windows.Forms.MessageBoxButtons.OK, _
                    MessageBoxIcon.Information, _
                    MessageBoxDefaultButton.Button1)
                Return
            Else
                If Len(pAdresOdbiorcy) = 0 Then
                    MessageBox.Show("Klient nie ma zdefiniowanego adresu eMail" & vbCrLf & _
                                    "w Katalogu Klientów programu.", _
                                    "Nie mogę wysłać wiadomości.", _
                        System.Windows.Forms.MessageBoxButtons.OK, _
                        MessageBoxIcon.Information, _
                        MessageBoxDefaultButton.Button1)
                    Return
                Else
                    'jest gotowy program Mozilla Thunderbird
                    rKomenda = rThunderbirdExe & " -compose to=" & Trim(pAdresOdbiorcy)
                    If Len(pCC) > 0 Then rKomenda &= ",cc=" & Trim(pCC)
                    If Len(pBCC) > 0 Then rKomenda &= ",bcc=" & Trim(pCC)
                    If Len(pTemat) > 0 Then rKomenda &= ",subject='" & Trim(pTemat) & "'"

                    If Len(pTresc) > 0 Then
                        'zamieniamy znaki  w pTresc
                        NewTresc = ""
                        For i = 1 To Len(pTresc)
                            ch = Mid(pTresc, i, 1)
                            Select Case Asc(ch)
                                Case &H9    'tab
                                    NewTresc &= "%09"
                                Case &HA    'LF
                                    NewTresc &= "%0a"
                                Case &HD    'CR
                                    NewTresc &= "%0d"
                                Case &H20   'SPace
                                    NewTresc &= "%20"
                                Case &H21   '!
                                    NewTresc &= "%21"

                                Case Else
                                    NewTresc &= ch
                            End Select
                        Next
                        rKomenda &= ",body=" & Trim(NewTresc)
                    End If

                    If Len(pZalacznik) > 0 Then
                        rKomenda &= ",attachment='"
                        'opracowujemy załączniki - muszą być jednoznacznie zdefiniowane
                        Dim AppPath As String = IO.Path.GetDirectoryName(pZalacznik)
                        Dim AppCount As Integer = 0
                        Dim AppName As String = Dir(pZalacznik)
                        While Len(AppName) > 0
                            AppCount += 1
                            If AppCount > 1 Then
                                rKomenda &= "," & AppPath & "\" & AppName
                            Else
                                rKomenda &= AppPath & "\" & AppName
                            End If
                            AppName = Dir()
                        End While
                        rKomenda &= "'"
                    End If

                    '-----------------------------
                    Dim aktDir As String = IO.Directory.GetCurrentDirectory
                    IO.Directory.SetCurrentDirectory(ParL_eMailThunderbirdPath)
                    '==============================================
                    Call Shell(rKomenda, AppWinStyle.NormalNoFocus, True)
                    '-------------------------------------
                    'Process.Start("cmd.exe", " /c " & rKomenda)
                    '==============================================
                    IO.Directory.SetCurrentDirectory(aktDir)

                End If
            End If
        End If
    End Sub


End Module
