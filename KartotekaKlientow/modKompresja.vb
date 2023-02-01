Imports System
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Imports ICSharpCode.SharpZipLib.Checksums
Imports ICSharpCode.SharpZipLib.Zip
'Imports ICSharpCode.SharpZipLib.GZip

Module modKompresja
    '=========================================================================
    'UWAGA: nie potrafie archiwizowac kilku plikow do jednego ZIPa
    '       Dla tej wersji programu trzeba ograniczyæ siê do jednego pliku w jednym ZIPie !!!!!
    '=========================================================================

    Public Sub ZipFile(ByVal FileListToZip As String, _
                       ByVal ZippedFile As String)

        Dim nCompressionLevel As Integer = 5    '0..9
        Dim nBlockSize As Integer = 2048
        Dim nBuffer(nBlockSize) As Byte
        Dim nSize As System.Int32 = 0
        Dim nSizeRead As Integer = 0
        Dim FileToZip As String = ""

        Dim SepPosition As Integer = 0

        'otwórz strumieñ pliku wyjœciowego
        Dim strmZipFile As System.IO.FileStream
        If System.IO.File.Exists(ZippedFile) Then
            'jesli wyjsciowy ZIP juz istnieje - otwórz go
            strmZipFile = System.IO.File.OpenWrite(ZippedFile)
        Else
            'jestnie istnieje - stworz
            strmZipFile = System.IO.File.Create(ZippedFile)
        End If
        Dim strmZipStream As New ZipOutputStream(strmZipFile)
        strmZipStream.SetLevel(nCompressionLevel)  'Compression Level: 0-9

        Do While Len(FileListToZip) > 0
            SepPosition = InStr(FileListToZip, "|")
            If SepPosition = 0 Then
                FileToZip = FileListToZip
                FileListToZip = ""
            Else
                FileToZip = Mid(FileListToZip, 1, SepPosition - 1)
                FileListToZip = Mid(FileListToZip, SepPosition + 1)
            End If

            'sprawdz czy taki pliK istnieje...
            If (Not System.IO.File.Exists(FileToZip)) Then
                'nie znalaz³em pliku do kompresji...
                'Throw New System.IO.FileNotFoundException("Plik " + FileToZip + " nie zosta³ znaleziony." & vbCrLf & _
                '                                          "Nie mogê wykonaæ kompresji.")
            Else
                'otwórz strumien z pliku wejsciowwego
                Dim strmStreamToZip As New System.IO.FileStream(FileToZip, _
                                           System.IO.FileMode.Open, _
                                           System.IO.FileAccess.Read)

                'Skoro istnieje plik wejsciowy, to istnieje nazwa dopisywanego pliku
                Dim ZippedEntry As String = Path.GetFileName(FileToZip)
                Dim myZipEntry As New ZipEntry(ZippedEntry)
                myZipEntry.DateTime = DateTime.Now
                myZipEntry.Size = strmStreamToZip.Length
                'bedziemy pisac do tego elementu
                strmZipStream.PutNextEntry(myZipEntry)

                Try
                    nSize = 0
                    While (nSize < strmStreamToZip.Length)
                        'czytaj porcje ze strumienia wejsciowego
                        nSizeRead = strmStreamToZip.Read(nBuffer, 0, nBuffer.Length)
                        'zapisz porcje w strumieniu wyjsciowym
                        strmZipStream.Write(nBuffer, 0, nSizeRead)
                        'tyle juz przeniesiono do ZIP
                        nSize = nSize + nSizeRead
                    End While

                Catch Ex As System.Exception
                    Throw Ex
                End Try

                strmStreamToZip.Close()
            End If
        Loop
        strmZipStream.Flush()
        strmZipStream.Finish()
        strmZipStream.Close()
    End Sub


    Public Sub UnzipFile(ByVal FileListToUnZip As String, _
                         ByVal ZippedFile As String)

        Dim SepPosition As Integer = 0
        Dim FileToUnZip As String = ""
        Dim FileToUnZipName As String = ""
        Dim FileToUnZipPath As String = ""

        'strumien odczytu z Zipa
        Dim strm As New ZipInputStream(File.OpenRead(ZippedFile))
        'pozycja w ZIPie
        Dim myEntry As ZipEntry
        Dim directoryName As String = ""
        Dim fileName As String = ""
        Dim FileList As String = ""

        'analizuj wszystkie elementy ZIPa
        myEntry = strm.GetNextEntry()
        Do Until myEntry Is Nothing
            directoryName = Path.GetDirectoryName(myEntry.Name)
            fileName = Path.GetFileName(myEntry.Name)
            'zbior w ZIPie nie jest pusty...
            If (fileName <> String.Empty) Then
                'sprawdz czy ten plik jest na liscie UNZIP FILE
                FileList = FileListToUnZip
                Do While Len(FileList) > 0
                    SepPosition = InStr(FileList, "|")
                    If SepPosition = 0 Then
                        FileToUnZip = FileList
                        FileList = ""
                    Else
                        FileToUnZip = Mid(FileList, 1, SepPosition - 1)
                        FileList = Mid(FileList, SepPosition + 1)
                    End If
                    FileToUnZipName = Path.GetFileName(FileToUnZip)
                    FileToUnZipPath = Path.GetDirectoryName(FileToUnZip)

                    If fileName = FileToUnZipName Then
                        'odtwarzamy ten plik z Zipa

                        '// create directory
                        Directory.CreateDirectory(FileToUnZipPath)
                        Dim StreamWriter As FileStream = File.Create(FileToUnZip)
                        Dim nSize As Integer = 2048
                        Dim nBuffer(nSize) As Byte
                        While (True)
                            nSize = strm.Read(nBuffer, 0, nBuffer.Length)
                            If (nSize > 0) Then
                                StreamWriter.Write(nBuffer, 0, nSize)
                            Else
                                Exit While
                            End If
                        End While
                        StreamWriter.Close()
                        Exit Do
                    End If
                Loop
            End If
            'nastêpna pozycjaw ZIPie
            myEntry = strm.GetNextEntry()
        Loop
        strm.Close()
    End Sub


    Public Sub ViewZipFile(ByVal ZippedFile As String, _
                            ByVal ListV As ListView)
        Dim strmZipInputStream As ZipInputStream = New ZipInputStream(File.OpenRead(ZippedFile))
        Dim objEntry As ZipEntry
        Dim i As Integer = 0

        Dim MaxSize As Integer = 2048
        Dim nSize As Integer = 0
        Dim abyData(MaxSize + 2) As Byte

        objEntry = strmZipInputStream.GetNextEntry()
        While IsNothing(objEntry) = False
            ListV.Items.Add(objEntry.Name.ToString, i)
            ListV.Items(i).SubItems.Add(objEntry.DateTime.ToString)
            ListV.Items(i).SubItems.Add(objEntry.Size.ToString)
            ListV.Items(i).SubItems.Add(objEntry.CompressedSize.ToString)
            i += 1

            If objEntry.IsFile Then
                'przewin ten zbior...
                While True
                    nSize = strmZipInputStream.Read(abyData, 0, MaxSize)
                    If nSize > 0 Then
                        'strBuilder.Append(New ASCIIEncoding().GetString(abyData, 0, nSize) + vbCrLf)
                    Else
                        Exit While
                    End If
                End While
            End If
            objEntry = strmZipInputStream.GetNextEntry()
        End While
        strmZipInputStream.Close()
    End Sub

End Module
