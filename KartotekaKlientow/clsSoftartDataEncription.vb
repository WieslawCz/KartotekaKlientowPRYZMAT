Option Strict Off
Option Explicit On

Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

Namespace Softart
    Public Class MyDataEncription

        Private _Key As Byte()
        Private _Vector As Byte()
        Private _EncryptorTransform As ICryptoTransform
        Private _DecryptorTransform As ICryptoTransform
        Private _UTFEncoder As UTF8Encoding


        '********************************
        '** Initializes the simple AES class.
        '** <param name="Key">The encryption key.</param>
        '** <param name="Vector">The IV.</param>
        '********************************
        Public Sub SimpleAES(ByVal key As Byte(), ByVal Vector As Byte())
            _Key = key
            _Vector = Vector

            'This is our encryption method
            Dim rm As RijndaelManaged = New RijndaelManaged

            'Create an encryptor and a decryptor using our 
            'encryption method, key, and vector.
            _EncryptorTransform = rm.CreateEncryptor(_Key, _Vector)
            _DecryptorTransform = rm.CreateDecryptor(_Key, _Vector)

            'Used to translate bytes to text and vice versa
            _UTFEncoder = New System.Text.UTF8Encoding()
        End Sub


        '*********************************
        '** Generates an encryption key.
        '** <returns>byte[] encryption key. Store it some place safe.</returns>
        '*********************************
        Public Function GenerateEncryptionKey() As Byte()
            'Generate a Key.
            Dim rm As RijndaelManaged = New RijndaelManaged
            rm.GenerateKey()
            Return rm.Key
        End Function


        '*********************************
        '** Generates a unique encryption vector
        '** <returns>byte[] vector value</returns>
        '*********************************
        Public Function GenerateEncryptionVector() As Byte()
            'Generate a Vector
            Dim rm As RijndaelManaged = New RijndaelManaged
            rm.GenerateIV()
            Return rm.IV
        End Function


        '*********************************
        '** Encrypt some text and return an encrypted byte array.
        '** <param name="TextValue">The Text to encrypt</param>
        '** <returns>byte[] of the encrypted value</returns>
        '*********************************
        Public Function Encrypt(ByVal TextValue As String) As Byte()
            'Translates our text value into a byte array.
            Dim bytes As Byte() = _UTFEncoder.GetBytes(TextValue)

            'Used to stream the data in and out of the CryptoStream.
            Dim MemStream As IO.MemoryStream = New MemoryStream()

            ' We will have to write the unencrypted bytes to the 
            ' stream, then read the encrypted result back from the stream.
            Dim cs As CryptoStream = New CryptoStream(MemStream, _
                                                      _EncryptorTransform, _
                                                      CryptoStreamMode.Write)
            cs.Write(bytes, 0, bytes.Length)
            cs.FlushFinalBlock()

            'Read encrypted value back out of the stream
            MemStream.Position = 0
            Dim encryptedBytes(MemStream.Length) As Byte
            MemStream.Read(encryptedBytes, 0, encryptedBytes.Length)

            'Clean up.
            cs.Close()
            MemStream.Close()

            Return encryptedBytes
        End Function


        '*********************************
        '** Decrypt some text and return an decrypted byte array.
        '** <param name="TextValue">The Text to encrypt</param>
        '** <returns>byte[] of the encrypted value</returns>
        '*********************************
        Public Function Decrypt(ByVal EncryptedValue As Byte()) As String
            'Write the encrypted value to the decryption stream
            Dim encryptedStream As MemoryStream = New MemoryStream()
            Dim decryptedStream As CryptoStream = New CryptoStream(encryptedStream, _
                                                                   _DecryptorTransform, _
                                                                   CryptoStreamMode.Write)
            decryptedStream.Write(EncryptedValue, 0, EncryptedValue.Length)
            decryptedStream.FlushFinalBlock()

            'Read the decrypted value from the stream.
            encryptedStream.Position = 0
            Dim decryptedBytes(encryptedStream.Length) As Byte

            encryptedStream.Read(decryptedBytes, 0, decryptedBytes.Length)
            encryptedStream.Close()

            Return _UTFEncoder.GetString(decryptedBytes)
        End Function
    End Class


End Namespace
