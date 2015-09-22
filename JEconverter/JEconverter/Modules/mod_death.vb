Imports System
Imports System.Security
Imports System.Security.Cryptography
Imports System.Text

Module mod_death

    ''' <summary>
    ''' Generate HASH Value
    ''' </summary>
    ''' <param name="str">String you want to be hash</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Function hashMe(ByVal str As String) As String
        Dim tmpSrc() As Byte
        Dim tmpHash() As Byte

        tmpSrc = ASCIIEncoding.ASCII.GetBytes(str)
        tmpHash = New MD5CryptoServiceProvider().ComputeHash(tmpSrc)

        Return ByteArrayToString(tmpHash)
    End Function

    Private Function ByteArrayToString(ByVal arrInput() As Byte) As String
        Dim i As Integer
        Dim sOutput As New StringBuilder(arrInput.Length)
        For i = 0 To arrInput.Length - 1
            sOutput.Append(arrInput(i).ToString("X2"))
        Next
        Return sOutput.ToString()
    End Function
End Module
