Imports System
Imports System.IO
Imports System.Security
Imports System.Security.Cryptography
Imports System.Runtime.InteropServices
Imports System.Text

Public Class Main
    Public Property TxtKey As Object

    Public Sub ZeroMemory(ByVal addr As IntPtr, ByVal size As Integer)
    End Sub

    Function GenerateKey() As String
        Dim desCrypto As DESCryptoServiceProvider = DESCryptoServiceProvider.Create()
        Return ASCIIEncoding.ASCII.GetString(desCrypto.Key)
    End Function

    Sub EncryptFile(ByVal sInputFileName As String,
                    ByVal sOuputFileName As String,
                    ByVal sKey As String)
        Dim fsInput As New FileStream(sInputFileName,
                           FileMode.Open, FileAccess.Read)
        Dim fsEncrypted As New FileStream(sOuputFileName,
                                          FileMode.Create, FileAccess.Write)

        Dim DES As New DESCryptoServiceProvider()

        DES.Key = ASCIIEncoding.ASCII.GetBytes(sKey)
        DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey)

        Dim desencrypt As ICryptoTransform = DES.CreateEncryptor()

        Dim cryptostream As New CryptoStream(fsEncrypted,
                                           desencrypt,
                                           CryptoStreamMode.Write)

        Dim bytearrayinput(fsInput.Length - 1) As Byte
        fsInput.Read(bytearrayinput, 0, bytearrayinput.Length)
        cryptostream.Write(bytearrayinput, 0, bytearrayinput.Length)
        cryptostream.Close()
    End Sub

    Sub DecryptFile(ByVal sInputFilename As String,
        ByVal sOutputFilename As String,
        ByVal sKey As String)

        Dim DES As New DESCryptoServiceProvider()

        DES.Key() = ASCIIEncoding.ASCII.GetBytes(sKey)

        DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey)

        Dim fsread As New FileStream(sInputFilename, FileMode.Open, FileAccess.Read)

        Dim desdecrypt As ICryptoTransform = DES.CreateDecryptor()
        Dim cryptostreamDecr As New CryptoStream(fsread, desdecrypt, CryptoStreamMode.Read)
        Dim fsDecrypted As New StreamWriter(sOutputFilename)
        fsDecrypted.Write(New StreamReader(cryptostreamDecr).ReadToEnd)
        fsDecrypted.Flush()
        fsDecrypted.Close()
    End Sub

    Private Sub ButBrowseEnc_Click(sender As Object, e As EventArgs) Handles ButBrowseEnc.Click
        If OpenFileDialog.ShowDialog() = DialogResult.OK Then
            TxtEncrypt.Text = OpenFileDialog.FileName.ToString
        End If
    End Sub
    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles Encrypt.Click

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog.FileOk

    End Sub


    Private Sub ButEncrypt_Click(sender As Object, e As EventArgs) Handles ButEncrypt.Click

        If Not TxtEncrypt.Text = "" Then
            Dim sSecretKey As String
            sSecretKey = GenerateKey()
            Dim gch As GCHandle = GCHandle.Alloc(sSecretKey, GCHandleType.Pinned)

            EncryptFile(TxtEncrypt.Text,
                        TxtEncrypt.Text + "_encrypted",
                        sSecretKey)

            Key.Show()
            Key.TxtCode.Text = sSecretKey


            ZeroMemory(gch.AddrOfPinnedObject(), sSecretKey.Length * 2)
            gch.Free()
        End If
    End Sub

    Private Sub ButDecrypt_Click(sender As Object, e As EventArgs) Handles ButDecrypt.Click
        If Not TxtDecrypt.Text = "" Then
            If Not TxtDec.Text = "" Then
                DecryptFile(TxtDecrypt.Text,
                            TxtDecrypt.Text + "_decrypted",
                            TxtKey.Text)
                MessageBox.Show("Completed!", "Decryption", MessageBoxButtons.OK)

            Else
                TxtDec.Text = "Please insert your Key!"
            End If
        Else
            TxtDecrypt.Text = "Please select your file!"
        End If
    End Sub

    Private Sub ButBrowseD_Click(sender As Object, e As EventArgs) Handles ButBrowseDec.Click
        If OpenFileDialog.ShowDialog() = DialogResult.OK Then
            TxtDecrypt.Text = OpenFileDialog.FileName.ToString
        End If
    End Sub
End Class
