Imports System
Imports System.IO
Imports System.IO.Compression
Imports Ionic.Zip

Public Class ZipHelper
    ''' <summary>
    ''' Creates an encrypted ZIP file containing all files from the specified directory
    ''' </summary>
    ''' <param name="sourceDirectory">Directory containing files to be zipped</param>
    ''' <param name="zipFilePath">Path where the ZIP file will be created</param>
    ''' <param name="password">Password for encryption</param>
    Public Shared Sub CreateEncryptedZip(ByVal sourceDirectory As String, ByVal zipFilePath As String, ByVal password As String)
        Using zip As New ZipFile()
            ' Set encryption
            zip.Password = password
            zip.Encryption = EncryptionAlgorithm.WinZipAes256
            
            ' Add all files from the directory
            Dim files As String() = Directory.GetFiles(sourceDirectory)
            For Each file As String In files
                zip.AddFile(file, "")
            Next
            
            ' Add subdirectories if any
            Dim dirs As String() = Directory.GetDirectories(sourceDirectory)
            For Each dir As String In dirs
                Dim dirName As String = Path.GetFileName(dir)
                zip.AddDirectory(dir, dirName)
            Next
            
            ' Save the ZIP file
            zip.Save(zipFilePath)
        End Using
    End Sub
End Class