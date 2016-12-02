Imports Renci.SshNet
Imports System.IO

Module Module1
    Private _city As String
    Private _sftpURL As String
    Private _sftpUser As String
    Private _sftpPass As String
    Private _sftpDirectory As String
    Private _saveAsFolder As String

    Sub Main(ByVal args() As String)
        _city = args(0)
        _sftpURL = args(1)
        _sftpUser = args(2)
        _sftpPass = args(3)
        _sftpDirectory = args(4)
        _saveAsFolder = args(5)
        grabOldXMLFiles()
    End Sub

    Sub grabOldXMLFiles()
        Console.WriteLine("Connecting to " & _sftpURL)
        Using sftp As New SftpClient(_sftpURL, _sftpUser, _sftpPass)
            sftp.Connect()
            sftp.ChangeDirectory(_sftpDirectory)
            If sftp.IsConnected Then
                Console.WriteLine("Connected!")
                'For Each file As String In Directory.GetFiles(_saveAsFolder)
                '    If file Like "*.xml" Or file Like "*.php*" Then
                '        Using filestream As New FileStream(file, FileMode.Open)
                '            sftp.BufferSize = 4 * 1024
                '            sftp.UploadFile(filestream, Path.GetFileName(file))
                '        End Using
                '    End If
                'Next
            Else
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("SFTP Connection Error!")
                Console.ResetColor()
            End If
        End Using
        Console.Write("Press Enter to exit")
        Console.Read()
    End Sub
End Module
