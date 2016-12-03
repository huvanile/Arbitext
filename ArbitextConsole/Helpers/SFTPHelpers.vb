Imports ArbitextConsole.Globals
Imports Renci.SshNet

Public Class SFTPHelpers
    Public Shared Function ConnectToSFTP()
        Console.WriteLine("Connecting to " & SftpURL)
        Sftp = New SftpClient(SftpURL, SftpUser, SftpPass)
        Try
            Sftp.Connect()
            Sftp.ChangeDirectory(SftpDirectory)
            If Sftp.IsConnected Then
                Console.WriteLine("Connected!")
                Return True
            Else
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("SFTP Connection Error!")
                Console.ResetColor()
                Return False
            End If
        Catch ex As Exception
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("SFTP Connection Error!")
            Console.ResetColor()
            Return False
        End Try
    End Function
End Class
