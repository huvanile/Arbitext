Imports ArbitextClassLibrary.SFTPHelpers
Imports ArbitextClassLibrary.Globals
Imports ArbitextClassLibrary.RSSHelpers

Module Main

    Public Sub Main(ByVal args() As String)
        Try
            City = args(0)
            TldUrl = "http://" & Trim(City) & ".craigslist.com"
            SftpURL = args(1)
            SftpUser = args(2)
            SftpPass = args(3)
            SftpDirectory = args(4)
            If ConnectToSFTP() Then
                Console.ForegroundColor = ConsoleColor.DarkCyan
                Console.WriteLine("Checking all feeds for dead leads, removing if needed")
                Console.ResetColor()
                RemoveDeadLeads(Sftp)
                SearchHelpers.allQuerySearch()
            End If
            Sftp.Disconnect()
            Sftp.Dispose()
        Catch
            MsgBox("Command line arguements are required in order to run this console app", vbCritical, ArbitextClassLibrary.Globals.Title)
        End Try
    End Sub

End Module
