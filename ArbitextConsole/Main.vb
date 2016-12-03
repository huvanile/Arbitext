Imports ArbitextConsole.SFTPHelpers
Imports ArbitextConsole.Globals
Imports ArbitextConsole.SearchHelpers

Module Main

    Sub Main(ByVal args() As String)
        City = args(0)
        TldUrl = "http://" & Trim(City) & ".craigslist.com"
        SftpURL = args(1)
        SftpUser = args(2)
        SftpPass = args(3)
        SftpDirectory = args(4)
        If connectToSFTP() Then
            allQuerySearch()
        End If
        Sftp.Disconnect()
        Sftp.Dispose()
    End Sub

End Module
