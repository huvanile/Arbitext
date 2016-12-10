Imports ArbitextClassLibrary.Globals
Imports ArbitextClassLibrary.SFTPHelpers

Module Main

    Public Sub Main(ByVal args() As String)
        City = args(0)
        TldUrl = "http://" & Trim(City) & ".craigslist.com"
        SftpURL = args(1)
        SftpUser = args(2)
        SftpPass = args(3)
        SftpDirectory = args(4)
        If ConnectToSFTP() Then
            SearchHelpers.allQuerySearch()
        End If
        Sftp.Disconnect()
        Sftp.Dispose()
    End Sub

End Module