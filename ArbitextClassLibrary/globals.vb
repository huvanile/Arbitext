Imports Renci.SshNet

Public Class Globals
    Public Shared City As String
    Public Shared SftpURL As String
    Public Shared SftpUser As String
    Public Shared SftpPass As String
    Public Shared SftpDirectory As String
    Public Shared TldUrl As String             'top-level domain name (e.g., http://dallas.craigslist.com)
    Public Shared Sftp As SftpClient
    Public Const wwwRoot As String = "http://huvanile.com/"
    Public Const _resultHook As String = "<li class=""result-row"" data-pid="
End Class
