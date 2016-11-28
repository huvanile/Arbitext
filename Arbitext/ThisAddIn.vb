Imports System.Threading

Public Class ThisAddIn
    Public Const Title As String = "Arbitext"
    Public Const ResultHook As String = "<li class=""result-row"" data-pid="
    Public Shared AppExcel As Excel.Application
    Public Shared Proceed As Boolean
    Public Shared AppIE As Object
    Public Shared frmPrefs As FrmPrefs
    Public Shared TldUrl As String             'top-level domain name (e.g., http://dallas.craigslist.com)
    Public Shared t1 As Thread
    Public Const wwwRoot As String = "http://huvanile.com/"
    Public Shared SaveAsFolder As String

    'sftp
    Public Shared SFTPUrl As String
    Public Shared SFTPUser As String
    Public Shared SFTPPass As String
    Public Shared SFTPDirectory As String

    'pushbullet
    Public Shared PbAPIKey As String
    Public Shared PbDeviceID As String

    'email
    Public Shared EmailsOK As Boolean
    Public Shared EmailAddress As String
    Public Shared EmailPassword As String

    'preferences
    Public Shared NotifyViaPBOK As Boolean
    Public Shared KeepIEVisibleOK As Boolean
    Public Shared MinTolerableProfit As Decimal
    Public Shared SaveWBFilePath As String                 'saveFileAs: output filepath
    Public Shared SaveWBFileName As String
    Public Shared OnWinnersOK As Boolean
    Public Shared OnMaybesOK As Boolean
    Public Shared City As String

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        AppExcel = Me.Application
        RegistryHelpers.addressNullPrefs()
        ThisAddIn.Proceed = True
        RegistryHelpers.loadVariablesFromRegistry()
    End Sub

End Class
