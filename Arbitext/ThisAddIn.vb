Public Class ThisAddIn
    Public Const Title As String = "Arbitext"
    Public Const ResultHook As String = "<li class=""result-row"" data-pid="
    Public Shared AppExcel As Excel.Application
    Public Shared Proceed As Boolean
    Public Shared AppIE As Object
    Public Shared frmPrefs As FrmPrefs
    Public Shared TldUrl As String             'top-level domain name (e.g., http://dallas.craigslist.com)

    'pushbullet
    Public Shared PbAPIKey As String
    Public Shared PbDeviceID As String

    'email
    Public Shared EmailsOK As Boolean
    Public Shared EmailAddress As String
    Public Shared EmailPassword As String

    'preferences
    Public Shared MaxResults As Double           'number of results to search (user defined)
    Public Shared NotifyViaPBOK As Boolean
    Public Shared KeepIEVisibleOK As Boolean
    Public Shared AutoTrashOK As Boolean
    Public Shared MinTolerableProfit As Decimal
    Public Shared PostTimingPref As String
    Public Shared SaveWBFilePath As String                 'saveFileAs: output filepath
    Public Shared SaveWBFileName As String
    Public Shared AutoTrashEBooksOK As Boolean
    Public Shared AutoTrashBindersOK As Boolean
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
