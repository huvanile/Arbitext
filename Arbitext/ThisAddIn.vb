Public Class ThisAddIn
    Public Const Title As String = "Arbitext"
    Public Const ResultHook As String = "<li class=""result-row"" data-pid="
    Public Shared AppExcel As Excel.Application
    Public Shared Proceed As Boolean
    Public Shared AppIE As Object
    Public Shared frmPrefs As FrmPrefs

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
    Public Shared MinTolerableProfit As Integer
    Public Shared PostTimingPref As String
    Public Shared SaveWBFilePath As String                 'saveFileAs: output filepath
    Public Shared SaveWBFileName As String
    Public Shared AutoTrashEBooksOK As Boolean
    Public Shared AutoTrashBindersOK As Boolean
    Public Shared OnWinnersOK As Boolean
    Public Shared OnMaybesOK As Boolean
    Public Shared City As String

    'about the CL post
    Public Shared PostURL As String            'CL post url
    Public Shared PostBody As String           'body of the post
    Public Shared PostCity As String           'CL post city (from parsing post page) 'public bc also used by wasCategorized function
    Public Shared PostAskingPrice As String    'CL asking price (from parsing post page) 'public bc also used by wasCategorized function
    Public Shared PostTitle As String          'CL post title (from parsing post page) 'public bc also used by wasCategorized function
    Public Shared TldUrl As String             'top-level domain name (e.g., http://dallas.craigslist.com)
    Public Shared PostUpdateDate As String     'CL post updated
    Public Shared PostDate As String           'CL post date
    Public Shared PostISBN As String           'CL post ISBN
    Public Shared PostSellingPrice As String   'BS selling price
    Public Shared ALaCarteFlag As Boolean      'flag for loose leaf editions
    Public Shared WeirdEditionFlag As Boolean  'flag for weird editions like teacher's edition
    Public Shared PdfFlag As Boolean           'flag for pdf files being sold as books on CL
    Public Shared OboFlag As Boolean           'flat for when OBO (or best offer) is present in CL post
    Public Shared PostCLImg As String          'url of image in craiglist post

    'about the bookscouter result
    Public Shared ImageURL As String           'url of book cover per the bookscouter search result
    Public Shared BsTitle As String            'title of book per the bookscouter search result
    Public Shared BsBuybackLink As String      'best buyback link from bookscouter search result

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        AppExcel = Me.Application
        RegistryHelpers.addressNullPrefs()
        ThisAddIn.Proceed = True
        RegistryHelpers.loadVariablesFromRegistry()
    End Sub
End Class
