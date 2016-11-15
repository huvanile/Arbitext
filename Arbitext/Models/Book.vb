Imports Arbitext.ArbitextHelpers
Imports System.Xml

Public Class Book
    Private _title As String
    Private _imageURL As String
    Private _isbn10 As String
    Private _isbn13 As String
    Private _author As String
    Private _isPDF As Boolean
    Private _isOBO As Boolean
    Private _isWeirdEdition As Boolean
    Private _aLaCarte As Boolean
    Private _buybackAmount As Decimal
    Private _buybackSite As String
    Private _buybackLink As String
    Private _bookscouterAPILink As String
    Private _bookscouterSiteLink As String
    Private _askingPrice As Decimal
    Private _aLaCarteFlag As Boolean      'flag for loose leaf editions when found in saleDescInPost
    Private _weirdEditionFlag As Boolean  'flag for weird editions like teacher's edition when found in saleDescInPost
    Private _pdfFlag As Boolean           'flag for pdf files being sold as books on saleDescInPost
    Private _oboFlag As Boolean           'flag for when OBO (or best offer) is present in saleDescInPost
    Private _saleDescInPost As String     'the part of the post used to describe this book. for single-book posts, this is teh whole post body.

#Region "Contructors"

    Sub New(isbn As String, askingPrice As Decimal, theSaleDesc As String)
        If isbn.Length = 13 Then
            _isbn13 = isbn
            _bookscouterAPILink = "http://api.bookscouter.com/prices.php?isbn=" & isbn & "&uid=" & randomUID() & ""
            _bookscouterSiteLink = "https://bookscouter.com/prices.php?isbn=" & isbn & "&all"

        ElseIf isbn.Length = 10 Then
            _isbn10 = isbn
            _bookscouterAPILink = "http://api.bookscouter.com/prices.php?isbn=" & isbn & "&uid=" & randomUID() & ""
            _bookscouterSiteLink = "https://bookscouter.com/prices.php?isbn=" & isbn & "&all"
        Else

        End If

        _saleDescInPost = theSaleDesc
        _askingPrice = askingPrice
        GetDataFromBookscouter()
    End Sub
#End Region

#Region "Read & Write Properties"

    Property AskingPrice As Decimal
        Get
            Return _askingPrice
        End Get
        Set(value As Decimal)
            _askingPrice = value
        End Set
    End Property

    Property Title As String
        Get
            Return _title.Trim
        End Get
        Set(value As String)
            _title = value.Trim
        End Set
    End Property

    Property Isbn10 As String
        Get
            Return _isbn10.Trim
        End Get
        Set(value As String)
            _isbn10 = value.Trim
        End Set
    End Property

    Property Isbn13 As String
        Get
            Return _isbn13.Trim
        End Get
        Set(value As String)
            _isbn13 = value.Trim
        End Set
    End Property

#End Region

#Region "Readonly Properties"

    ReadOnly Property IsWinner As Boolean
        Get
            If Profit > ThisAddIn.MinTolerableProfit _
            And Not aLaCarte _
            And Not isWeirdEdition _
            And Not isPDF Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    ReadOnly Property IsMaybe As Boolean
        Get
            If (Profit >= 1 Or PriceDelta <= 0.25) _
            And Not _buybackAmount = 0 _
            And Not IsWinner _
            And Not aLaCarte _
            And Not isWeirdEdition _
            And Not isPDF Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    ReadOnly Property PriceDelta As Decimal
        Get
            If _buybackAmount > 0 Then
                Return 1 - (MinAskingPriceForDesiredProfit / _askingPrice)
            Else
                Return 0
            End If
        End Get
    End Property

    ReadOnly Property MinAskingPriceForDesiredProfit As Decimal
        Get
            If _buybackAmount > 0 Then
                If _buybackAmount - ThisAddIn.MinTolerableProfit > 0 Then
                    Return Math.Round(_buybackAmount - ThisAddIn.MinTolerableProfit, 0)
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        End Get
    End Property
    ReadOnly Property SaleDescInPost As String
        Get
            Return _saleDescInPost
        End Get
    End Property

    ReadOnly Property Profit As Decimal
        Get
            Dim tmp As Decimal = 0
            If _buybackAmount > 0 Then
                tmp = ThisAddIn.AppExcel.Round(_buybackAmount - _askingPrice, 2)
                If tmp > 0 Then Return tmp Else Return 0
            Else
                Return 0
            End If
        End Get
    End Property

    ReadOnly Property ProfitPercentage As Decimal
        Get
            If _buybackAmount > 0 Then
                Return Profit / _askingPrice
            Else
                Return 0
            End If
        End Get
    End Property

    ReadOnly Property BookscouterAPILink As String
        Get
            Return "http://api.bookscouter.com/prices.php?isbn=" & _isbn13 & "&uid=" & randomUID() & ""
        End Get
    End Property

    ReadOnly Property BookscouterSiteLink As String
        Get
            Return "https://bookscouter.com/prices.php?isbn=" & _isbn13 & "&all"
        End Get
    End Property

    ReadOnly Property Author As String
        Get
            Return _author.Trim
        End Get
    End Property

    ReadOnly Property ImageURL As String
        Get
            Return _imageURL.Trim
        End Get
    End Property

    ReadOnly Property BuybackAmount As Short
        Get
            Return _buybackAmount
        End Get
    End Property

    ReadOnly Property BuybackSite As String
        Get
            Return _buybackSite.Trim
        End Get
    End Property

    ReadOnly Property BuybackLink As String
        Get
            Return _buybackLink.Trim
        End Get
    End Property

#End Region

#Region "FLAG Properties"

    Property isPDF() As Boolean
        Get
            If CraigslistHelpers.isPDF(_saleDescInPost) Then Return True Else Return False
        End Get
        Set(value As Boolean)
            _isPDF = value
        End Set
    End Property

    Property isOBO() As Boolean
        Get
            If CraigslistHelpers.isOBO(_saleDescInPost) Then Return True Else Return False
        End Get
        Set(value As Boolean)
            _isOBO = value
        End Set
    End Property

    Property isWeirdEdition() As Boolean
        Get
            If CraigslistHelpers.isWeirdEdition(_saleDescInPost) Then Return True Else Return False
        End Get
        Set(value As Boolean)
            _isWeirdEdition = value
        End Set
    End Property

    Property aLaCarte() As Boolean
        Get
            If CraigslistHelpers.aLaCarte(_saleDescInPost) Then Return True Else Return False
        End Get
        Set(value As Boolean)
            _aLaCarte = value
        End Set
    End Property

#End Region

#Region "Methods"

    Public Function GetDataFromBookscouter() As Boolean
        Dim wc As New Net.WebClient
        Dim isbn As String : If _isbn13 = "" Then isbn = _isbn10 Else isbn = _isbn13
        Dim bXML() As Byte = wc.DownloadData(_bookscouterAPILink)
        Dim sXML As String = New UTF8Encoding().GetString(bXML)
        Dim doc As XmlDocument = New XmlDocument
        Dim nodes As XmlNodeList
        doc.LoadXml(sXML)
        wc = Nothing
        bXML = Nothing
        If doc.ChildNodes.Count >= 2 Then 'its 2 if its a bad isbn, and 0 if not response at all
            nodes = doc.GetElementsByTagName("title") : _title = nodes(0).InnerText.Trim()
            nodes = doc.GetElementsByTagName("image") : _imageURL = nodes(0).InnerText.Trim()
            _imageURL = Replace(ImageURL, "._SL75_.", "._SL400_.")
            nodes = doc.GetElementsByTagName("isbn10") : _isbn10 = nodes(0).InnerText.Trim()
            nodes = doc.GetElementsByTagName("isbn13") : _isbn13 = nodes(0).InnerText.Trim()
            nodes = doc.GetElementsByTagName("author") : _author = nodes(0).InnerText.Trim() & vbCrLf
            nodes = doc.GetElementsByTagName("amount") : _buybackAmount = nodes(0).InnerText.Trim
            nodes = doc.GetElementsByTagName("vendor") : _buybackSite = nodes(0).InnerText.Trim()
            nodes = doc.GetElementsByTagName("link") : _buybackLink = nodes(0).InnerText.Trim()
            Return True
        Else
            _title = "(unknown)"
            _imageURL = "(unknown)"
            _author = "(unknown)"
            _buybackAmount = "(unknown)"
            _buybackSite = "(unknown)"
            _buybackLink = "(unknown)"
            Return False
        End If
        nodes = Nothing
        doc = Nothing
    End Function

#End Region

End Class
