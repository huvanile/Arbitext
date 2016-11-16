Imports Arbitext.ArbitextHelpers
Imports Arbitext.ExcelHelpers
Imports Arbitext.StringHelpers
Imports System.Xml

Public Class Book
    Private _title As String
    Private _imageURL As String
    Private _isParsable As Boolean
    Private _isbn10 As String
    Private _isbn13 As String
    Private _author As String
    Private _buybackAmount As Decimal
    Private _buybackSite As String
    Private _buybackLink As String
    Private _bookscouterAPILink As String
    Private _bookscouterSiteLink As String
    Private _askingPrice As Decimal
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
        If GetDataFromBookscouter() Then _isParsable = True Else _isParsable = False
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

    ReadOnly Property AmazonSearchURL As String
        Get
            Return "http://www.amazon.com/s/ref=nb_sb_noss?url=search-alias%3Daps&field-keywords=" & replacePlusWithSpace(_title)
        End Get
    End Property

    ReadOnly Property IsParsable As Boolean
        Get
            If Isbn13 = "" Or Isbn13 Like "*(*" Then Return False Else Return True
        End Get
    End Property

    ReadOnly Property IsTrash(post As Post) As Boolean
        Get
            IsTrash = False
            If IsParsable Then
                If isPDF(post) Then IsTrash = True
                If isWeirdEdition(post) Then IsTrash = True
                If aLaCarte(post) Then IsTrash = True
                If Profit <= 0 Then IsTrash = True
            End If
        End Get
    End Property

    ReadOnly Property IsWinner(post As Post) As Boolean
        Get
            IsWinner = False
            If IsParsable Then
                If Profit > ThisAddIn.MinTolerableProfit And Not IsTrash(post) Then IsWinner = True
            End If
        End Get
    End Property

    ReadOnly Property IsMaybe(post As Post) As Boolean
        Get
            IsMaybe = False
            If IsParsable Then
                If Not IsWinner(post) And Not IsTrash(post) Then IsMaybe = True
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

    ReadOnly Property isPDF(post As Post) As Boolean
        Get
            If post.Books.Count > 1 Then
                If CraigslistHelpers.isPDF(_saleDescInPost) Then Return True Else Return False
            Else
                If CraigslistHelpers.isPDF(post.Title) Or CraigslistHelpers.isPDF(post.Body) Then Return True Else Return False
            End If
        End Get
    End Property

    ReadOnly Property isOBO(post As Post) As Boolean
        Get
            If post.Books.Count > 1 Then
                If CraigslistHelpers.isOBO(_saleDescInPost) Then Return True Else Return False
            Else
                If CraigslistHelpers.isOBO(post.Title) Or CraigslistHelpers.isOBO(post.Body) Then Return True Else Return False
            End If
        End Get
    End Property

    ReadOnly Property isWeirdEdition(post As Post) As Boolean
        Get
            If post.Books.Count > 1 Then
                If CraigslistHelpers.isWeirdEdition(_saleDescInPost) Then Return True Else Return False
            Else
                If CraigslistHelpers.isWeirdEdition(post.Title) Or CraigslistHelpers.isWeirdEdition(post.Body) Then Return True Else Return False
            End If
        End Get
    End Property

    ReadOnly Property aLaCarte(post As Post) As Boolean
        Get
            If post.Books.Count > 1 Then
                If CraigslistHelpers.aLaCarte(_saleDescInPost) Then Return True Else Return False
            Else
                If CraigslistHelpers.aLaCarte(post.Title) Or CraigslistHelpers.aLaCarte(post.Body) Then Return True Else Return False
            End If
        End Get
    End Property

    ReadOnly Property WasAlreadyChecked() As Boolean
        Get
            WasAlreadyChecked = False
            WasAlreadyChecked = findInResultSheet("Unparsable")
            If Not WasAlreadyChecked Then WasAlreadyChecked = findInResultSheet("Winners")
            If Not WasAlreadyChecked Then WasAlreadyChecked = findInResultSheet("Maybes")
            If Not WasAlreadyChecked Then WasAlreadyChecked = findInResultSheet("Trash")
            If Not WasAlreadyChecked Then WasAlreadyChecked = findInResultSheet("Automated Checks")
        End Get
    End Property

    Private Function findInResultSheet(sheet As String)
        If doesWSExist(sheet) Then
            If canFind(_title, sheet,, False, False) Then
                With ThisAddIn.AppExcel.Sheets(sheet)
                    Dim theRow As Int16 : theRow = .range(canFind(_title, sheet,, True, False)).row
                    If .range("d" & theRow).value2 = _title _
                        AndAlso .range("e" & theRow).value2 = _isbn13 _
                        AndAlso .range("f" & theRow).value2 = _askingPrice Then
                        Return True
                    Else
                        Return False
                    End If
                End With
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function
#End Region

#Region "Methods"

    Public Function GetDataFromBookscouter() As Boolean
        Try

            Dim wc As New Net.WebClient
            Dim isbn As String : If _isbn13 = "" Then isbn = _isbn10 Else isbn = _isbn13
            ThisAddIn.AppExcel.DisplayStatusBar = True
            ThisAddIn.AppExcel.StatusBar = "Learning about book: " & isbn
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
                _imageURL = Replace(ImageURL, "._SL75_.", "._SL300_.")
                nodes = doc.GetElementsByTagName("isbn10") : _isbn10 = nodes(0).InnerText.Trim()
                nodes = doc.GetElementsByTagName("isbn13") : _isbn13 = nodes(0).InnerText.Trim()
                nodes = doc.GetElementsByTagName("author") : _author = nodes(0).InnerText.Trim() & vbCrLf
                nodes = doc.GetElementsByTagName("amount") : _buybackAmount = nodes(0).InnerText.Trim
                nodes = doc.GetElementsByTagName("vendor") : _buybackSite = nodes(0).InnerText.Trim()
                nodes = doc.GetElementsByTagName("link") : _buybackLink = nodes(0).InnerText.Trim()
                Return True
            Else
                _isbn10 = "(unknown)"
                _isbn13 = "(unknown)"
                _title = "(unknown)"
                _imageURL = "(unknown)"
                _author = "(unknown)"
                _buybackAmount = -1
                _buybackSite = "(unknown)"
                _buybackLink = "(unknown)"
                Return False
            End If
            wc = Nothing
            nodes = Nothing
            doc = Nothing
        Catch ex As Exception
            _isbn10 = "(unknown)"
            _isbn13 = "(unknown)"
            _title = "(unknown)"
            _imageURL = "(unknown)"
            _author = "(unknown)"
            _buybackAmount = -1
            _buybackSite = "(unknown)"
            _buybackLink = "(unknown)"
            Return False
        End Try
    End Function

#End Region

End Class
