﻿Imports ArbitextClassLibrary.StringHelpers
Imports System.Xml
Imports System.Text

Public Class Book
    Private _title As String
    Private _imageURL As String
    Private _isParsable As Boolean
    Private _isbn10 As String
    Private _isbn13 As String
    Private _author As String
    Private _hasSiblings As Boolean         'is the book part of a multipost?
    Private _parentTitle As String          'title of the parent post
    Private _parentBody As String           'html body of the parent post
    Private _parentPostDate As String       'the date of the original / parent post
    Private _buybackAmount As Decimal
    Private _buybackSite As String
    Private _buybackLink As String
    Private _isbnFromPost As String
    Private _bookscouterAPILink As String
    Private _bookscouterSiteLink As String
    Private _askingPrice As Decimal
    Private _saleDescInPost As String     'the part of the post used to describe this book. for single-book posts, this is teh whole post body.

#Region "Contructors"

    Sub New(isbn As String, askingPrice As Decimal, theSaleDesc As String, queryBS As Boolean, hasSiblings As Boolean, parentTitle As String, parentBody As String, parentPostDate As String)
        _isbnFromPost = isbn.Trim
        _hasSiblings = hasSiblings
        _parentBody = parentBody
        _parentTitle = parentTitle
        _parentPostDate = parentPostDate
        If isbn.Trim.Length = 13 Or isbn.Trim.Length = 10 Then
            If isbn.Length = 13 Then _isbn13 = isbn Else _isbn10 = isbn
            _bookscouterAPILink = "http://api.bookscouter.com/prices.php?isbn=" & isbn & "&uid=" & randomUID() & ""
            _bookscouterSiteLink = "https://href.li/?https://bookscouter.com/prices.php?isbn=" & isbn & "&all"
            _saleDescInPost = theSaleDesc.Trim
            _askingPrice = askingPrice
            If queryBS Then
                If GetDataFromBookscouter() Then _isParsable = True Else _isParsable = False
            Else
                _isParsable = True
            End If
        Else
            _isParsable = False
        End If
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
            Return _title
        End Get
        Set(value As String)
            _title = value.Trim
        End Set
    End Property

    Property Isbn10 As String
        Get
            Return _isbn10
        End Get
        Set(value As String)
            _isbn10 = value.Trim
        End Set
    End Property

    Property Isbn13 As String
        Get
            Return _isbn13
        End Get
        Set(value As String)
            _isbn13 = value.Trim
        End Set
    End Property

#End Region

#Region "Readonly Properties"

    ''' <summary>
    ''' Unique book sale ID
    ''' </summary>
    ''' <returns>Unique book sale ID based on post date, ISBN from post, and asking price</returns>
    ReadOnly Property ID As String
        Get
            Dim tmp As New StringBuilder
            tmp.Append(Year(_parentPostDate) & Month(_parentPostDate) & Day(_parentPostDate) & Hour(_parentPostDate) & Minute(_parentPostDate))
            tmp.Append(Trim(IsbnFromPost))
            tmp.Append(Trim(AskingPrice))
            Return tmp.ToString
        End Get
    End Property

    ReadOnly Property IsbnFromPost As String
        Get
            Return _isbnFromPost
        End Get
    End Property

    ReadOnly Property AmazonSearchURL As String
        Get
            Return "https://href.li/?http://www.amazon.com/s/ref=nb_sb_noss?url=search-alias%3Daps&field-keywords=" & replacePlusWithSpace(_title)
        End Get
    End Property

    ReadOnly Property IsParsable As Boolean
        Get
            If (Isbn13 = "" Or Isbn13 Like "*(*") _
            And (Isbn10 = "" Or Isbn10 Like "*(*") Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    ReadOnly Property IsTrash() As Boolean
        Get
            IsTrash = False
            If IsParsable Then
                If isPDF() Then IsTrash = True
                If isWeirdEdition() Then IsTrash = True
                If aLaCarte() Then IsTrash = True
                If Profit <= 0 Then IsTrash = True
            End If
        End Get
    End Property

    ReadOnly Property IsWinner() As Boolean
        Get
            IsWinner = False
            If IsParsable Then
                If Profit > 15 AndAlso AskingPrice > 0 AndAlso Not IsTrash() Then IsWinner = True
            End If
        End Get
    End Property

    ReadOnly Property IsMaybe() As Boolean
        Get
            IsMaybe = False
            If IsParsable Then
                If Not IsWinner() AndAlso Not IsTrash() AndAlso AskingPrice > 0 Then IsMaybe = True
            End If
        End Get
    End Property

    ''' <summary>
    ''' HVSB = High Value Stale Book.  These can sell online for a fair amount, weren't originally profitable based on asking price, and are now stale.
    ''' These can be profitable if they take a low-ball offer.
    ''' </summary>
    ''' <returns>True if the book is a HVSB</returns>
    ReadOnly Property IsHVSB As Boolean
        Get
            IsHVSB = False
            If IsParsable _
                AndAlso IsWinner _
                AndAlso Not IsMaybe _
                AndAlso AskingPrice > 0 _
                AndAlso BuybackAmount >= 35 _
                AndAlso DateDiff(DateInterval.Day, CDate(_parentPostDate), Now()) >= 14 Then
                Return True
            End If
        End Get
    End Property

    ''' <summary>
    ''' HVOBO = High Value Or Best Offer.  These can sell online for a fair amoubt, aren't profitable at the given asking price, and the seller said they'd take the best offer.
    ''' </summary>
    ''' <returns></returns>
    ReadOnly Property IsHVOBO As Boolean
        Get
            IsHVOBO = False
            If IsParsable _
            AndAlso Not IsWinner _
            AndAlso Not IsMaybe _
            AndAlso Not IsHVSB _
            AndAlso (isOBO Or AskingPrice = -2) _
            AndAlso BuybackAmount >= 35 Then
                Return True
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
                If _buybackAmount - 15 > 0 Then
                    Return Math.Round(_buybackAmount - 15, 0)
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
            If _buybackAmount > 0 Then
                Dim tmp As Decimal = Math.Round(_buybackAmount - _askingPrice, 2)
                If tmp > 0 Then Return tmp Else Return 0
            Else
                Return 0
            End If
        End Get
    End Property

    ReadOnly Property ProfitPercentage As Decimal
        Get
            If _buybackAmount > 0 Then
                Dim tmp As Decimal = Profit / _askingPrice
                If tmp > 0 Then Return tmp Else Return 0
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
            Return "https://href.li/?https://bookscouter.com/prices.php?isbn=" & _isbn13 & "&all"
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

    ReadOnly Property isPDF() As Boolean
        Get
            If _hasSiblings Then
                If CraigslistHelpers.isPDF(_saleDescInPost) Then Return True Else Return False
            Else
                If CraigslistHelpers.isPDF(_parentTitle) Or CraigslistHelpers.isPDF(_parentBody) Then Return True Else Return False
            End If
        End Get
    End Property

    ReadOnly Property isOBO() As Boolean
        Get
            If _hasSiblings Then
                If CraigslistHelpers.isOBO(_saleDescInPost) Then Return True Else Return False
            Else
                If CraigslistHelpers.isOBO(_parentTitle) Or CraigslistHelpers.isOBO(_parentBody) Then Return True Else Return False
            End If
        End Get
    End Property

    ReadOnly Property isWeirdEdition() As Boolean
        Get
            If _hasSiblings Then
                If CraigslistHelpers.isWeirdEdition(_saleDescInPost) Then Return True Else Return False
            Else
                If CraigslistHelpers.isWeirdEdition(_parentTitle) Or CraigslistHelpers.isWeirdEdition(_parentBody) Then Return True Else Return False
            End If
        End Get
    End Property

    ReadOnly Property aLaCarte() As Boolean
        Get
            If _hasSiblings Then
                If CraigslistHelpers.aLaCarte(_saleDescInPost) Then Return True Else Return False
            Else
                If CraigslistHelpers.aLaCarte(_parentTitle) Or CraigslistHelpers.aLaCarte(_parentBody) Then Return True Else Return False
            End If
        End Get
    End Property

    'ReadOnly Property WasAlreadyChecked() As Boolean
    '    Get
    '        WasAlreadyChecked = False
    '        If Not WasAlreadyChecked Then WasAlreadyChecked = findInResultSheet("Winners")
    '        If Not WasAlreadyChecked Then WasAlreadyChecked = findInResultSheet("Maybes")
    '        If Not WasAlreadyChecked Then WasAlreadyChecked = findInResultSheet("Trash")
    '        If Not WasAlreadyChecked Then WasAlreadyChecked = findInResultSheet("HVSBs")
    '        If Not WasAlreadyChecked Then WasAlreadyChecked = findInResultSheet("Automated Checks")
    '    End Get
    'End Property

    'Private Function findInResultSheet(sheet As String)
    '    Try
    '        If canFindInResultCol("L", ID, sheet) Then Return True Else Return False
    '    Catch ex As Exception
    '        Return False
    '    End Try

    'End Function

#End Region

#Region "Methods"

    Public Function GetDataFromBookscouter() As Boolean
        Try
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
