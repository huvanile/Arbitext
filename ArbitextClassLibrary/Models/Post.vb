Imports mshtml
Imports ArbitextClassLibrary.CraigslistHelpers
Imports ArbitextClassLibrary.StringHelpers
Imports System.Text

Public Class Post
    Private _url As String                'CL post url
    Private _body As String               'body of the post
    Private _city As String               'CL post city (from parsing post page) 
    Private _askingPrice As Decimal       'CL asking price (from parsing post page) 
    Private _title As String              'CL post title (from parsing post page) 
    Private _updateDate As String         'CL post updated
    Private _postDate As String           'CL post date
    Private _isbn As String               'CL post ISBN
    Private _image As String              'url of image in craiglist post
    Private _html As String               'html of post
    Private _books As List(Of Book)
    Private _phones As List(Of Phone)
    Private _isParsable As Boolean        'is this post parseable?
    Private _searchterm As String

#Region "Constructors"

    Sub New(url As String, Optional learnAboutBooks As Boolean = True)
        _url = url
        _body = ""
        _city = "-"
        _askingPrice = -1
        _title = "?"
        _updateDate = "-"
        _postDate = "?"
        _isbn = "(?)"
        _image = ""
        _html = ""
        _books = Nothing
        If LearnAboutPost() Then
            If learnAboutBooks Then
                _books = New List(Of Book)
                findBooksInPost("<br>")
                findBooksInPost("<br><br>")
                findBooksInPost("<p>")
                findBooksInPost("")
                If _books.Count > 0 Then _isParsable = True Else _isParsable = False
            Else
                If _askingPrice <> -1 And Not _isbn Like "(" Then _isParsable = True Else _isParsable = False
            End If
        Else
            _isParsable = False
        End If
    End Sub

    Sub New(learnAboutPhones As Boolean, url As String)
        _url = url
        _body = ""
        _city = "-"
        _askingPrice = -1
        _title = "?"
        _updateDate = "-"
        _postDate = "?"
        _searchterm = "(?)"
        _image = ""
        _html = ""
        _books = Nothing
        If LearnAboutPost(True) Then
            _phones = New List(Of Phone)
            findPhonesInPost()
            If _phones.Count > 0 And _askingPrice <> -1 And _askingPrice <> 1 Then _isParsable = True Else _isParsable = False
        Else
            If _askingPrice <> -1 And _askingPrice <> 1 And Not _searchterm Like "(" Then _isParsable = True Else _isParsable = False
        End If
    End Sub

    Sub New(preCheckedPost As Post)
        _url = preCheckedPost.URL
        _body = preCheckedPost.Body
        _city = preCheckedPost.City
        _askingPrice = preCheckedPost.AskingPrice
        _title = preCheckedPost.Title
        _updateDate = preCheckedPost.UpdateDate
        _postDate = preCheckedPost.PostDate
        _isbn = preCheckedPost.Isbn
        _image = preCheckedPost.Image
        _html = preCheckedPost._html
        _books = New List(Of Book)
        findBooksInPost("<br>", False)
        findBooksInPost("<br><br>", False)
        findBooksInPost("<p>", False)
        findBooksInPost("", False)
        If _books.Count > 0 Then _isParsable = True Else _isParsable = False
    End Sub

    Sub New(preCheckedPost As Post, forPhones As Boolean)
        _url = preCheckedPost.URL
        _body = preCheckedPost.Body
        _city = preCheckedPost.City
        _askingPrice = preCheckedPost.AskingPrice
        _title = preCheckedPost.Title
        _updateDate = preCheckedPost.UpdateDate
        _postDate = preCheckedPost.PostDate
        _searchterm = preCheckedPost.SearchTerm
        _image = preCheckedPost.Image
        _html = preCheckedPost._html
        _phones = New List(Of Phone)
        findPhonesInPost()
        If _phones.Count > 0 Then _isParsable = True Else _isParsable = False
    End Sub

#End Region

#Region "Easy Properties"

    Property IsParsable As Boolean
        Get
            Return _isParsable
        End Get
        Set(value As Boolean)
            _isParsable = value
        End Set
    End Property

    ReadOnly Property PostDate As String
        Get
            Return _postDate.Trim
        End Get
    End Property

    Property UpdateDate As String
        Get
            Return _updateDate.Trim
        End Get
        Set(value As String)
            _updateDate = value.Trim
        End Set
    End Property

    ReadOnly Property Title As String
        Get
            Return _title.Trim
        End Get
    End Property

    ReadOnly Property AskingPrice As Decimal
        Get
            Return _askingPrice
        End Get
    End Property

    ReadOnly Property URL As String
        Get
            Return _url.Trim
        End Get
    End Property

    ReadOnly Property SearchTerm As String
        Get
            Return _searchterm
        End Get
    End Property
    ReadOnly Property Isbn As String
        Get
            Return _isbn.Trim
        End Get
    End Property

    ReadOnly Property Image As String
        Get
            Return _image.Trim
        End Get
    End Property

    ReadOnly Property Body As String
        Get
            Return _body.Trim
        End Get
    End Property

    ReadOnly Property City As String
        Get
            Return _city.Trim
        End Get
    End Property

#End Region

#Region "Not Easy Properties"

    ReadOnly Property IsMagazinePost As Boolean
        Get
            If LCase(_body) Like "*magazine*" Or LCase(_title) Like "*magazine*" Then Return True Else Return False
        End Get
    End Property

    ReadOnly Property AmazonSearchURL As String
        Get
            Return "https://href.li/?http://www.amazon.com/s/ref=nb_sb_noss?url=search-alias%3Daps&field-keywords=" & replacePlusWithSpace(_title)
        End Get
    End Property

    'ReadOnly Property WasAlreadyChecked() As Boolean
    '    Get
    '        WasAlreadyChecked = False
    '        findInResultSheet("Unparsable")
    '        If Not WasAlreadyChecked Then WasAlreadyChecked = findInResultSheet("Winners")
    '        If Not WasAlreadyChecked Then WasAlreadyChecked = findInResultSheet("Maybes")
    '        If Not WasAlreadyChecked Then WasAlreadyChecked = findInResultSheet("Trash")
    '        If Not WasAlreadyChecked Then WasAlreadyChecked = findInResultSheet("Automated Checks")
    '    End Get
    'End Property

    'Private Function findInResultSheet(sheet As String)
    '    If doesWSExist(sheet) Then
    '        If canFindInResultCol("C", _title, sheet) Then Return True Else Return False
    '    Else
    '        Return False
    '    End If
    'End Function

    Property Phones As List(Of Phone)
        Get
            Return _phones
        End Get
        Set(value As List(Of Phone))
            _phones = value
        End Set
    End Property

    Property Books As List(Of Book)
        Get
            Return _books
        End Get
        Set(value As List(Of Book))
            _books = value
        End Set
    End Property

    ReadOnly Property isMulti() As Boolean
        Get
            If CraigslistHelpers.isMulti(_body) Or CraigslistHelpers.isMulti(_title) Then Return True Else Return False
        End Get
    End Property

#End Region

#Region "Methods"

    Sub findPhonesInPost()
        Dim tmpphone As Phone
        If _html.Contains("<span>model name / number: <b>") Then
            _searchterm = getMakeAndModel(_html)
            If _searchterm.Length < 5 Then _searchterm = _title
        Else
            _searchterm = _title
            If _searchterm.Length < 5 Then _searchterm = getMakeAndModel(_html)
        End If
        tmpphone = New Phone(_searchterm, _askingPrice, _body, True, False, _title, _body, _updateDate)
        _phones.Add(tmpphone)
    End Sub

    Sub findBooksInPost(Optional splitter As String = "", Optional queryBS As Boolean = True)
        Dim tmpISBN As String = ""
        Dim tmpAskingPrice As Decimal = -1
        Dim s As Long = 0                   'splitholder start position, increments as books are parsed
        Dim t As Integer = 0                'to loop through the splitholder <BR> results
        Dim splitholder() As String
        If splitter = "" Then
            tmpISBN = getISBN(_body, URL)
            tmpAskingPrice = getAskingPrice(_html)
            Dim tmpBook = New Book(tmpISBN, tmpAskingPrice, clean(_body, False, False, False, False, False, False), queryBS, CraigslistHelpers.isMulti(_body), _title, _body, _postDate)
            Dim match As Boolean = False
            For Each b As Book In _books
                If b.Equals(tmpBook) Or b.Title = tmpBook.Title Then
                    match = True
                    Exit For
                End If
            Next b
            If Not match And bookCountFromString(tmpBook.SaleDescInPost) < 4 Then _books.Add(tmpBook)
            tmpBook = Nothing
        Else
            Dim tmpPostBody As String = Replace(_body, Chr(10), "")
            tmpPostBody = Replace(tmpPostBody, Chr(13), "")
            splitholder = Split(tmpPostBody, splitter)
            For t = s To UBound(splitholder)
                tmpISBN = getISBN(splitholder(t), URL)
                tmpAskingPrice = getAskingPrice(splitholder(t))
                If Not tmpISBN Like "(*" And Not tmpAskingPrice = -1 Then 'it's actually a book result!
                    Dim tmpBook = New Book(tmpISBN, tmpAskingPrice, clean(splitholder(t), False, False, False, False, False, False), queryBS, CraigslistHelpers.isMulti(_body), _title, _body, _postDate)
                    Dim match As Boolean = False
                    For Each b As Book In _books
                        If b.Equals(tmpBook) Or b.Title = tmpBook.Title Then
                            match = True
                            Exit For
                        End If
                    Next b
                    If Not match And bookCountFromString(tmpBook.SaleDescInPost) < 4 Then _books.Add(tmpBook)
                    tmpBook = Nothing
                End If
            Next t
        End If
    End Sub

    Function LearnAboutPost(Optional forPhones As Boolean = False) As Boolean

        Dim tmp As String = ""
        Dim wc As New Net.WebClient
        Dim bHTML() As Byte = wc.DownloadData(_url)
        _html = New UTF8Encoding().GetString(bHTML)
        Dim doc As mshtml.IHTMLDocument2 = New mshtml.HTMLDocument
        doc.clear()
        doc.write(_html)
        Dim allElements As mshtml.IHTMLElementCollection
        allElements = doc.all
        doc.close()
        Dim element As mshtml.IHTMLElement
        Try

            'find title
            _title = doc.title

            If Not _title Like "*Page Not Found*" Then
                'find price
                For Each element In allElements.tags("span")
                    If element.className = "price" Then
                        _askingPrice = Trim(element.innerText)
                        Exit For
                    End If
                Next
                If _askingPrice = -1 Then
                    _askingPrice = getAskingPrice(_html)
                End If

                'find location (MAP)
                For Each element In allElements.tags("small")
                    Dim tmpLoc As String = Trim(element.innerText)
                    tmpLoc = Replace(tmpLoc, ")", "")
                    tmpLoc = Replace(tmpLoc, "(", "")
                    tmpLoc = Replace(tmpLoc, "\n", "")
                    tmpLoc = Replace(tmpLoc, "\r", "")
                    _city = tmpLoc
                    Exit For
                Next

                ''find og:image
                For Each element In allElements
                    If element.tagName = "META" Then
                        Dim metaTag As HTMLMetaElement = element
                        If LCase(metaTag.content) Like "*images.craigslist*" Then
                            _image = metaTag.content
                            Exit For
                        End If
                    End If
                Next

                'find date posted
                For Each element In allElements.tags("time")
                    If LCase(element.parentElement.innerText) Like "*posted:*" Then
                        _postDate = Trim(element.parentElement.innerText)
                        _postDate = Replace(_postDate, "posted:", "").Trim
                        Exit For
                    End If
                Next

                'find date updated
                For Each element In allElements.tags("time")
                    If LCase(element.parentElement.innerText) Like "*updated:*" Then
                        _updateDate = Trim(element.innerText)
                        Exit For
                    End If
                Next
                If _updateDate = "" Or _updateDate = "-" Then _updateDate = _postDate

                'get posting body
                Dim z As Long : z = 0
                Dim splitholder
                Dim m As String : m = ""
                z = Strings.InStr(1, _html, "<section id=""postingbody"">") 'start of section
                m = Right(_html, Len(_html) - z)
                z = Strings.InStr(1, m, "</div>")
                m = Right(m, Len(m) - z)
                z = Strings.InStr(1, m, "</div>")
                m = Right(m, Len(m) - z - 5)
                splitholder = Split(m, "</section>") 'end boundary of section 
                m = Trim(splitholder(0))
                _body = m

                If forPhones Then

                Else
                    'get isbn
                    _isbn = getISBN(_html, _url)
                End If


                'clean up
                element = Nothing
                doc = Nothing
                wc = Nothing
                bHTML = Nothing
                allElements = Nothing

                If forPhones Then
                    Return True
                Else
                    If _isbn Like "*(*" Or _askingPrice = -1 Then Return False Else Return True
                End If

            Else
                'page not found title
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try
    End Function

#End Region

End Class
