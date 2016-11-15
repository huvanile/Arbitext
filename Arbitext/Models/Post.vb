Imports mshtml
Imports Arbitext.CraigslistHelpers
Imports Arbitext.ExcelHelpers
Imports Arbitext.StringHelpers
Imports Arbitext.ArbitextHelpers

Public Class Post
    Private _url As String                'CL post url
    Private _body As String               'body of the post
    Private _city As String               'CL post city (from parsing post page) 'public bc also used by wasCategorized function
    Private _askingPrice As Decimal       'CL asking price (from parsing post page) 'public bc also used by wasCategorized function
    Private _title As String              'CL post title (from parsing post page) 'public bc also used by wasCategorized function
    Private _updateDate As String         'CL post updated
    Private _postDate As String           'CL post date
    Private _isbn As String               'CL post ISBN
    Private _image As String              'url of image in craiglist post
    Private _html As String               'html of post
    Private _book As Book
    Private _books As List(Of Book)

#Region "Constructors"

    ' Default constructor
    Sub New(url As String)
        _url = url
        _isbn = ""
        _askingPrice = -1
        _postDate = "?"
        _updateDate = "-"
        _title = "?"
        _city = "-"
        _body = ""
        _image = ""
        _book = Nothing
        _books = Nothing

        LearnAboutPost()

        If isMulti Then
            _books = New List(Of Book)
            tryMultiSplit("<br>")
            tryMultiSplit("<br><br>")
            tryMultiSplit("<p>")
        Else
            _book = New Book(_isbn, _askingPrice, _body)
        End If

    End Sub

#End Region

#Region "Easy Properties"

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

    ReadOnly Property AmazonSearchURL As String
        Get
            Return "http://www.amazon.com/s/ref=nb_sb_noss?url=search-alias%3Daps&field-keywords=" & replacePlusWithSpace(_title)
        End Get
    End Property

    ''' <summary>
    ''' Used to determine if a post has already been seen and categorized or not
    ''' </summary>
    ''' <returns>True if alreadyChecked, False if not</returns>
    ReadOnly Property alreadyChecked() As Boolean
        Get
            alreadyChecked = False

            With ThisAddIn.AppExcel.Sheets("Automated Checks")

                If isMulti Then

                    'to check for parsable multiposts
                    If Not _isbn = "" And Not _isbn Like "*(*" Then
                        If canFind(_isbn, "Automated Checks") Then
                            If Trim(.Range("g" & .Range(canFind(_isbn, "Automated Checks", , True, False)).Row).Value) = _askingPrice _
                            And .Range("l" & .Range(canFind(_isbn, "Automated Checks", , True, False)).Row).Value = _city _
                            And .Range("f" & .Range(canFind(_isbn, "Automated Checks", , True, False)).Row).Value = _isbn Then
                                Return True
                            End If
                        End If
                    End If

                    'to check for unparsable multiposts
                    If canFind(_title, "Automated Checks") Then
                        If Trim(.Range("g" & .Range(canFind(_title, "Automated Checks", , True, False)).Row).Value) = _askingPrice _
                        And LCase(.Range("f" & .Range(canFind(_title, "Automated Checks", , True, False)).Row).Value) Like "*multi*" _
                        And .Range("l" & .Range(canFind(_title, "Automated Checks", , True, False)).Row).Value = _city Then
                            Return True
                        End If
                    End If

                Else 'not multi

                    If canFind(_title, "Automated Checks") Then
                        If Trim(.Range("g" & .Range(canFind(_title, "Automated Checks", , True, False)).Row).Value) = _askingPrice _
                        And .Range("l" & .Range(canFind(_title, "Automated Checks", , True, False)).Row).Value = _city Then
                            Return True
                        End If
                    End If

                End If

            End With
        End Get
    End Property

    Property Books As List(Of Book)
        Get
            Return _books
        End Get
        Set(value As List(Of Book))
            _books = value
        End Set
    End Property

    Property Book As Book
        Get
            Return _book
        End Get
        Set(value As Book)
            _book = value
        End Set
    End Property

    ReadOnly Property bookCount() As Integer
        Get
            Return bookCountFromString(_html)
        End Get
    End Property

    ReadOnly Property isMulti() As Boolean
        Get
            If CraigslistHelpers.isMulti(_body) Or CraigslistHelpers.isMulti(_title) Then Return True Else Return False
        End Get
    End Property

#End Region

#Region "Methods"

    Sub LearnAboutMultipost()

    End Sub

    Sub tryMultiSplit(splitter As String)
        Dim tmpISBN As String = ""
        Dim tmpAskingPrice As Decimal = -1
        Dim s As Long = 0                   'splitholder start position, increments as books are parsed
        Dim t As Integer = 0                'to loop through the splitholder <BR> results
        Dim splitholder() As String
        Dim tmpPostBody As String = Replace(_body, Chr(10), "")
        tmpPostBody = Replace(tmpPostBody, Chr(13), "")
        splitholder = Split(tmpPostBody, splitter)
        For t = s To UBound(splitholder)
            tmpISBN = getISBN(splitholder(t), URL)
            tmpAskingPrice = getAskingPrice(splitholder(t))
            If Not tmpISBN Like "(*" And Not tmpAskingPrice = -1 Then 'it's actually a book result!
                Dim tmpBook = New Book(tmpISBN, tmpAskingPrice, clean(splitholder(t), False, False, False, False, False, False))
                Dim match As Boolean = False
                For Each z As Book In _books
                    If z.Equals(tmpBook) Then 'shouldn't ever be more than 3 really
                        match = True
                        Exit For
                    End If
                Next z
                If Not match And bookCountFromString(tmpBook.SaleDescInPost) < 4 Then _books.Add(tmpBook)
                tmpBook = Nothing
            End If
        Next t
    End Sub

    Sub LearnAboutPost()
        Dim tmp As String = ""
        Dim wc As New Net.WebClient
        Dim bHTML() As Byte = wc.DownloadData(_url)
        Dim sHTML As String = New UTF8Encoding().GetString(bHTML)
        Dim doc As IHTMLDocument = New HTMLDocument
        'doc.clear()
        doc.write(sHTML)
        Dim allElements As IHTMLElementCollection = doc.all
        Dim element As IHTMLElement

        'find title
        _title = doc.title

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

        'find location (HAP)
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
        Dim metaTag As HTMLMetaElement
        For Each element In allElements
            If element.tagName = "META" Then
                metaTag = element
                If LCase(metaTag.content) Like "*images.craigslist*" Then
                    _image = metaTag.content
                    Exit For
                End If
            End If
        Next
        metaTag = Nothing

        Dim allTimes As IHTMLElementCollection = allElements.tags("time")

        'find date posted
        For Each element In allTimes
            If LCase(element.parentElement.innerText) Like "*posted:*" Then
                _postDate = Trim(element.parentElement.innerText)
                _postDate = Replace(_postDate, "posted:", "").Trim
                Exit For
            End If
        Next

        'find date updated
        For Each element In allTimes
            If LCase(element.parentElement.innerText) Like "*updated:*" Then
                _updateDate = Trim(element.innerText)
                Exit For
            End If
        Next

        allTimes = Nothing

        'get posting body
        Dim z As Long : z = 0
        Dim splitholder
        Dim m As String : m = ""
        z = Strings.InStr(1, sHTML, "<section id=""postingbody"">") 'start of section
        m = Right(sHTML, Len(sHTML) - z)
        z = Strings.InStr(1, m, "</div>")
        m = Right(m, Len(m) - z)
        z = Strings.InStr(1, m, "</div>")
        m = Right(m, Len(m) - z - 5)
        splitholder = Split(m, "</section>") 'end boundary of section 
        m = Trim(splitholder(0))
        _body = m

        'get isbn
        If Not isMulti Then
            _isbn = getISBN(_body, _url)
            If Not _isbn.Length = 13 And Not _isbn.Length = 10 Then _isbn = getISBN(_title, _url)
        End If

        'clean up
        doc.close()
        element = Nothing
        'doc = Nothing
        wc = Nothing
        bHTML = Nothing
        allElements = Nothing

    End Sub

#End Region

End Class
