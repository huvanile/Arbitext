Imports mshtml
Imports Arbitext.CraigslistHelpers
Imports Arbitext.ExcelHelpers
Imports Arbitext.StringHelpers
Imports Arbitext.ArbitextHelpers

Public Class Post
    Private _url As String                'CL post url
    Private _body As String               'body of the post
    Private _city As String               'CL post city (from parsing post page) 'public bc also used by wasCategorized function
    Private _askingPrice As Integer       'CL asking price (from parsing post page) 'public bc also used by wasCategorized function
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
            LearnAboutMultipost()
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

    ReadOnly Property AskingPrice As Integer
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
        Dim tmpISBN As String
        Dim tmpAskingPrice As String
        Dim s As Long            'splitholder start position, increments as books are parsed
        Dim t As Integer         'to loop through the splitholder <BR> results
        Dim splitholder() As String
        Dim tmpPostBody As String
        Dim tmpBook As Book

        'fist try the "1br" separator approach 
        tmpISBN = ""
        tmpAskingPrice = ""
        s = 0
        splitholder = Split(_body, "<br>")
        t = 0
        For t = s To UBound(splitholder)
            tmpISBN = getISBN(splitholder(t), URL)
            tmpAskingPrice = getAskingPrice(splitholder(t))
            If Not tmpISBN Like "(*" And Not tmpAskingPrice Like "*(*" Then 'it's actually a book result!
                tmpBook = New Book(tmpISBN, tmpAskingPrice, clean(splitholder(t), False, False, False, False, False, False))
                _books.Add(tmpBook)
                tmpBook = Nothing
                Exit For 'this single book is found, go on to next splitholder result
            End If
        Next t

        'then try the "p" separator approach, adding all new books found as we go
        tmpISBN = ""
        tmpAskingPrice = ""
        s = 0
        t = 0
        splitholder = Split(_body, "<p>")
        For t = s To UBound(splitholder)
            tmpISBN = getISBN(splitholder(t), URL)
            tmpAskingPrice = getAskingPrice(splitholder(t))
            If Not tmpISBN Like "(*" And Not tmpAskingPrice Like "*(*" Then 'it's actually a book result!
                If _books.Count > 0 Then
                    For i As Short = 0 To _books.Count 'should this be a 1 base instead of a zero base?
                        If _books(i).Isbn13 <> tmpISBN And _books(i).Isbn10 <> tmpISBN Then
                            tmpBook = New Book(tmpISBN, tmpAskingPrice, clean(splitholder(t), False, False, False, False, False, False))
                            _books.Add(tmpBook)
                            tmpBook = Nothing
                        End If
                    Next i
                Else
                    tmpBook = New Book(tmpISBN, tmpAskingPrice, clean(splitholder(t), False, False, False, False, False, False))
                    _books.Add(tmpBook)
                    tmpBook = Nothing
                End If
                Exit For 'this single book is found, go on to next splitholder result
            End If
        Next t

        'then try the "2br" separator approach, adding all new books found as we go
        tmpISBN = ""
        tmpAskingPrice = ""
        s = 0
        t = 0
        tmpPostBody = Replace(_body, Chr(10), "")
        tmpPostBody = Replace(tmpPostBody, Chr(13), "")
        splitholder = Split(tmpPostBody, "<br><br>")
        For t = s To UBound(splitholder)
            tmpISBN = getISBN(splitholder(t), URL)
            tmpAskingPrice = getAskingPrice(splitholder(t))
            If Not tmpISBN Like "(*" And Not tmpAskingPrice Like "*(*" Then 'it's actually a book result!
                If _books.Count > 0 Then
                    For i As Short = 0 To _books.Count 'should this be a 1 base instead of a zero base?
                        If _books(i).Isbn13 <> tmpISBN And _books(i).Isbn10 <> tmpISBN Then
                            tmpBook = New Book(tmpISBN, tmpAskingPrice, clean(splitholder(t), False, False, False, False, False, False))
                            _books.Add(tmpBook)
                            tmpBook = Nothing
                        End If
                    Next i
                Else
                    tmpBook = New Book(tmpISBN, tmpAskingPrice, clean(splitholder(t), False, False, False, False, False, False))
                    _books.Add(tmpBook)
                    tmpBook = Nothing
                End If
                Exit For 'this single book is found, go on to next splitholder result
            End If
        Next t

    End Sub

    Sub LearnAboutPost()
        Dim tmp As String = ""
        Dim wc As New Net.WebClient
        Dim bHTML() As Byte = wc.DownloadData(_url)
        Dim sHTML As String = New UTF8Encoding().GetString(bHTML)
        Dim doc As IHTMLDocument2 = New HTMLDocument
        doc.clear()
        doc.write(sHTML)
        Dim allElements As IHTMLElementCollection = doc.all
        Dim element As IHTMLElement

        'find title
        _title = doc.title

        'find price
        Dim allSpans As IHTMLElementCollection = allElements.tags("span")
        For Each element In allSpans
            If element.className = "price" Then
                _askingPrice = Trim(element.innerText)
                Exit For
            End If
        Next
        allSpans = Nothing
        If _askingPrice = -1 Then
            _askingPrice = getAskingPrice(_html)
        End If

        'find location
        Dim allSmalls As IHTMLElementCollection = allElements.tags("small")
        For Each element In allSmalls
            Dim tmpLoc As String = Trim(element.innerText)
            tmpLoc = Replace(tmpLoc, ")", "")
            tmpLoc = Replace(tmpLoc, "(", "")
            tmpLoc = Replace(tmpLoc, "\n", "")
            tmpLoc = Replace(tmpLoc, "\r", "")
            _city = tmpLoc
            Exit For
        Next
        allSmalls = Nothing

        'find og:image
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
                _postDate = Trim(element.innerText)
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
        Dim allSections As IHTMLElementCollection = allElements.tags("section")
        For Each element In allSections
            If element.id = "postingbody" Then
                _body = Trim(element.innerText)
                Exit For
            End If
        Next
        allSections = Nothing

        'clean up
        doc.close()
        element = Nothing
        doc = Nothing
        wc = Nothing
        bHTML = Nothing
        allElements = Nothing

    End Sub

#End Region

End Class
