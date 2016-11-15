Imports Arbitext.ExcelHelpers
Imports Arbitext.ArbitextHelpers

Public Class SearchSession

    '======== COUNTS ============
    Dim _multiCount As Integer 'count of multi-posts
    Dim _dunnoCount As Integer 'count of ones we couldn't figure out
    Dim _atRow As Integer                 'count of search results that were auto-trashed in this session
    Dim _tooOldPosts As Integer           'count of posts that were too old to show based on preferences
    Dim _skippedResultCount As Integer   'count of skipped search results
    Dim _resultCount As Integer          'count of search results
    Dim _wasAlreadyCategorized As Integer           'count of results already in trash... a subset of skipped results
    Dim _outOfTown As Integer            'count of "nearby" (but not really results)... a subset of skipped results
    Dim _magCount As Integer             'count of magazines in search results
    Dim _keeperCount As Integer          'count of potential winners!
    Dim _negCount As Integer             'count of ones to maybe negoitate for

#Region "Properties"
    Property MultiCount As Integer
        Get
            Return _multiCount
        End Get
        Set(value As Integer)
            _multiCount = value
        End Set
    End Property

    Property DunnoCount As Integer
        Get
            Return _dunnoCount
        End Get
        Set(value As Integer)
            _dunnoCount = value
        End Set
    End Property

    Property ATRow As Integer
        Get
            Return _atRow
        End Get
        Set(value As Integer)
            _atRow = value
        End Set
    End Property

    Property TooOldPosts As Integer
        Get
            Return _tooOldPosts
        End Get
        Set(value As Integer)
            _tooOldPosts = value
        End Set
    End Property

    Property SkippedResultCount As Integer
        Get
            Return _skippedResultCount
        End Get
        Set(value As Integer)
            _skippedResultCount = value
        End Set
    End Property

    Property ResultCount As Integer
        Get
            Return _resultCount
        End Get
        Set(value As Integer)
            _resultCount = value
        End Set
    End Property

    Property WasAlreadyCategorized As Integer
        Get
            Return _wasAlreadyCategorized
        End Get
        Set(value As Integer)
            _wasAlreadyCategorized = value
        End Set
    End Property

    Property OutOfTown As Integer
        Get
            Return _outOfTown
        End Get
        Set(value As Integer)
            _outOfTown = value
        End Set
    End Property

    Property MagCount As Integer
        Get
            Return _magCount
        End Get
        Set(value As Integer)
            _magCount = value
        End Set
    End Property

    Property KeeperCount As Integer
        Get
            Return _keeperCount
        End Get
        Set(value As Integer)
            _keeperCount = value
        End Set
    End Property

    Property NegCount As Integer
        Get
            Return _negCount
        End Get
        Set(value As Integer)
            _negCount = value
        End Set
    End Property
#End Region

#Region "constructors"

    Sub New()
        'set or reset global variables
        ThisAddIn.TldUrl = "http://" & ThisAddIn.City & ".craigslist.org"
        _keeperCount = 0
        _negCount = 0
        _resultCount = 0
        _wasAlreadyCategorized = 0
        _outOfTown = 0
        If ThisAddIn.MaxResults = 0 Then ThisAddIn.MaxResults = 1000000 'just some crazy  high number that we'll never hit if teh user set unlimited results
        _atRow = 0
        _skippedResultCount = 0
        _tooOldPosts = 0
        _magCount = 0
        _multiCount = 0
        _dunnoCount = 0
    End Sub

#End Region

#Region "Methods"

    Sub WrapUp()
        With ThisAddIn.AppExcel
            .Columns("g").Style = "Currency" : .Range("g2").NumberFormat = "General"
            .Columns("h").Style = "Currency"
            .Columns("i").Style = "Currency"
            .Columns("j").Style = "Percent"
            .Columns("k").Style = "Currency"
            .Columns("n").Style = "Percent"

            Dim doneMessage As String
            .ScreenUpdating = True
            doneMessage = "Automated check done examining " & _resultCount & " posts!" & vbCrLf & vbCrLf &
            "KEEPERS: " & _keeperCount & vbCrLf & vbCrLf &
            "TOTAL SKIPPED RESULTS: " & _skippedResultCount & vbCrLf &
            " - Auto-trashed results: " & _atRow & vbCrLf &
            " - Multi-posts: " & _multiCount & vbCrLf &
            " - Magazine posts: " & _magCount & vbCrLf &
            " - Results already in trash: " & _wasAlreadyCategorized & vbCrLf &
            " - ""Nearby"" (but not really) results: " & _outOfTown & vbCrLf &
            " - Too old posts (based on preferences): " & _tooOldPosts & vbCrLf &
            " - Results we otherwise couldn't figure out: " & _dunnoCount
            .StatusBar = False
            .Goto(.Range("A5"), True)

        End With
    End Sub

    Sub WriteSearchResult(post As Post)
        Dim r As Integer = lastUsedRow() + 1

        If post.isMulti Then
            With ThisAddIn.AppExcel
                _multiCount = _multiCount + 1

                If post.Books.Count = 0 Then
                    'doesn't look like we can parse it because bookcount didn't work
                    .Range("b" & r & ":n" & r).Interior.ColorIndex = 20
                    WriteTheEasyStuff(r, post)
                    WriteDummyValues(r)
                    .Range("g" & r).Value2 = post.AskingPrice
                    .Range("f" & r).Value2 = "(multipost, unknown book count)"
                    _dunnoCount = _dunnoCount + 1
                Else 'bookcount worked, let's try to parse it
                    writeMultipostBooks(r, post)
                    'may need to do something right here to catch unparsable multiposts
                End If
            End With
        Else 'not multi
            With ThisAddIn.AppExcel
                writeTheEasyStuff(r, post)
                writeDummyValues(r)
                .Range("g" & r).Value2 = post.AskingPrice 'asking price
                .Range("f" & r).Value2 = post.Book.Isbn13 'isbn
                If Not .Range("F" & r).Value2 Like "*(*" Then
                    .ActiveSheet.Hyperlinks.Add(anchor:= .Range("f" & r), Address:=post.Book.BookscouterSiteLink, TextToDisplay:="'" & .Range("f" & r).Value)
                End If

                'write the numbers if we're good on the result so far
                If Not post.Book.Isbn13 Like "*(*" And post.Book.Isbn13 <> "" Then

                    .Range("h" & r).Value2 = post.Book.BuybackAmount
                    .ActiveSheet.Hyperlinks.Add(anchor:= .Range("h" & r), Address:=post.Book.BuybackLink, TextToDisplay:="'" & .Range("h" & r).Value)

                    'work out the profit if the online selling price came back
                    If Not .Range("h" & r).Value2 Like "*(*" _
                    And Not .Range("g" & r).Value2 Like "*(*" Then
                        WriteProfitCalcs(r, post.Book)
                    End If

                Else
                    _dunnoCount = _dunnoCount + 1
                End If 'paren check on columns F, G, and H

                If post.Book.aLaCarte Then .Rows(r).font.colorindex = 3
                If post.Book.isOBO Then .Rows(r).font.bold = True
                If post.Book.isWeirdEdition Then .Rows(r).font.colorindex = 46
                If post.Book.isPDF Then .Rows(r).font.colorindex = 53
                If post.Book.IsWinner Then HandleWinner(post, post.Book, r)
                If post.Book.IsMaybe Then HandleMaybe(post, post.Book, r)

                AutoCategorizeRowIfNeeded(r, post)
            End With
        End If 'multi handling
    End Sub

    Sub AutoCategorizeRowIfNeeded(ByVal theR As Integer, post As Post)
        If ThisAddIn.AutoTrashOK Then 'checkbox 8 "auto-trash"
            With ThisAddIn.AppExcel
                If IsNumeric(.Range("i" & theR).Value2) Then

                    'first the trash
                    If .Range("n" & theR).Interior.ColorIndex <> 6 _
                    And .Range("i" & theR).Value <= 0 Then
                        categorizeRow(theR, "Trash")
                        ATRow = ATRow + 1
                        SkippedResultCount = SkippedResultCount + 1
                    End If

                    'auto trash those ebooks if that pref is set
                    If ThisAddIn.AutoTrashEBooksOK _
                    AndAlso Not IsDBNull(.Rows(theR).Font.ColorIndex) _
                    AndAlso .Rows(theR).Font.ColorIndex = 53 Then
                        categorizeRow(theR, "Trash")
                        ATRow = ATRow + 1
                        SkippedResultCount = SkippedResultCount + 1
                    End If

                    'auto trash those loose leaf editions if that pref is set
                    If ThisAddIn.AutoTrashBindersOK _
                    AndAlso Not IsDBNull(.Rows(theR).Font.ColorIndex) _
                    AndAlso .Rows(theR).Font.ColorIndex = 3 Then
                        categorizeRow(theR, "Trash")
                        ATRow = ATRow + 1
                        SkippedResultCount = SkippedResultCount + 1
                    End If

                    'then the maybes
                    If .Range("n" & theR).Interior.ColorIndex = 6 Then
                        categorizeRow(theR, "Maybes")
                        ATRow = ATRow + 1
                        SkippedResultCount = SkippedResultCount + 1
                    End If

                    'then the winners
                    If .Range("n" & theR).Interior.ColorIndex = 35 Then
                        categorizeRow(theR, "Keepers")
                        ATRow = ATRow + 1
                        SkippedResultCount = SkippedResultCount + 1
                    End If

                End If
            End With
        End If
    End Sub

    Sub WriteProfitCalcs(ByVal theR As Integer, book As Book)
        With ThisAddIn.AppExcel
            If Not IsNumeric(.Range("g" & theR).Value2) Then
                .Range("i" & theR).Value2 = "?" 'profit
                .Range("j" & theR).Value2 = "?" 'profit margin
                .Range("k" & theR).Value2 = "?" 'min asking price for desired profit margin
                .Range("n" & theR).Value2 = "?" 'price delta
            Else
                .Range("i" & theR).Value2 = book.Profit
                .Range("j" & theR).Value2 = book.ProfitPercentage
                .Range("k" & theR).Value2 = book.MinAskingPriceForDesiredProfit
                .Range("n" & theR).Value2 = book.PriceDelta
            End If
        End With
    End Sub

    Sub WriteTheEasyStuff(ByVal theR As Integer, post As Post)
        With ThisAddIn.AppExcel
            thinInnerBorder(.Range("b" & theR & ":n" & theR))
            .Range("b" & theR).Value2 = theR - 3
            .Range("c" & theR).Value2 = post.PostDate 'date posted
            .Range("d" & theR).Value2 = post.URL 'post url
            .ActiveSheet.Hyperlinks.Add(anchor:= .Range("d" & theR), Address:= .Range("d" & theR).Value, TextToDisplay:= .Range("d" & theR).Value)
            .Range("m" & theR).Value2 = post.UpdateDate
            .Range("l" & theR).Value2 = post.City
            .Range("e" & theR).Value2 = post.Title
            .ActiveSheet.Hyperlinks.Add(anchor:= .Range("e" & theR), Address:=post.AmazonSearchURL, TextToDisplay:= .Range("e" & theR).Value)
        End With
    End Sub

    Sub WriteDummyValues(ByVal theR As Integer)
        'write more dummy values so we can over-write them later if needed
        With ThisAddIn.AppExcel
            .Range("h" & theR).Value2 = "?" 'online selling price
            .Range("i" & theR).Value2 = "?" 'profit
            .Range("j" & theR).Value2 = "?" 'profit margin
            .Range("k" & theR).Value2 = "?" 'min asking price for desired profit margin
            .Range("n" & theR).Value2 = "?" 'price delta
        End With
    End Sub

    Private Function writeMultipostBooks(ByVal theR As Integer, post As Post) As Boolean
        Try
            For b As Short = 0 To post.Books.Count
                With ThisAddIn.AppExcel
                    .StatusBar = "Currently on result number " & _resultCount & " [BOOK " & b + 1 & "] (old, in trash, or otherwise skipped: " & _skippedResultCount & ") (keepers: " & _keeperCount & ") (maybes: " & _negCount & ") [MULTIPOSTS: " & _multiCount & "]"
                    If Not post.alreadyChecked Then
                        theR = lastUsedRow("Automated Checks") + 1
                        If Not wasCategorized(post, post.Books(b).Isbn13, post.Books(b).AskingPrice) Then
                            .Range("b" & theR & ":n" & theR).Interior.ColorIndex = 20
                            WriteTheEasyStuff(theR, post)
                            '.Range("e" & theR).Value2 = post.Books(b).Title
                            .Range("f" & theR).Value2 = post.Books(b).Isbn13 'isbn
                            If Not post.Books(b).Isbn13 Like "*(*" And post.Books(b).Isbn13 <> "" Then
                                .ActiveSheet.Hyperlinks.Add(anchor:= .Range("f" & theR), Address:=post.Books(b).BookscouterSiteLink, TextToDisplay:="'" & .Range("f" & theR).Value)
                            End If
                            .Range("g" & theR).Value2 = post.Books(b).AskingPrice 'asking price
                            .Range("h" & theR).Value2 = post.Books(b).BuybackAmount
                            .ActiveSheet.Hyperlinks.Add(anchor:= .Range("h" & theR), Address:=post.Books(b).BuybackLink & .Range("h" & theR).Value)
                            WriteProfitCalcs(theR, post.Books(b))
                            If post.Books(b).aLaCarte Then .Rows(theR).font.colorindex = 3
                            If post.Books(b).isOBO Then .Rows(theR).font.bold = True
                            If post.Books(b).isWeirdEdition Then .Rows(theR).font.colorindex = 46
                            If post.Books(b).isPDF Then .Rows(theR).font.colorindex = 53
                            If post.Books(b).IsWinner Then HandleWinner(post, post.Books(b), theR)
                            If post.Books(b).IsMaybe Then HandleMaybe(post, post.Books(b), theR)
                            AutoCategorizeRowIfNeeded(theR, post)
                        Else
                            _atRow += 1
                        End If 'was trashed check
                    Else
                        _wasAlreadyCategorized += 1
                    End If 'already checked check
                End With
            Next b
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Sub HandleWinner(post As Post, book As Book, r As Integer)
        _keeperCount += 1
        If ThisAddIn.OnWinnersOK Then
            PushbulletHelpers.sendPushbulletNote(ThisAddIn.Title, "Textbook Winner Found: " & book.Title)
            EmailHelpers.sendSilentNotification(EmailHelpers.emailBodyString(post, book).ToString, "Textbook Winner Found: " & book.Title)
        End If
        SoundHelpers.PlayWAV("money")
        ThisAddIn.AppExcel.Range("b" & r & ":n" & r).Interior.ColorIndex = 35
    End Sub

    Sub HandleMaybe(post As Post, book As Book, r As Integer)
        _negCount += 1
        If ThisAddIn.OnMaybesOK Then
            PushbulletHelpers.sendPushbulletNote(ThisAddIn.Title, "Possible Textbook Lead Found: " & book.Title)
            EmailHelpers.sendSilentNotification(EmailHelpers.emailBodyString(post, book).ToString, "Possible Textbook Lead Found: " & book.Title)
        End If
        ThisAddIn.AppExcel.Range("b" & r & ":n" & r).Interior.ColorIndex = 6
    End Sub


#End Region

End Class
