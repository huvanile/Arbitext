Imports Arbitext.ExcelHelpers
Imports Arbitext.ArbitextHelpers
Imports Arbitext.CraigslistHelpers

Public Class MultiplePostsAnalysis
    Dim _checkedPostsNotBooks As List(Of Post)      'list of posts checked before populating the post object with the accompanying books
    Dim _checkedPostsAndBooks As List(Of Post)      'list of posts checked with all of the books included
#Region "constructors"

    Sub New()
        ThisAddIn.TldUrl = "http://" & ThisAddIn.City & ".craigslist.org"
        If ThisAddIn.MaxResults = 0 Then ThisAddIn.MaxResults = 1000000 'just some crazy  high number that we'll never hit if teh user set unlimited results
        _checkedPostsNotBooks = New List(Of Post)
        _checkedPostsAndBooks = New List(Of Post)
        allQuerySearch()
    End Sub

#End Region

#Region "Methods"

    Private Sub allQuerySearch()
        ThisAddIn.Proceed = True
        ThisAddIn.AppExcel.ScreenUpdating = True
        If Not doesWSExist("Color Legend") Then BuildWSColorLegend.buildWSColorLegend()
        If ThisAddIn.Proceed Then
            Dim searchURL As String = ""

            'isbn search
            searchURL = ThisAddIn.TldUrl & "/search/sss?query=isbn&sort=rel"
            If ThisAddIn.PostTimingPref Like "*Today*" Then searchURL = searchURL & "&postedToday=1"
            oneQuerySearch(searchURL)

            'textbook search
            If ThisAddIn.Proceed Then
                searchURL = ThisAddIn.TldUrl & "/search/sss?query=textbook&sort=rel"
                If ThisAddIn.PostTimingPref Like "*Today*" Then searchURL = searchURL & "&postedToday=1"
                oneQuerySearch(searchURL)
            End If

            'books search
            If ThisAddIn.Proceed Then
                searchURL = ThisAddIn.TldUrl & "/search/bka?query=college"
                If ThisAddIn.PostTimingPref Like "*Today*" Then searchURL = searchURL & "&postedToday=1"
                oneQuerySearch(searchURL)
            End If

            MsgBox("Done! " & vbCrLf & vbCrLf & "Partially checked posts:  " & _checkedPostsNotBooks.Count & vbCrLf & vbCrLf & "Fully checked posts:  " & _checkedPostsAndBooks.Count, vbOKOnly, ThisAddIn.Title)
            ThisAddIn.AppExcel.ScreenUpdating = True
            ThisAddIn.AppExcel.StatusBar = False
        End If
    End Sub

    ''' <summary>
    ''' basically this one filters out the unneeded posts and then requests that we only write the good ones
    ''' this one also scrapes the entire cl search response page and loops through each posting
    ''' </summary>
    ''' <param name="searchURL">URL of search result page</param>
    Private Sub oneQuerySearch(searchURL As String)
        Dim postNotBooks As Post                                    'this is a partially populated post object, just the post and not the books
        Dim postAndBooks As Post                                    'this is a fully populated post object, including the books
        Dim wc As New System.Net.WebClient
        Dim resultURL As String : resultURL = ""                     'URL of search result page
        Dim updatedSearchURL As String : updatedSearchURL = ""       'search result URL... this gets iterated through pagination
        Dim startPos As Integer : startPos = 1 'this is the start position of the search in the search results and will be incremented to find different results
        Dim searchPage As String : searchPage = ""                   'HTML of whole search result page

        'set or reset global variables
        searchPage = wc.DownloadString(searchURL)
        If Not searchPage Like "*Nothing found for that search*" Then
            'start scraping
            Do While InStr(startPos, searchPage, ThisAddIn.ResultHook) > 0

                'if max results not yet exceeded
                If _checkedPostsNotBooks.Count < ThisAddIn.MaxResults Then
                    ThisAddIn.AppExcel.DisplayStatusBar = True
                    ThisAddIn.AppExcel.StatusBar = "Cursory examination of post " & getLinkFromCLSearchResults(searchPage, startPos) & ", will ignore if out of town"

                    'if not a nearby result
                    If Not getLinkFromCLSearchResults(searchPage, startPos) Like "*http*" And Not getLinkFromCLSearchResults(searchPage, startPos) Like "*.org*" Then 'this prevents it from showing "nearby results"
                        ThisAddIn.AppExcel.StatusBar = "Learning more about post " & getLinkFromCLSearchResults(searchPage, startPos) & "(search result number " & _checkedPostsNotBooks.Count + 1 & ")"

                        postNotBooks = New Post(ThisAddIn.TldUrl & getLinkFromCLSearchResults(searchPage, startPos), False)

                        'ThisAddIn.AppExcel.StatusBar = "Currently on result number " & _checkedPosts.Count & " (old, in trash, or otherwise skipped: " & SearchSession.SkippedResultCount & ") (keepers: " & SearchSession.KeeperCount & ") (maybes: " & SearchSession.NegCount & ") [MULTIPOSTS: " & SearchSession.MultiCount & "]"
                        ThisAddIn.AppExcel.StatusBar = "Writing search result number " & _checkedPostsNotBooks.Count + 1 & " to workbook"

                        'make sure it wasn't aleady checked this session
                        Dim match As Boolean = False
                        For Each p In _checkedPostsNotBooks
                            If p.Equals(postNotBooks) Then
                                match = True
                                Exit For
                            End If
                        Next

                        If Not match _
                        AndAlso Not postNotBooks.IsMagazinePost _
                        AndAlso Not postNotBooks.IsTooOld Then
                            If postNotBooks.IsParsable Then
                                postAndBooks = New Post(postNotBooks)
                                WriteSearchResult(postAndBooks)
                                _checkedPostsAndBooks.Add(postAndBooks)
                            Else
                                If Not doesWSExist("Unparseable posts") Then createWS("Unparseable posts")
                                Dim r As Int16 = lastUsedRow("Unparseable posts") + 1
                                With ThisAddIn.AppExcel.Sheets("Unparseable posts")
                                    .range("a" & r).value2 = postNotBooks.Title
                                    .range("b" & r).value2 = postNotBooks.URL
                                    .Hyperlinks.Add(anchor:= .Range("b" & r), Address:=postNotBooks.URL, TextToDisplay:=postNotBooks.URL)
                                End With
                            End If
                        End If
                        _checkedPostsNotBooks.Add(postNotBooks)
                    End If

                    'do the pagination
                    startPos = InStr(startPos, searchPage, ThisAddIn.ResultHook) + 100
                    If _checkedPostsNotBooks.Count Mod 100 = 0 Then
                        If InStr(1, searchURL, "&") = 0 Then
                            updatedSearchURL = searchURL & "?s=" & _checkedPostsNotBooks.Count
                        Else
                            updatedSearchURL = searchURL & "&s=" & _checkedPostsNotBooks.Count
                        End If
                        searchPage = wc.DownloadString(updatedSearchURL)
                        startPos = 1 'to start searching at the top of the newly loaded search page
                    End If

                Else 'max count check
                    GoTo maxReached
                End If

            Loop
        Else
            MsgBox("Nothing found for the search:" & vbCrLf & vbCrLf & searchURL, vbInformation, ThisAddIn.Title)
        End If
        wc = Nothing
        Exit Sub
maxReached:

        'wrap up

        wc = Nothing
        ThisAddIn.Proceed = False
    End Sub

    Sub WriteSearchResult(post As Post)
        Dim destSheet As String
        For Each b As Book In post.Books
            If Not b.WasAlreadyChecked Then
                b.GetDataFromBookscouter()
                If post.IsParsable AndAlso b.IsParsable Then
                    If ThisAddIn.AutoCategorizeOK Then
                        If b.IsWinner(post) Then
                            If Not doesWSExist("Winners") Then BuildWSResults.buildResultWS("Winners")
                            destSheet = "Winners"
                        ElseIf b.IsMaybe(post) Then
                            If Not doesWSExist("Maybes") Then BuildWSResults.buildResultWS("Maybes")
                            destSheet = "Maybes"
                        ElseIf b.IsTrash(post) Then
                            If Not doesWSExist("Trash") Then BuildWSResults.buildResultWS("Trash") Else unFilterTrash()
                            destSheet = "Trash"
                        Else
                            If Not doesWSExist("Automated Checks") Then BuildWSResults.buildResultWS("Automated Checks")
                            destSheet = "Automated Checks"
                        End If
                    Else
                        If Not doesWSExist("Automated Checks") Then BuildWSResults.buildResultWS("Automated Checks")
                        destSheet = "Automated Checks"
                    End If
                Else
                    If Not doesWSExist("Unparseable posts") Then createWS("Unparseable posts")
                    Dim r As Int16 = lastUsedRow("Unparseable posts") + 1
                    With ThisAddIn.AppExcel.Sheets("Unparseable posts")
                        .range("a" & r).value2 = post.Title
                        .range("b" & r).value2 = post.URL
                        .Hyperlinks.Add(anchor:= .Range("b" & r), Address:=post.URL, TextToDisplay:=post.URL)
                    End With
                    Exit Sub
                End If

                With ThisAddIn.AppExcel.Sheets(destSheet)
                    .activate
                    Dim r As Integer = lastUsedRow() + 1
                    thinInnerBorder(.Range("a" & r & ":k" & r))

                    'write post stuff
                    .Range("a" & r).Value2 = post.PostDate 'date posted
                    .Range("b" & r).Value2 = post.UpdateDate 'date updated
                    .Hyperlinks.Add(anchor:= .Range("c" & r), Address:=post.URL, TextToDisplay:=post.Title)
                    .range("k" & r).value2 = post.City

                    'write book stuff
                    If b.IsParsable Then
                        .Hyperlinks.Add(anchor:= .Range("d" & r), Address:=b.AmazonSearchURL, TextToDisplay:=b.Title)
                        .Hyperlinks.Add(anchor:= .Range("e" & r), Address:=b.BookscouterSiteLink, TextToDisplay:="'" & b.IsbnFromPost)
                        .Range("f" & r).Value2 = b.AskingPrice
                        .Range("g" & r).Value2 = b.BuybackAmount
                        .range("h" & r).value2 = b.Profit
                        .range("i" & r).value2 = b.ProfitPercentage
                        .range("j" & r).value2 = b.MinAskingPriceForDesiredProfit
                        If b.aLaCarte(post) Then .Rows(r).font.colorindex = 3
                        If b.isOBO(post) Then .Rows(r).font.bold = True
                        If b.isWeirdEdition(post) Then .Rows(r).font.colorindex = 46
                        If b.isPDF(post) Then .Rows(r).font.colorindex = 53
                        If b.IsWinner(post) Then HandleWinner(post, b, r)
                        If b.IsMaybe(post) Then HandleMaybe(post, b, r)
                    End If
                End With
            Else
                'already checked, so don't write it
            End If
        Next b

    End Sub

    'Sub WrapUp()
    '    Dim doneMessage As String
    '    .ScreenUpdating = True
    '        doneMessage = "Automated check done examining " & _resultCount & " posts!" & vbCrLf & vbCrLf &
    '        "KEEPERS: " & _keeperCount & vbCrLf & vbCrLf &
    '        "TOTAL SKIPPED RESULTS: " & _skippedResultCount & vbCrLf &
    '        " - Auto-trashed results: " & _atRow & vbCrLf &
    '        " - Magazine posts: " & _magCount & vbCrLf &
    '        " - Results already in trash: " & _wasAlreadyCategorized & vbCrLf &
    '        " - ""Nearby"" (but not really) results: " & _outOfTown & vbCrLf &
    '        " - Too old posts (based on preferences): " & _tooOldPosts & vbCrLf &
    '        " - Results we otherwise couldn't figure out: " & _dunnoCount
    '        ThisAddIn.AppExcel.StatusBar = False
    '    ThisAddIn.AppExcel.Goto(.Range("A5"), True)
    'End Sub

    Sub HandleWinner(post As Post, book As Book, r As Integer)
        If ThisAddIn.OnWinnersOK Then
            PushbulletHelpers.sendPushbulletNote(ThisAddIn.Title, "Textbook Winner Found: " & book.Title)
            EmailHelpers.sendSilentNotification(EmailHelpers.emailBodyString(post, book).ToString, "Textbook Winner Found: " & book.Title)
        End If
        SoundHelpers.PlayWAV("money")
        ThisAddIn.AppExcel.Range("a" & r & ":k" & r).Interior.ColorIndex = 35
    End Sub

    Sub HandleMaybe(post As Post, book As Book, r As Integer)
        If ThisAddIn.OnMaybesOK Then
            PushbulletHelpers.sendPushbulletNote(ThisAddIn.Title, "Possible Textbook Lead Found: " & book.Title)
            EmailHelpers.sendSilentNotification(EmailHelpers.emailBodyString(post, book).ToString, "Possible Textbook Lead Found: " & book.Title)
        End If
        ThisAddIn.AppExcel.Range("a" & r & ":k" & r).Interior.ColorIndex = 6
    End Sub

#End Region
End Class
