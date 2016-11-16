Imports Arbitext.ExcelHelpers
Imports Arbitext.ArbitextHelpers
Imports Arbitext.CraigslistHelpers
Imports Microsoft.Office.Interop.Excel

Public Class MultiplePostsAnalysis
    Dim _checkedPosts As List(Of Post)    'list of posts checked in this session

#Region "constructors"

    Sub New()
        ThisAddIn.TldUrl = "http://" & ThisAddIn.City & ".craigslist.org"
        If ThisAddIn.MaxResults = 0 Then ThisAddIn.MaxResults = 1000000 'just some crazy  high number that we'll never hit if teh user set unlimited results
        _checkedPosts = New List(Of Post)
        allQuerySearch()
    End Sub

#End Region

#Region "Methods"

    Private Sub allQuerySearch()
        ThisAddIn.Proceed = True
        ThisAddIn.AppExcel.ScreenUpdating = True
        prepareWBForAutomatedSearch()
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

            MsgBox("Done!" & _checkedPosts.Count & " post checked", vbOKOnly, ThisAddIn.Title)
            ThisAddIn.AppExcel.ScreenUpdating = True
            ThisAddIn.AppExcel.StatusBar = False
        End If
    End Sub

    Private Sub prepareWBForAutomatedSearch()
        unFilterTrash()
        If Not doesWSExist("Automated Checks") Then BuildWSResults.buildResultWS("Automated Checks")
        If Not doesWSExist("Color Legend") Then BuildWSColorLegend.buildWSColorLegend()
        ThisAddIn.AppExcel.Sheets("Automated Checks").activate
        If lastUsedRow("Automated Checks") > 3 Then
            Select Case MsgBox("Should we empty the 'Automated Checks' page before starting the search?", vbYesNoCancel, ThisAddIn.Title)
                Case vbYes
                    DeleteWS("Automated Checks")
                    BuildWSResults.buildResultWS("Automated Checks")
                Case vbNo

                Case Else
                    ThisAddIn.AppExcel.ScreenUpdating = True
                    ThisAddIn.AppExcel.StatusBar = False
                    ThisAddIn.Proceed = False
            End Select
        End If
    End Sub

    ''' <summary>
    ''' basically this one filters out the unneeded posts and then requests that we only write the good ones
    ''' this one also scrapes the entire cl search response page and loops through each posting
    ''' </summary>
    ''' <param name="searchURL">URL of search result page</param>
    Private Sub oneQuerySearch(searchURL As String)
        Dim post As Post
        Dim wc As New System.Net.WebClient
        Dim resultURL As String : resultURL = ""                     'URL of search result page
        Dim updatedSearchURL As String : updatedSearchURL = ""       'search result URL... this gets iterated through pagination
        Dim startPos As Integer : startPos = 1 'this is the start position of the search in the search results and will be incremented to find different results
        Dim searchPage As String : searchPage = ""                   'HTML of whole search result page

        'set or reset global variables
        searchPage = wc.DownloadString(searchURL)
        'start scraping
        Do While InStr(startPos, searchPage, ThisAddIn.ResultHook) > 0

            'if max results not yet exceeded
            If _checkedPosts.Count < ThisAddIn.MaxResults Then

                'if not a nearby result
                If Not getLinkFromCLSearchResults(searchPage, startPos) Like "*http*" And Not getLinkFromCLSearchResults(searchPage, startPos) Like "*.org*" Then 'this prevents it from showing "nearby results"
                    ThisAddIn.AppExcel.StatusBar = "Learning about result number " & _checkedPosts.Count + 1

                    post = New Post(ThisAddIn.TldUrl & getLinkFromCLSearchResults(searchPage, startPos))

                    'ThisAddIn.AppExcel.StatusBar = "Currently on result number " & _checkedPosts.Count & " (old, in trash, or otherwise skipped: " & SearchSession.SkippedResultCount & ") (keepers: " & SearchSession.KeeperCount & ") (maybes: " & SearchSession.NegCount & ") [MULTIPOSTS: " & SearchSession.MultiCount & "]"
                    ThisAddIn.AppExcel.StatusBar = "Writing result number " & _checkedPosts.Count + 1 & " to workbook"

                    'make sure it wasn't aleady checked this session
                    Dim match As Boolean = False
                    For Each p In _checkedPosts
                        If p.Equals(post) Then
                            match = True
                            Exit For
                        End If
                    Next

                    If Not match _
                    AndAlso Not post.IsMagazinePost _
                    AndAlso Not post.IsTooOld Then
                        If post.IsParsable Then
                            WriteSearchResult(post)
                        Else
                            If Not doesWSExist("Unparseable posts") Then createWS("Unparseable posts")
                            Dim r As Int16 = ExcelHelpers.lastUsedRow("Unparseable posts") + 1
                            ThisAddIn.AppExcel.Sheets("Unparseable posts").range("a" & r).value2 = post.Title
                            ThisAddIn.AppExcel.Sheets("Unparseable posts").range("b" & r).value2 = post.URL
                        End If
                    End If
                    _checkedPosts.Add(post)
                End If

                'do the pagination
                startPos = InStr(startPos, searchPage, ThisAddIn.ResultHook) + 100
                If _checkedPosts.Count Mod 100 = 0 Then
                    If InStr(1, searchURL, "&") = 0 Then
                        updatedSearchURL = searchURL & "?s=" & _checkedPosts.Count
                    Else
                        updatedSearchURL = searchURL & "&s=" & _checkedPosts.Count
                    End If
                    searchPage = wc.DownloadString(updatedSearchURL)
                    startPos = 1 'to start searching at the top of the newly loaded search page
                End If

            Else 'max count check
                GoTo maxReached
            End If

        Loop
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
                    'determine sheet to write result in
                    If post.IsParsable AndAlso b.IsParsable Then
                        If ThisAddIn.AutoCategorizeOK Then
                            If b.IsWinner(post) Then
                                If Not doesWSExist("Winners") Then BuildWSResults.buildResultWS("Winners")
                                destSheet = "Winners"
                            ElseIf b.IsMaybe(post) Then
                                If Not doesWSExist("Maybes") Then BuildWSResults.buildResultWS("Maybes")
                                destSheet = "Maybes"
                            ElseIf b.IsTrash(post) Then
                                If Not doesWSExist("Trash") Then BuildWSResults.buildResultWS("Trash")
                                destSheet = "Trash"
                            Else
                                destSheet = "Automated Checks"
                            End If
                        Else
                            destSheet = "Automated Checks"
                        End If
                    Else
                    If Not doesWSExist("Unparseable books") Then BuildWSResults.buildResultWS("Unparsable books")
                    destSheet = "Unparsable books"
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
                            .Hyperlinks.Add(anchor:= .Range("e" & r), Address:=b.BookscouterSiteLink, TextToDisplay:=b.Isbn13)
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
