Imports Arbitext.ExcelHelpers
Imports Arbitext.ArbitextHelpers
Imports Arbitext.CraigslistHelpers
Imports Microsoft.Office.Interop.Excel

Public Class AutomatedChecks
    Shared searchSession As SearchSession

    Public Shared Sub AutomatedChecks()
        ThisAddIn.Proceed = True
        ThisAddIn.AppExcel.ScreenUpdating = True
        automatedChecksPageCheck()
        prepareWBForAutomatedSearch()
        If ThisAddIn.Proceed Then
            Dim searchURL As String = ""
            searchSession = New SearchSession

            'isbn search
            searchURL = ThisAddIn.TldUrl & "/search/sss?query=isbn&sort=rel"
            If ThisAddIn.PostTimingPref Like "*today*" Then searchURL = searchURL & "&postedToday=1"
            automatedSearch(searchURL)

            'textbook search
            If ThisAddIn.Proceed Then
                searchURL = ThisAddIn.TldUrl & "/search/sss?query=textbook&sort=rel"
                If ThisAddIn.PostTimingPref Like "*today*" Then searchURL = searchURL & "&postedToday=1"
                automatedSearch(searchURL)
            End If

            'books search
            If ThisAddIn.Proceed Then
                searchURL = ThisAddIn.TldUrl & ".craigslist.org/search/ndf/bks"
                If ThisAddIn.PostTimingPref Like "*today*" Then searchURL = searchURL & "&postedToday=1"
                automatedSearch(searchURL)
            End If

            'remove dupes
            If ThisAddIn.Proceed Then
                Dim oColumns(0) As Object
                For i As Short = 2 To 13
                    oColumns.SetValue(i, 0)
                Next
                Dim header As Excel.XlYesNoGuess = XlYesNoGuess.xlYes
                ThisAddIn.AppExcel.Range("B4:n" & lastUsedRow()).Select()
                ThisAddIn.AppExcel.ActiveSheet.Range("$B$4:$n$" & lastUsedRow("Automated Checks")).RemoveDuplicates(Columns:=oColumns, Header:=header)
            End If

            searchSession.WrapUp()
        End If
    End Sub

    Private Shared Sub automatedSearch(searchURL As String)
        Dim post As Post
        Dim wc As New System.Net.WebClient
        'basically this one filters out the unneeded posts and then requests that we only write the good ones
        'this one also scrapes the entire cl search response page and loops through each posting

        Dim resultURL As String : resultURL = ""                     'URL of search result page
        Dim updatedSearchURL As String : updatedSearchURL = ""       'search result URL... this gets iterated through pagination
        Dim startPos As Integer : startPos = 1 'this is the start position of the search in the search results and will be incremented to find different results
        Dim searchPage As String : searchPage = ""                   'HTML of whole search result page

        'set or reset global variables
        searchPage = wc.DownloadString(searchURL)
        'start scraping
        Do While InStr(startPos, searchPage, ThisAddIn.ResultHook) > 0

            'if max results not yet exceeded
            If searchSession.ResultCount < ThisAddIn.MaxResults Then
                searchSession.ResultCount += 1
                ThisAddIn.AppExcel.StatusBar = "Currently on result number " & searchSession.ResultCount & " (old, in trash, or otherwise skipped: " & searchSession.SkippedResultCount & ") (keepers: " & searchSession.KeeperCount & ") (maybes: " & searchSession.NegCount & ") [MULTIPOSTS: " & searchSession.MultiCount & "]"

                'if not a nearby result
                If Not getLinkFromCLSearchResults(searchPage, startPos) Like "*http*" Then 'this prevents it from showing "nearby results"

                    'its time to learn a bit about the post
                    post = New Post(ThisAddIn.TldUrl & getLinkFromCLSearchResults(searchPage, startPos))

                    'if not already written to the automated checks page
                    If Not post.alreadyChecked() Then

                        'if not already in trash
                        If Not wasCategorized(post, , , "Trash") And Not wasCategorized(post, , , "Maybes") And Not wasCategorized(post, , , "Keepers") Then

                            'if not a magazine post
                            If Not LCase(post.Title) Like "*magazine*" Then

                                'POST DATE PREFERENCES
                                If post.UpdateDate = "" Or post.UpdateDate = "-" Then post.UpdateDate = post.PostDate
                                Select Case ThisAddIn.PostTimingPref
                                    Case "timingPostedToday", "timingUpdatedToday"  'if only today's new posts requested
                                        If FormatDateTime(post.PostDate, vbShortDate) = FormatDateTime(Now(), vbShortDate) Then
                                            searchSession.WriteSearchResult(post)
                                        Else
                                            searchSession.TooOldPosts += 1
                                            searchSession.SkippedResultCount += 1
                                        End If
                                    Case "timingUpdated7Days" 'if posts updated in last 7 days
                                        If ThisAddIn.AppExcel.WorksheetFunction.Days360(FormatDateTime(post.UpdateDate, vbShortDate), FormatDateTime(DateTime.Now.AddDays(-7), vbShortDate)) <= 0 Then
                                            searchSession.WriteSearchResult(post)
                                        Else
                                            searchSession.TooOldPosts += 1
                                            searchSession.SkippedResultCount += 1
                                        End If
                                    Case "timingUpdated14Days" 'if post updated in last 14 days
                                        If ThisAddIn.AppExcel.WorksheetFunction.Days360(FormatDateTime(post.UpdateDate, vbShortDate), FormatDateTime(DateTime.Now.AddDays(-14), vbShortDate)) <= 0 Then
                                            searchSession.WriteSearchResult(post)
                                        Else
                                            searchSession.TooOldPosts += 1
                                            searchSession.SkippedResultCount += 1
                                        End If
                                    Case Else
                                        searchSession.WriteSearchResult(post)
                                End Select
                            Else
                                searchSession.MagCount += 1
                                searchSession.SkippedResultCount += 1
                            End If

                        Else 'was categorized result
                            searchSession.WasAlreadyCategorized += 1
                            searchSession.SkippedResultCount += 1
                        End If

                    End If 'was it already written?

                Else 'a "nearby result" that was also skipped
                    searchSession.OutOfTown += 1
                    searchSession.SkippedResultCount += 1
                End If

                'do the pagination
                startPos = InStr(startPos, searchPage, ThisAddIn.ResultHook) + 100
                If searchSession.ResultCount Mod 100 = 0 Then
                    If InStr(1, searchURL, "&") = 0 Then
                        updatedSearchURL = searchURL & "?s=" & searchSession.ResultCount
                    Else
                        updatedSearchURL = searchURL & "&s=" & searchSession.ResultCount
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
        searchSession.WrapUp()
        wc = Nothing
        ThisAddIn.Proceed = False
    End Sub

    Private Shared Sub prepareWBForAutomatedSearch()
        unFilterTrash()
        If lastUsedRow("Automated Checks") > 3 Then
            Select Case MsgBox("Should we empty the 'Automated Checks' page before starting the search?", vbYesNoCancel, ThisAddIn.Title)
                Case vbYes
                    BuildWSAutomatedChecks.BuildWSAutomatedChecks()
                Case vbNo

                Case Else
                    ThisAddIn.AppExcel.ScreenUpdating = True
                    ThisAddIn.AppExcel.StatusBar = False
                    ThisAddIn.Proceed = False
            End Select
        End If
    End Sub

End Class
