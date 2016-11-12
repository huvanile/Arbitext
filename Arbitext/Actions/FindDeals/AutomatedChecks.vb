Imports Arbitext.RegistryHelpers
Imports Arbitext.ExcelHelpers
Imports Arbitext.ArbitextHelpers
Imports Arbitext.PushbulletHelpers
Imports Arbitext.BookscouterHelpers
Imports Arbitext.CraigslistHelpers
Imports Arbitext.EmailHelpers
Imports Arbitext.SoundHelpers
Imports Arbitext.StringHelpers
Imports Microsoft.Office.Interop.Excel

Public Class AutomatedChecks

    Shared searchURL As String

    '======== COUNTS ============
    Shared atRow As String                 'count of search results that were auto-trashed in this session
    Shared tooOldPosts As String           'count of posts that were too old to show based on preferences
    Shared skippedResultCount As Integer   'count of skipped search results
    Shared resultCount As Integer          'count of search results
    Shared wasAlreadyCategorized As Integer           'count of results already in trash... a subset of skipped results
    Shared outOfTown As Integer            'count of "nearby" (but not really results)... a subset of skipped results
    Shared magCount As Integer             'count of magazines in search results
    Shared keeperCount As Integer          'count of potential winners!
    Shared multiCount As Integer           'count of multi-posts
    Shared dunnoCount As Integer           'count of ones we couldn't figure out
    Shared negCount As Integer             'count of ones to maybe negoitate for

    '======== OTHER ====================
    Shared r As Integer
    Shared wasParsed As Boolean

    Public Shared Sub AutomatedChecks()
        ThisAddIn.Proceed = True
        ThisAddIn.AppExcel.ScreenUpdating = True
        automatedChecksPageCheck()
        prepareWBForAutomatedSearch()
        If ThisAddIn.Proceed Then

            'set or reset global variables
            ThisAddIn.TldUrl = "http://" & ThisAddIn.City & ".craigslist.org"
            keeperCount = 0
            negCount = 0
            resultCount = 0
            wasAlreadyCategorized = 0
            outOfTown = 0
            If ThisAddIn.MaxResults = 0 Then ThisAddIn.MaxResults = 1000000 'just some crazy  high number that we'll never hit if teh user set unlimited results
            atRow = 0
            skippedResultCount = 0
            tooOldPosts = 0
            magCount = 0
            multiCount = 0
            dunnoCount = 0

            'isbn search
            searchURL = ThisAddIn.TldUrl & "/search/sss?query=isbn&sort=rel"
            If ThisAddIn.PostTimingPref Like "*today*" Then searchURL = searchURL & "&postedToday=1"
            automatedSearch()

            'textbook search
            If ThisAddIn.Proceed Then
                searchURL = ThisAddIn.TldUrl & "/search/sss?query=textbook&sort=rel"
                If ThisAddIn.PostTimingPref Like "*today*" Then searchURL = searchURL & "&postedToday=1"
                automatedSearch()
            End If

            'books search
            If ThisAddIn.Proceed Then
                searchURL = ThisAddIn.TldUrl & ".craigslist.org/search/ndf/bks"
                If ThisAddIn.PostTimingPref Like "*today*" Then searchURL = searchURL & "&postedToday=1"
                automatedSearch()
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

            wrapUp()
        End If
    End Sub

    Private Shared Sub automatedSearch()
        Dim wc As New System.Net.WebClient
        'basically this one filters out the unneeded posts and then requests that we only write the good ones
        'this one also scrapes the entire cl search response page and loops through each posting

        Dim resultURL As String : resultURL = ""                     'URL of search result page
        Dim updatedSearchURL As String : updatedSearchURL = ""       'search result URL... this gets iterated through pagination
        Dim startPos As Integer : startPos = 1 'this is the start position of the search in the search results and will be incremented to find different results
        Dim searchPage As String : searchPage = ""                   'HTML of whole search result page

        'set or reset global variables
        ThisAddIn.PostURL = ""
        searchPage = wc.DownloadString(searchURL)
        'start scraping
        Do While InStr(startPos, searchPage, ThisAddIn.ResultHook) > 0

            'grab url of next post
            ThisAddIn.PostURL = ThisAddIn.TldUrl & getLinkFromCLSearchResults(searchPage, startPos)

            'if max results not yet exceeded
            If resultCount < ThisAddIn.MaxResults Then
                resultCount = resultCount + 1
                ThisAddIn.AppExcel.StatusBar = "Currently on result number " & resultCount & " (old, in trash, or otherwise skipped: " & skippedResultCount & ") (keepers: " & keeperCount & ") (maybes: " & negCount & ") [MULTIPOSTS: " & multiCount & "]"

                'if not a nearby result
                If Not getLinkFromCLSearchResults(searchPage, startPos) Like "*http*" Then 'this prevents it from showing "nearby results"

                    'its time to learn a bit about the post
                    learnAboutPost()

                    'if not already written to the automated checks page
                    If Not alreadyChecked() Then

                        'if not already in trash
                        If Not wasCategorized(, , "Trash") And Not wasCategorized(, , "Maybes") And Not wasCategorized(, , "Keepers") Then

                            'if not a magazine post
                            If Not LCase(ThisAddIn.PostTitle) Like "*magazine*" Then

                                'POST DATE PREFERENCES
                                If ThisAddIn.PostUpdateDate = "" Or ThisAddIn.PostUpdateDate = "-" Then ThisAddIn.PostUpdateDate = ThisAddIn.PostDate
                                Select Case ThisAddIn.PostTimingPref
                                    Case "timingPostedToday", "timingUpdatedToday"  'if only today's new posts requested
                                        If FormatDateTime(ThisAddIn.PostDate, vbShortDate) = FormatDateTime(Now(), vbShortDate) Then
                                            writeSearchResult()
                                        Else
                                            tooOldPosts = tooOldPosts + 1
                                            skippedResultCount = skippedResultCount + 1
                                        End If
                                    Case "timingUpdated7Days" 'if posts updated in last 7 days
                                        If ThisAddIn.AppExcel.WorksheetFunction.Days360(FormatDateTime(ThisAddIn.PostUpdateDate, vbShortDate), FormatDateTime(DateTime.Now.AddDays(-7), vbShortDate)) <= 0 Then
                                            writeSearchResult()
                                        Else
                                            tooOldPosts = tooOldPosts + 1
                                            skippedResultCount = skippedResultCount + 1
                                        End If
                                    Case "timingUpdated14Days" 'if post updated in last 14 days
                                        If ThisAddIn.AppExcel.WorksheetFunction.Days360(FormatDateTime(ThisAddIn.PostUpdateDate, vbShortDate), FormatDateTime(DateTime.Now.AddDays(-14), vbShortDate)) <= 0 Then
                                            writeSearchResult()
                                        Else
                                            tooOldPosts = tooOldPosts + 1
                                            skippedResultCount = skippedResultCount + 1
                                        End If
                                    Case Else
                                        writeSearchResult()
                                End Select
                            Else
                                magCount = magCount + 1
                                skippedResultCount = skippedResultCount + 1
                            End If

                        Else 'was categorized result
                            wasAlreadyCategorized = wasAlreadyCategorized + 1
                            skippedResultCount = skippedResultCount + 1
                        End If

                    End If 'was it already written?

                Else 'a "nearby result" that was also skipped
                    outOfTown = outOfTown + 1
                    skippedResultCount = skippedResultCount + 1
                End If

                'do the pagination
                startPos = InStr(startPos, searchPage, ThisAddIn.ResultHook) + 100
                If resultCount Mod 100 = 0 Then
                    If InStr(1, searchURL, "&") = 0 Then
                        updatedSearchURL = searchURL & "?s=" & resultCount
                    Else
                        updatedSearchURL = searchURL & "&s=" & resultCount
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
        wrapUp()
        wc = Nothing
        ThisAddIn.Proceed = False
    End Sub

    Private Shared Sub prepareWBForAutomatedSearch()
        unFilterTrash()

        Select Case MsgBox("Should we empty the 'Automated Checks' page before starting the search?", vbYesNoCancel, ThisAddIn.Title)
            Case vbYes
                BuildWSAutomatedChecks.BuildWSAutomatedChecks()
            Case vbNo

            Case Else
                ThisAddIn.AppExcel.ScreenUpdating = True
                ThisAddIn.AppExcel.StatusBar = False
                ThisAddIn.Proceed = False
        End Select

    End Sub

    Private Shared Sub writeSearchResult()
        r = lastUsedRow() + 1

        If isMulti(ThisAddIn.PostTitle) Or isMulti(ThisAddIn.PostBody) Then
            With ThisAddIn.AppExcel
                multiCount = multiCount + 1

                If bookCount(ThisAddIn.PostBody) = 0 Then
                    'doesn't look like we can parse it because bookcount didn't work
                    .Range("b" & r & ":n" & r).Interior.ColorIndex = 20
                    writeTheEasyStuff(r)
                    writeDummyValues(r)
                    .Range("g" & r).Value2 = ThisAddIn.PostAskingPrice
                    .Range("f" & r).Value2 = "(multipost, unknown book count)"
                    dunnoCount = dunnoCount + 1
                Else
                    'bookcount worked, let's try to parse it
                    doAMultiAutoCheck("2br", r)
                    If Not wasParsed Then doAMultiAutoCheck("1br", r)
                    If Not wasParsed Then doAMultiAutoCheck("p", r)

                    If (.Range("f" & r).Value2 Like "(" Or .Range("f" & r).Value2 = "") And Not wasParsed Then
                        'doesn't look like we can parse it because something went wrong
                        .Range("b" & r & ":n" & r).Interior.ColorIndex = 20
                        writeTheEasyStuff(r)
                        writeDummyValues(r)
                        .Range("g" & r).Value = ThisAddIn.PostAskingPrice
                        .Range("f" & r).Value2 = "(multipost, unparsable)"
                        dunnoCount = dunnoCount + 1
                    End If

                End If
            End With
        Else
            With ThisAddIn.AppExcel
                writeTheEasyStuff(r)
                writeDummyValues(r)
                .Range("g" & r).Value2 = ThisAddIn.PostAskingPrice 'asking price
                If ThisAddIn.PostISBN = "(unknown)" Or ThisAddIn.PostISBN = "" Or ThisAddIn.PostISBN = "(not listed in ad)" Then
                    ThisAddIn.PostISBN = getISBN(ThisAddIn.PostTitle)
                    If ThisAddIn.PostISBN = "(unknown)" Or ThisAddIn.PostISBN = "" Or ThisAddIn.PostISBN = "(not listed in ad)" Then ThisAddIn.PostISBN = getISBN(ThisAddIn.PostBody)
                End If
                .Range("f" & r).Value2 = ThisAddIn.PostISBN 'isbn
                If Not .Range("F" & r).Value2 Like "*(*" Then
                    .ActiveSheet.Hyperlinks.Add(anchor:= .Range("f" & r), Address:="https://bookscouter.com/prices.php?isbn=" & ThisAddIn.PostISBN & "&all", TextToDisplay:="'" & .Range("f" & r).Value)
                End If

                'write the numbers if we're good on the result so far
                If Not ThisAddIn.PostISBN Like "*(*" And ThisAddIn.PostISBN <> "" Then
                    askBSAboutBook()
                    .Range("h" & r).Value2 = ThisAddIn.PostSellingPrice
                    .ActiveSheet.Hyperlinks.Add(anchor:= .Range("h" & r), Address:=ThisAddIn.BsBuybackLink, TextToDisplay:="'" & .Range("h" & r).Value)

                    'work out the profit if the online selling price came back
                    If Not .Range("h" & r).Value2 Like "*(*" _
                And Not .Range("g" & r).Value2 Like "*(*" Then
                        writeProfitCalcs(r)
                    End If

                Else
                    dunnoCount = dunnoCount + 1
                End If 'paren check on columns F, G, and H

                doFlagChecks(ThisAddIn.PostBody, r)
                doFlagChecks(ThisAddIn.PostTitle, r)
                autoCategorizeRowIfNeeded(r)
            End With
        End If 'multi handling
    End Sub

    Private Shared Sub autoCategorizeRowIfNeeded(ByVal theR As Integer)
        If ThisAddIn.AutoTrashOK Then 'checkbox 8 "auto-trash"
            With ThisAddIn.AppExcel
                If IsNumeric(.Range("i" & theR).Value2) Then

                    'first the trash
                    If .Range("n" & theR).Interior.ColorIndex <> 6 _
                    And .Range("i" & theR).Value <= 0 Then
                        categorizeRow(theR, "Trash")
                        atRow = atRow + 1
                        skippedResultCount = skippedResultCount + 1
                    End If

                    'auto trash those ebooks if that pref is set
                    If ThisAddIn.AutoTrashEBooksOK _
                    AndAlso Not IsDBNull(.Rows(theR).Font.ColorIndex) _
                    AndAlso .Rows(theR).Font.ColorIndex = 53 Then
                        categorizeRow(theR, "Trash")
                        atRow = atRow + 1
                        skippedResultCount = skippedResultCount + 1
                    End If

                    'auto trash those loose leaf editions if that pref is set
                    If ThisAddIn.AutoTrashBindersOK _
                    AndAlso Not IsDBNull(.Rows(theR).Font.ColorIndex) _
                    AndAlso .Rows(theR).Font.ColorIndex = 3 Then
                        categorizeRow(theR, "Trash")
                        atRow = atRow + 1
                        skippedResultCount = skippedResultCount + 1
                    End If

                    'then the maybes
                    If .Range("n" & theR).Interior.ColorIndex = 6 Then
                        categorizeRow(theR, "Maybes")
                        atRow = atRow + 1
                        skippedResultCount = skippedResultCount + 1
                    End If

                    'then the winners
                    If .Range("n" & theR).Interior.ColorIndex = 35 Then
                        categorizeRow(theR, "Keepers")
                        atRow = atRow + 1
                        skippedResultCount = skippedResultCount + 1
                    End If

                End If
            End With
        End If
    End Sub

    Private Shared Sub writeProfitCalcs(ByVal theR As Integer)
        With ThisAddIn.AppExcel
            If Not IsNumeric(.Range("g" & theR).Value2) Then
                .Range("i" & theR).Value2 = "?" 'profit
                .Range("j" & theR).Value2 = "?" 'profit margin
                .Range("k" & theR).Value2 = "?" 'min asking price for desired profit margin
                .Range("n" & theR).Value2 = "?" 'price delta
            ElseIf .Range("h" & theR).Value2 = 0 Then
                .Range("i" & theR).Value2 = 0 'profit
                .Range("j" & theR).Value2 = 0 'profit margin
                .Range("k" & theR).Value2 = 0 'min asking price for desired profit margin
                .Range("n" & theR).Value2 = 0 'price delta
            Else
                .Range("i" & theR).FormulaLocal = "=h" & theR & "-g" & theR 'profit
                .Range("j" & theR).FormulaLocal = "=i" & theR & "/g" & theR 'profit margin
                .Range("k" & theR).FormulaLocal = "=if(round(h" & theR & "-" & ThisAddIn.MinTolerableProfit & ",0)>0,round(h" & theR & "-" & ThisAddIn.MinTolerableProfit & ",0),0)" 'min asking price for desired profit margin
                .Range("n" & theR).FormulaLocal = "=IFERROR(1-(K" & theR & "/G" & theR & "),""-"")" 'delta b/w my price and theirs
            End If
        End With
    End Sub

    Private Shared Sub doAMultiAutoCheck(checkMethod As String, ByVal theR As Integer)
        Dim i As Integer        'book count
        Dim tmpISBN As String : tmpISBN = ""
        Dim tmpAskingPrice As String : tmpAskingPrice = ""
        Dim s As Long           'splitholder start position, increments as books are parsed
        Dim t As Integer         'to loop through the splitholder <BR> results
        Dim splitholder() As String
        Dim tmpPostBody As String
        s = 0
        wasParsed = False

        For i = 1 To bookCount(ThisAddIn.PostBody)
            With ThisAddIn.AppExcel
                .StatusBar = "Currently on result number " & resultCount & " [BOOK " & i & "] (old, in trash, or otherwise skipped: " & skippedResultCount & ") (keepers: " & keeperCount & ") (maybes: " & negCount & ") [MULTIPOSTS: " & multiCount & "]"

                Select Case checkMethod
                    Case "1br"
                        splitholder = Split(ThisAddIn.PostBody, "<br>")
                    Case "p"
                        splitholder = Split(ThisAddIn.PostBody, "<p>")
                    Case "2br"
                        tmpPostBody = Replace(ThisAddIn.PostBody, Chr(10), "")
                        tmpPostBody = Replace(tmpPostBody, Chr(13), "")
                        splitholder = Split(tmpPostBody, "<br><br>")
                End Select

                For t = s To UBound(splitholder)
                    tmpISBN = getISBN(splitholder(t))
                    tmpAskingPrice = getAskingPrice(splitholder(t))
                    If Not tmpISBN Like "(*" And Not tmpAskingPrice Like "*(*" Then 'match!
                        wasParsed = True
                        theR = lastUsedRow("Automated Checks") + 1

                        If Not alreadyChecked(tmpISBN, tmpAskingPrice) Then

                            If Not wasCategorized(tmpISBN, tmpAskingPrice) Then

                                'write results
                                .Range("b" & theR & ":n" & theR).Interior.ColorIndex = 20
                                writeTheEasyStuff(theR)
                                .Range("f" & theR).Value2 = tmpISBN 'isbn
                                If Not tmpISBN Like "*(*" And tmpISBN <> "" Then
                                    .ActiveSheet.Hyperlinks.Add(anchor:= .Range("f" & theR), Address:="https://bookscouter.com/prices.php?isbn=" & tmpISBN & "&all", TextToDisplay:="'" & .Range("f" & theR).Value)
                                    askBSAboutBook(tmpISBN)
                                End If
                                .Range("g" & theR).Value2 = tmpAskingPrice 'asking price
                                If .Range("h" & theR).Value2 = "" Then .Range("h" & theR).Value2 = ThisAddIn.PostSellingPrice 'bookscouter selling price 'don't hit bookscouter if there's already a sell price present (maybe from a prior pass)
                                .ActiveSheet.Hyperlinks.Add(anchor:= .Range("h" & theR), Address:=ThisAddIn.BsBuybackLink & .Range("h" & theR).Value)
                                s = t + 1
                                writeProfitCalcs(theR)
                                doFlagChecks(splitholder(t), theR)
                                autoCategorizeRowIfNeeded(theR)
                                Exit For

                            Else
                                s = t + 1
                                Exit For
                            End If 'was trashed check

                        Else
                            s = t + 1
                            Exit For
                        End If 'already checked check

                    End If
                Next t

            End With
        Next i

    End Sub

    Private Shared Sub writeTheEasyStuff(ByVal theR As Integer)
        With ThisAddIn.AppExcel
            thinInnerBorder(.Range("b" & theR & ":n" & theR))
            .Range("b" & theR).Value2 = theR - 3
            .Range("c" & theR).Value2 = ThisAddIn.PostDate 'date posted
            .Range("d" & theR).Value2 = ThisAddIn.PostURL 'post url
            .ActiveSheet.Hyperlinks.Add(anchor:= .Range("d" & theR), Address:= .Range("d" & theR).Value, TextToDisplay:= .Range("d" & theR).Value)
            .Range("m" & theR).Value2 = ThisAddIn.PostUpdateDate
            .Range("l" & theR).Value2 = ThisAddIn.PostCity
            .Range("e" & theR).Value2 = ThisAddIn.PostTitle
            If ThisAddIn.PostTitle = "" Or ThisAddIn.PostDate = "?" Then
                'every once in a while the python parsing doesnt work completely... this tries to accomodate
                learnAboutPost(ThisAddIn.PostURL)
            End If
            .ActiveSheet.Hyperlinks.Add(anchor:= .Range("e" & theR), Address:="http://www.amazon.com/s/ref=nb_sb_noss?url=search-alias%3Daps&field-keywords=" & replacePlusWithSpace(ThisAddIn.PostTitle), TextToDisplay:= .Range("e" & theR).Value)
        End With
    End Sub

    Private Shared Sub writeDummyValues(ByVal theR As Integer)
        'write more dummy values so we can over-write them later if needed
        With ThisAddIn.AppExcel
            .Range("h" & theR).Value2 = "?" 'online selling price
            .Range("i" & theR).Value2 = "?" 'profit
            .Range("j" & theR).Value2 = "?" 'profit margin
            .Range("k" & theR).Value2 = "?" 'min asking price for desired profit margin
            .Range("n" & theR).Value2 = "?" 'price delta
        End With
    End Sub

    Private Shared Sub doFlagChecks(ByVal theStr As String, ByVal theR As Integer)
        Dim winFlag As Boolean : winFlag = False
        With ThisAddIn.AppExcel
            'flag the a la carte results
            If aLaCarte(theStr) Then
                .Rows(theR).Font.ColorIndex = 3
                ThisAddIn.ALaCarteFlag = True
            End If

            'flag "or best offer"
            If isOBO(theStr) Then
                .Rows(theR).Font.Bold = True
                ThisAddIn.OboFlag = True
            End If

            'flag custom or international editions
            If isWeirdEdition(theStr) Then
                .Rows(theR).Font.ColorIndex = 46
                ThisAddIn.WeirdEditionFlag = True
            End If

            'flag custom or international editions
            If isPDF(theStr) Then
                .Rows(theR).Font.ColorIndex = 53
                ThisAddIn.PdfFlag = True
            End If

            'email alerts for the WINNERS!!!
            If IsNumeric(.Range("g" & theR)) _
        And IsNumeric(.Range("h" & theR)) _
        And IsNumeric(.Range("i" & theR)) Then
                If CInt(.Range("i" & theR).Value2) >= CInt(ThisAddIn.MinTolerableProfit) Then

                    Select Case .Rows(theR).Font.ColorIndex
                        Case 3 'loose leaf
                            If Not ThisAddIn.AutoTrashBindersOK Then 'its not a winner if this pref is set and if it's a binder
                                keeperCount = keeperCount + 1
                                If ThisAddIn.OnWinnersOK Then
                                    sendPushbulletNote(ThisAddIn.Title, "Winner found!")
                                    sendSilentNotification(emailBodyString(theR, "winner"), "Definite Textbook Lead: " & ThisAddIn.PostTitle)
                                End If
                                PlayWAV("money")
                                winFlag = True
                                .Range("b" & theR & ":n" & theR).Interior.ColorIndex = 35

                            End If
                        Case 53 'ebook
                            If Not ThisAddIn.AutoTrashEBooksOK Then 'its not a winner if this pref is set and if it's an ebook

                                keeperCount = keeperCount + 1
                                If ThisAddIn.OnWinnersOK Then
                                    sendPushbulletNote(ThisAddIn.Title, "Winner found!")
                                    sendSilentNotification(emailBodyString(theR, "winner"), "Definite Textbook Lead: " & ThisAddIn.PostTitle)
                                End If
                                PlayWAV("money")
                                winFlag = True
                                .Range("b" & theR & ":n" & theR).Interior.ColorIndex = 35

                            End If
                        Case Else

                            keeperCount = keeperCount + 1
                            If ThisAddIn.OnWinnersOK Then
                                sendPushbulletNote(ThisAddIn.Title, "Winner found!")
                                sendSilentNotification(emailBodyString(theR, "winner"), "Definite Textbook Lead: " & ThisAddIn.PostTitle)
                            End If
                            PlayWAV("money")
                            winFlag = True
                            .Range("b" & theR & ":n" & theR).Interior.ColorIndex = 35

                    End Select

                ElseIf .Range("i" & theR).Value2 >= 1 Then 'below minimum tolerable profit threshold but still profitable, so it's a maybe

                    Select Case .Rows(theR).Font.ColorIndex
                        Case 3 'loose leaf
                            If Not ThisAddIn.AutoTrashBindersOK Then
                                .Range("b" & theR & ":n" & theR).Interior.ColorIndex = 6
                                negCount = negCount + 1
                                If ThisAddIn.OnMaybesOK Then
                                    sendPushbulletNote(ThisAddIn.Title, "Maybe found!")
                                    sendSilentNotification(emailBodyString(theR, "maybe"), "Possible Textbook Lead: " & ThisAddIn.PostTitle)
                                End If
                            End If
                        Case 53 'ebook
                            If Not ThisAddIn.AutoTrashEBooksOK Then
                                .Range("b" & theR & ":n" & theR).Interior.ColorIndex = 6
                                negCount = negCount + 1
                                If ThisAddIn.OnMaybesOK Then
                                    sendPushbulletNote(ThisAddIn.Title, "Maybe found!")
                                    sendSilentNotification(emailBodyString(theR, "maybe"), "Possible Textbook Lead: " & ThisAddIn.PostTitle)
                                End If
                            End If
                        Case Else
                            .Range("b" & theR & ":n" & theR).Interior.ColorIndex = 6
                            negCount = negCount + 1
                            If ThisAddIn.OnMaybesOK Then
                                sendPushbulletNote(ThisAddIn.Title, "Maybe found!")
                                sendSilentNotification(emailBodyString(theR, "maybe"), "Possible Textbook Lead: " & ThisAddIn.PostTitle)
                            End If
                    End Select

                End If
            End If

            'flag the ones with negotiation potential
            If IsNumeric(.Range("n" & theR).Value2) _
            AndAlso .Range("h" & theR).Value2 <> 0 _
            AndAlso .Range("n" & theR).Value2 <= 0.25 _
            AndAlso Not winFlag Then
                Select Case .Rows(theR).Font.ColorIndex
                    Case 3 'loose leaf
                        If Not ThisAddIn.AutoTrashBindersOK Then
                            .Range("b" & theR & ":n" & theR).Interior.ColorIndex = 6
                            negCount = negCount + 1
                            If ThisAddIn.OnMaybesOK Then
                                sendPushbulletNote(ThisAddIn.Title, "Maybe found!")
                                sendSilentNotification(emailBodyString(theR, "maybe"), "Possible Textbook Lead: " & ThisAddIn.PostTitle)
                            End If
                        End If
                    Case 53 'ebook
                        If Not ThisAddIn.AutoTrashEBooksOK Then
                            .Range("b" & theR & ":n" & theR).Interior.ColorIndex = 6
                            negCount = negCount + 1
                            If ThisAddIn.OnMaybesOK Then
                                sendPushbulletNote(ThisAddIn.Title, "Maybe found!")
                                sendSilentNotification(emailBodyString(theR, "maybe"), "Possible Textbook Lead: " & ThisAddIn.PostTitle)
                            End If
                        End If
                    Case Else
                        .Range("b" & theR & ":n" & theR).Interior.ColorIndex = 6
                        negCount = negCount + 1
                        If ThisAddIn.OnMaybesOK Then
                            sendPushbulletNote(ThisAddIn.Title, "Maybe found!")
                            sendSilentNotification(emailBodyString(theR, "maybe"), "Possible Textbook Lead: " & ThisAddIn.PostTitle)
                        End If
                End Select
            End If
        End With
    End Sub



    Public Shared Function alreadyChecked(Optional theISBN As String = "", Optional theAskingPrice As String = "") As Boolean
        alreadyChecked = False

        With ThisAddIn.AppExcel.Sheets("Automated Checks")

            If isMulti(ThisAddIn.PostTitle) Or isMulti(ThisAddIn.PostBody) Then

                'to check for parsable multiposts
                If Not theISBN = "" And Not theISBN Like "*(*" Then
                    If canFind(theISBN, "Automated Checks") Then
                        If Trim(.Range("g" & .Range(canFind(theISBN, "Automated Checks", , True, False)).Row).Value) = theAskingPrice _
                        And .Range("l" & .Range(canFind(theISBN, "Automated Checks", , True, False)).Row).Value = ThisAddIn.PostCity _
                        And .Range("f" & .Range(canFind(theISBN, "Automated Checks", , True, False)).Row).Value = theISBN Then
                            alreadyChecked = True
                            Exit Function
                        End If
                    End If
                End If

                'to check for unparsable multiposts
                If canFind(ThisAddIn.PostTitle, "Automated Checks") Then
                    If Trim(.Range("g" & .Range(canFind(ThisAddIn.PostTitle, "Automated Checks", , True, False)).Row).Value) = ThisAddIn.PostAskingPrice _
                    And LCase(.Range("f" & .Range(canFind(ThisAddIn.PostTitle, "Automated Checks", , True, False)).Row).Value) Like "*multi*" _
                    And .Range("l" & .Range(canFind(ThisAddIn.PostTitle, "Automated Checks", , True, False)).Row).Value = ThisAddIn.PostCity Then
                        alreadyChecked = True
                    End If
                End If

            Else 'not multi

                If canFind(ThisAddIn.PostTitle, "Automated Checks") Then
                    If Trim(.Range("g" & .Range(canFind(ThisAddIn.PostTitle, "Automated Checks", , True, False)).Row).Value) = ThisAddIn.PostAskingPrice _
                    And .Range("l" & .Range(canFind(ThisAddIn.PostTitle, "Automated Checks", , True, False)).Row).Value = ThisAddIn.PostCity Then
                        alreadyChecked = True
                    End If
                End If

            End If

        End With

    End Function

    Private Shared Function emailBodyString(theRow As Integer, scenario As String) As String
        Dim message As String = ""
        Select Case scenario
            Case "maybe"
                message = "<h2 style='color:orange; text-align:left'>..:: Textbook lead with negotiation potential found! ::..</h2>" & vbNewLine
            Case "winner"
                message = "<h2 style='color:green; text-align:left'>..:: Definite textbook lead found! ::..</h2>" & vbNewLine
        End Select
        message = message & "<hr/>" & vbNewLine
        message = message & "<h3 style='text-decoration: underline;'>Lead Details</h3>" & vbNewLine
        message = message & "<p><b>Post Title:</b>  " & ThisAddIn.PostTitle & "</p>" & vbNewLine
        message = message & "<p><b>Post URL:</b>  " & ThisAddIn.PostURL & "</p>" & vbNewLine
        message = message & "<p><b>Date Posted:</b>  " & ThisAddIn.PostDate & "</p>" & vbNewLine
        message = message & "<p><b>Date Last Updated:</b>  " & ThisAddIn.PostUpdateDate & "</p>" & vbNewLine
        message = message & "<p><b>ISBN:</b>  <a href=""https://bookscouter.com/prices.php?isbn=" & ThisAddIn.PostISBN & """>" & ThisAddIn.PostISBN & "</a></p>" & vbNewLine
        message = message & "<p><b>City:</b>  " & ThisAddIn.PostCity & "</p>" & vbNewLine
        message = message & "<hr/>" & vbNewLine
        message = message & "<h3 style='text-decoration: underline;'>Financial Details</h3>" & vbNewLine
        message = message & "<p><b>Asking Price:</b>  $" & ThisAddIn.PostAskingPrice & "</p>" & vbNewLine
        message = message & "<p><b>Online Selling Price:</b>  $" & ThisAddIn.PostSellingPrice & "</p>" & vbNewLine
        If ThisAddIn.PostSellingPrice <> "" And IsNumeric(ThisAddIn.PostSellingPrice) And IsNumeric(ThisAddIn.PostAskingPrice) Then
            message = message & "<p><b>Profit:</b>  $" & ThisAddIn.AppExcel.Round(ThisAddIn.PostSellingPrice - ThisAddIn.PostAskingPrice, 2) & "</p>" & vbNewLine
            message = message & "<p><b>Profit Margin:</b>  " & 100 * ThisAddIn.AppExcel.Sheets("Automated Checks").Range("j" & theRow).Value & "%</p>" & vbNewLine
            message = message & "<p><b>Asking price I'd need to get minimum desired profit (rounded):</b>  $" & ThisAddIn.AppExcel.Sheets("Automated Checks").Range("k" & theRow).Value & "</p>" & vbNewLine
            message = message & "<p><b>Delta between my minimum price and their asking price:</b>  " & 100 * ThisAddIn.AppExcel.Sheets("Automated Checks").Range("n" & theRow).Value & "%</p>" & vbNewLine
        Else
            message = message & "<p><b>Profit:</b>  It's weird, but couldn't compute profit for this one</p>" & vbNewLine
            message = message & "<p><b>Profit Margin:</b>  It's weird, but couldn't compute profit for this one</p>" & vbNewLine
        End If
        message = message & "<hr/>" & vbNewLine
        message = message & "<h3 style='text-decoration: underline;'>Flags</h3>" & vbNewLine
        If ThisAddIn.PdfFlag Then
            message = message & "<p style='color:red;><b>eBook Flag?</b>  " & ThisAddIn.PdfFlag & "</p>" & vbNewLine
        Else
            message = message & "<p><b>eBook Flag?</b>  " & ThisAddIn.PdfFlag & "</p>" & vbNewLine
        End If
        If ThisAddIn.WeirdEditionFlag Then
            message = message & "<p style='color:red;><b>Weird Edition Flag?</b>  " & ThisAddIn.WeirdEditionFlag & "</p>" & vbNewLine
        Else
            message = message & "<p><b>Weird Edition Flag?</b>  " & ThisAddIn.WeirdEditionFlag & "</p>" & vbNewLine
        End If
        If ThisAddIn.ALaCarteFlag Then
            message = message & "<p style='color:red;><b>A La Carte Edition Flag?</b>  " & ThisAddIn.ALaCarteFlag & "</p>" & vbNewLine
        Else
            message = message & "<p><b>A La Carte Edition Flag?</b>  " & ThisAddIn.ALaCarteFlag & "</p>" & vbNewLine
        End If
        If ThisAddIn.OboFlag Then
            message = message & "<p style='color:GREEN;><b>""Or Best Offer"" Flag?</b>  " & ThisAddIn.OboFlag & "</p>" & vbNewLine
        Else
            message = message & "<p><b>""Or Best Offer"" Flag?</b>  " & ThisAddIn.OboFlag & "</p>" & vbNewLine
        End If
        message = message & "<hr/>" & vbNewLine
        message = message & "<h3 style='text-decoration: underline;'>Body of Craigslist Post</h3>" & vbNewLine
        message = message & ThisAddIn.PostBody & vbNewLine
        message = message & "<hr/>" & vbNewLine
        message = message & "<p><img src='http://replygif.net/i/159.gif'/></p>"
        emailBodyString = message
    End Function

    Private Shared Sub wrapUp()
        With ThisAddIn.AppExcel
            .Columns("g").Style = "Currency" : .Range("g2").NumberFormat = "General"
            .Columns("h").Style = "Currency"
            .Columns("i").Style = "Currency"
            .Columns("j").Style = "Percent"
            .Columns("k").Style = "Currency"
            .Columns("n").Style = "Percent"

            Dim doneMessage As String
            .ScreenUpdating = True
            doneMessage = "Automated check done examining " & resultCount & " posts!" & vbCrLf & vbCrLf &
            "KEEPERS: " & keeperCount & vbCrLf & vbCrLf &
            "TOTAL SKIPPED RESULTS: " & skippedResultCount & vbCrLf &
            " - Auto-trashed results: " & atRow & vbCrLf &
            " - Multi-posts: " & multiCount & vbCrLf &
            " - Magazine posts: " & magCount & vbCrLf &
            " - Results already in trash: " & wasAlreadyCategorized & vbCrLf &
            " - ""Nearby"" (but not really) results: " & outOfTown & vbCrLf &
            " - Too old posts (based on preferences): " & tooOldPosts & vbCrLf &
            " - Results we otherwise couldn't figure out: " & dunnoCount
            .StatusBar = False
            .Goto(.Range("A5"), True)

        End With
    End Sub

End Class
