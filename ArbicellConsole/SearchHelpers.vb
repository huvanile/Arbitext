Imports ArbitextClassLibrary.Globals
Imports ArbitextClassLibrary
Imports ArbitextClassLibrary.StringHelpers
Imports ArbitextClassLibrary.CraigslistHelpers
Imports ArbitextClassLibrary.RSSHelpers
Imports System.IO

Public Class SearchHelpers
    Private Shared _checkedPostsAndPhones As List(Of Post)      'list of posts checked with all of the books included

    Public Shared Sub allQuerySearch()
        _checkedPostsAndPhones = New List(Of Post)
        Console.ForegroundColor = ConsoleColor.DarkGreen
        Console.WriteLine("Starting phones search...")
        Console.ResetColor()
        'oneQuerySearch(TldUrl & "/search/sss?query=cell%20phone")
        'oneQuerySearch(TldUrl & "/search/sss?query=cellphone")
        'oneQuerySearch(TldUrl & "/search/sss?query=smart%20phone")
        'oneQuerySearch(TldUrl & "/search/sss?query=smartphone")
        'oneQuerySearch(TldUrl & "/search/sss?query=mobile%20phone")
        'oneQuerySearch(TldUrl & "/search/sss?query=mobilephone")
        oneQuerySearch(TldUrl & "/search/moa")
        Console.WriteLine("Done!")
    End Sub


    Private Shared Sub oneQuerySearch(searchURL As String)
        Dim postAndPhones As Post                                    'this is a fully populated post object, including the books
        Dim wc As New Net.WebClient
        Dim resultURL As String : resultURL = ""                     'URL of search result page
        Dim updatedSearchURL As String : updatedSearchURL = ""       'search result URL... this gets iterated through pagination
        Dim startPos As Integer : startPos = 1 'this is the start position of the search in the search results and will be incremented to find different results
        Dim searchPage As String : searchPage = ""                   'HTML of whole search result page

        'set or reset global variables
        searchPage = wc.DownloadString(searchURL)
        If Not searchPage Like "*Nothing found for that search*" Then
            Do While InStr(startPos, searchPage, _resultHook) > 0
                Console.WriteLine("On Result Number " & _checkedPostsAndPhones.Count + 1)
                Console.WriteLine("Cursory examination of post " & getLinkFromCLSearchResults(searchPage, startPos))

                'if not a nearby result
                If Not getLinkFromCLSearchResults(searchPage, startPos) Like "*http*" And Not getLinkFromCLSearchResults(searchPage, startPos) Like "*.org*" Then 'this prevents it from showing "nearby results"
                    Console.WriteLine("Learning more about post " & getLinkFromCLSearchResults(searchPage, startPos))

                    postAndPhones = New Post(True, TldUrl & getLinkFromCLSearchResults(searchPage, startPos))

                    'make sure it wasn't aleady checked this session
                    Dim match As Boolean = False
                    For Each p In _checkedPostsAndPhones
                        If p.Equals(postAndPhones) Then
                            match = True
                            Exit For
                        End If
                    Next
                    Console.ForegroundColor = ConsoleColor.DarkGray
                    Console.WriteLine("- Posts:  Checked checked:  " & _checkedPostsAndPhones.Count)
                    Console.WriteLine("- Posts:  Unparseable:  " & _checkedPostsAndPhones.Where(Function(x) x.IsParsable = False).Count)
                    Console.WriteLine("- Posts:  Parseable: " & _checkedPostsAndPhones.Where(Function(x) x.IsParsable = True).Count)
                    Console.WriteLine("- Phones:  Winners: " & _checkedPostsAndPhones.SelectMany(Function(x) x.Phones).Where(Function(y) y.IsWinner = True).Count)
                    Console.WriteLine("- Phones:  Maybes: " & _checkedPostsAndPhones.SelectMany(Function(x) x.Phones).Where(Function(y) y.IsMaybe = True).Count)
                    Console.WriteLine("- Phones:  HVSB: " & _checkedPostsAndPhones.SelectMany(Function(x) x.Phones).Where(Function(y) y.IsHVSP = True).Count)
                    Console.WriteLine("- Phones:  HVOBO: " & _checkedPostsAndPhones.SelectMany(Function(x) x.Phones).Where(Function(y) y.IsHVOBO = True).Count)
                    Console.WriteLine("- Phones:  Trash: " & _checkedPostsAndPhones.SelectMany(Function(x) x.Phones).Where(Function(y) y.IsTrash = True).Count)
                    Console.ResetColor()

                    If Not match Then
                        If postAndPhones.IsParsable Then
                            WriteSearchResult(postAndPhones)
                            _checkedPostsAndPhones.Add(postAndPhones)
                        Else
                            Console.ForegroundColor = ConsoleColor.DarkRed
                            Console.WriteLine("Unparseable post: " & postAndPhones.URL)
                            Console.ResetColor()
                            postAndPhones.IsParsable = False
                        End If
                    Else
                        postAndPhones.IsParsable = False
                    End If
                    _checkedPostsAndPhones.Add(postAndPhones)
                End If

                'do the pagination
                startPos = InStr(startPos, searchPage, _resultHook) + 100
                If _checkedPostsAndPhones.Count Mod 100 = 0 Then
                    If InStr(1, searchURL, "&") = 0 Then
                        updatedSearchURL = searchURL & "?s=" & _checkedPostsAndPhones.Count
                    Else
                        updatedSearchURL = searchURL & "&s=" & _checkedPostsAndPhones.Count
                    End If
                    searchPage = wc.DownloadString(updatedSearchURL)
                    startPos = 1 'to start searching at the top of the newly loaded search page
                End If

            Loop
        Else
            Console.ForegroundColor = ConsoleColor.DarkRed
            Console.WriteLine("Nothing found for the search:" & vbCrLf & vbCrLf & searchURL)
            Console.ResetColor()
        End If
        wc = Nothing
        Exit Sub
    End Sub

    Private Shared Sub WriteSearchResult(post As Post)
        Dim resultType As String
        Dim rssFeed As RSSFeed
        Dim theCity As String = StrConv(City, VbStrConv.ProperCase)
        Dim desc As String = ""
        Dim title As String = ""
        Dim outfile As String = ""
        Dim proceed As Boolean
        For Each p As Phone In post.Phones
            proceed = True
            Console.WriteLine("Querying ThePriceGeek about Phone")
            p.GetDataFromThePriceGeek()
            Console.WriteLine("Writing Phone Info")
            If post.IsParsable AndAlso p.IsParsable Then
                If p.IsWinner() Then
                    Console.ForegroundColor = ConsoleColor.Green
                    Console.Write("WINNER phone WINNER phone WINNER!  ")
                    Console.ResetColor()
                    resultType = "Winners"
                    desc = "Profitable phone deals (winners) in the " & theCity & " area."
                    title = theCity & " Winners- Phones"
                    outfile = theCity & " Winners- Phones.xml"
                ElseIf p.IsMaybe() Then
                    Console.ForegroundColor = ConsoleColor.Cyan
                    Console.Write("MAYBE phone MAYBE phone MAYBE!  ")
                    Console.ResetColor()
                    resultType = "Maybes"
                    desc = "Potentially profitable phone deals (maybes) in the " & theCity & " area."
                    title = theCity & " Maybes- Phones"
                    outfile = theCity & " Maybes- Phones.xml"
                ElseIf p.IsHVSP() Then
                    Console.ForegroundColor = ConsoleColor.Magenta
                    Console.Write("HIGH VALUE STALE PHONE!  ")
                    Console.ResetColor()
                    resultType = "HVSPs"
                    desc = "High value stale phones in the " & theCity & " area.  These phones can be sold for a profit, but only if the seller (who hasn't been successful selling them at the current asking price) will come down on the price a bit."
                    title = theCity & " Stale Phones of Value"
                    outfile = theCity & " High Value Stale Phones.xml"
                ElseIf p.IsHVOBO() Then
                    Console.ForegroundColor = ConsoleColor.Magenta
                    Console.Write("VALUABLE 'OR BEST OFFER' PHONE!")
                    Console.ResetColor()
                    resultType = "HVOBOs"
                    desc = "High value 'or best offer' phones in the " & theCity & " area.  These phones can be sold for a profit, but only if the seller (who indicated that they'd consider the best offer) will come down on the price a bit."
                    title = theCity & " Valuable 'Or Best Offer' Phones"
                    outfile = theCity & " Valuable Or Best Offer Phones.xml"
                Else
                    Console.ForegroundColor = ConsoleColor.DarkRed
                    Console.WriteLine("Trash phone lead found in: " & post.URL)
                    Console.ResetColor()
                    proceed = False
                End If
            Else
                Console.ForegroundColor = ConsoleColor.DarkRed
                Console.WriteLine("Unparseable post or phone:  " & post.URL)
                Console.ResetColor()
                proceed = False
            End If

            If proceed Then
                Debug.Print(p.ID, " ", resultType, post.URL)
                If Not AlreadyInRSSFeed(p.ID, resultType, Sftp, SftpDirectory, City, SftpURL, "phone") Then
                    Console.ForegroundColor = ConsoleColor.Yellow
                    Console.WriteLine(" NOT YET IN THE FEEDS- ADDING!")
                    Console.ResetColor()

                    If Not FeedAlreadyExists("phone", resultType, Sftp, SftpDirectory, City) Then
                        rssFeed = New RSSFeed(title, wwwRoot & "showfeed.php?feed=" & replaceSpacesWithTwenty(Path.GetFileName(outfile)), desc, resultType, outfile)
                    Else
                        rssFeed = New RSSFeed(wwwRoot & "leads/" & replaceSpacesWithTwenty(outfile))

                    End If


                    Dim postLink As String = "https://href.li/?" & post.URL 'arbitext:postLink 
                    Dim resultURL As String = wwwRoot & "showitem.php?item=" & replaceSpacesWithTwenty(Path.GetFileName(rssFeed.FileName)) & "|" & p.ID
                    Dim theDesc As String = getDesc(resultType, post.City, p.AskingPrice, p.Profit, p.Median)

                    WriteRSSItem(rssFeed.Document, p.SearchTerm, resultURL, CDate(post.UpdateDate), theDesc, p.ID, post.URL, post.Title, post.City,
                                 p.AskingPrice, p.TpgURL, p.Median, p.Mean, p.Profit, p.ProfitPercentage, p.isOBO, post.Image)

                    PushUpdatedXML(rssFeed, Sftp)
                Else
                    Console.ForegroundColor = ConsoleColor.DarkYellow
                    Console.WriteLine("  Already in the feeds, no need to add")
                    Console.ResetColor()
                End If
            End If 'proceed check
        Next p

    End Sub

End Class
