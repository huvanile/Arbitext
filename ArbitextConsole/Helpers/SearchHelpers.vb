﻿Imports ArbitextClassLibrary.Globals
Imports ArbitextClassLibrary
Imports ArbitextClassLibrary.StringHelpers
Imports ArbitextClassLibrary.CraigslistHelpers
Imports ArbitextClassLibrary.RSSHelpers
Imports System.IO

Public Class SearchHelpers
    Private Shared _checkedPostsNotBooks As List(Of Post)      'list of posts checked before populating the post object with the accompanying books
    Private Shared _checkedPostsAndBooks As List(Of Post)      'list of posts checked with all of the books included

    Public Shared Sub allQuerySearch()
        _checkedPostsNotBooks = New List(Of Post)
        _checkedPostsAndBooks = New List(Of Post)
        Console.ForegroundColor = ConsoleColor.DarkGreen
        Console.WriteLine("Starting...")
        Console.ResetColor()
        oneQuerySearch(TldUrl & "/search/sss?query=isbn")
        oneQuerySearch(TldUrl & "/search/sss?query=textbook")
        oneQuerySearch(TldUrl & "/search/bka?query=college")
        oneQuerySearch(TldUrl & "/search/bka?query=university")
        oneQuerySearch(TldUrl & "/search/bka?query=text")
        oneQuerySearch(TldUrl & "/search/bka?query=978")
        Console.WriteLine("Done!")
    End Sub

    Private Shared Sub oneQuerySearch(searchURL As String)
        Dim postNotBooks As Post                                    'this is a partially populated post object, just the post and not the books
        Dim postAndBooks As Post                                    'this is a fully populated post object, including the books
        Dim wc As New Net.WebClient
        Dim resultURL As String : resultURL = ""                     'URL of search result page
        Dim updatedSearchURL As String : updatedSearchURL = ""       'search result URL... this gets iterated through pagination
        Dim startPos As Integer : startPos = 1 'this is the start position of the search in the search results and will be incremented to find different results
        Dim searchPage As String : searchPage = ""                   'HTML of whole search result page

        'set or reset global variables
        searchPage = wc.DownloadString(searchURL)
        If Not searchPage Like "*Nothing found for that search*" Then
            Do While InStr(startPos, searchPage, _resultHook) > 0
                Console.WriteLine("On Result Number " & _checkedPostsNotBooks.Count + 1)
                Console.WriteLine("Cursory examination of post " & getLinkFromCLSearchResults(searchPage, startPos))

                'if not a nearby result
                If Not getLinkFromCLSearchResults(searchPage, startPos) Like "*http*" And Not getLinkFromCLSearchResults(searchPage, startPos) Like "*.org*" Then 'this prevents it from showing "nearby results"
                    Console.WriteLine("Learning more about post " & getLinkFromCLSearchResults(searchPage, startPos))

                    postNotBooks = New Post(TldUrl & getLinkFromCLSearchResults(searchPage, startPos), False)

                    'make sure it wasn't aleady checked this session
                    Dim match As Boolean = False
                    For Each p In _checkedPostsNotBooks
                        If p.Equals(postNotBooks) Then
                            match = True
                            Exit For
                        End If
                    Next
                    Console.ForegroundColor = ConsoleColor.DarkGray
                    Console.WriteLine("- Posts:  Partially checked: " & _checkedPostsNotBooks.Count)
                    Console.WriteLine("- Posts:  Fully checked:  " & _checkedPostsAndBooks.Count)
                    Console.WriteLine("- Posts:  Unparseable:  " & _checkedPostsNotBooks.Where(Function(x) x.IsParsable = False).Count)
                    Console.WriteLine("- Posts:  Parseable: " & _checkedPostsAndBooks.Where(Function(x) x.IsParsable = True).Count)
                    Console.WriteLine("- Books:  Winners: " & _checkedPostsAndBooks.SelectMany(Function(x) x.Books).Where(Function(y) y.IsWinner = True).Count)
                    Console.WriteLine("- Books:  Maybes: " & _checkedPostsAndBooks.SelectMany(Function(x) x.Books).Where(Function(y) y.IsMaybe = True).Count)
                    Console.WriteLine("- Books:  HVSB: " & _checkedPostsAndBooks.SelectMany(Function(x) x.Books).Where(Function(y) y.IsHVSB = True).Count)
                    Console.WriteLine("- Books:  HVOBO: " & _checkedPostsAndBooks.SelectMany(Function(x) x.Books).Where(Function(y) y.IsHVOBO = True).Count)
                    Console.WriteLine("- Books:  Trash: " & _checkedPostsAndBooks.SelectMany(Function(x) x.Books).Where(Function(y) y.IsTrash = True).Count)
                    Console.ResetColor()

                    If Not match _
                    AndAlso Not postNotBooks.IsMagazinePost Then
                        If postNotBooks.IsParsable Then
                            postAndBooks = New Post(postNotBooks)

                            WriteSearchResult(postAndBooks)
                            _checkedPostsAndBooks.Add(postAndBooks)
                        Else
                            Console.ForegroundColor = ConsoleColor.DarkRed
                            Console.WriteLine("Unparseable post: " & postNotBooks.URL)
                            Debug.Print("Unparseable post: " & postNotBooks.URL)
                            Console.ResetColor()
                            postNotBooks.IsParsable = False
                        End If
                    Else
                        postNotBooks.IsParsable = False
                    End If
                    _checkedPostsNotBooks.Add(postNotBooks)
                End If

                'do the pagination
                startPos = InStr(startPos, searchPage, _resultHook) + 100
                If _checkedPostsNotBooks.Count Mod 100 = 0 Then
                    If InStr(1, searchURL, "&") = 0 Then
                        updatedSearchURL = searchURL & "?s=" & _checkedPostsNotBooks.Count
                    Else
                        updatedSearchURL = searchURL & "&s=" & _checkedPostsNotBooks.Count
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
        For Each b As Book In post.Books
            proceed = True
            Console.WriteLine("Querying BookScouter about Book")
            b.GetDataFromBookscouter()
            Console.WriteLine("Writing Book Info")
            If post.IsParsable AndAlso b.IsParsable Then
                If b.IsWinner() Then
                    Console.ForegroundColor = ConsoleColor.Green
                    Console.Write("WINNER WINNER WINNER!")
                    Console.ResetColor()
                    resultType = "Winners"
                    desc = "Profitable book deals (winners) in the " & theCity & " area."
                    title = theCity & " Winners- Books"
                    outfile = title & ".xml"
                ElseIf b.IsMaybe() Then
                    Console.ForegroundColor = ConsoleColor.Cyan
                    Console.Write("MAYBE MAYBE MAYBE!")
                    Console.ResetColor()
                    resultType = "Maybes"
                    desc = "Potentially profitable book deals (maybes) in the " & theCity & " area."
                    title = theCity & " Maybes- Books"
                    outfile = title & ".xml"
                ElseIf b.IsHVSB() Then
                    Console.ForegroundColor = ConsoleColor.Magenta
                    Console.Write("VALUABLE STALE BOOK!")
                    Console.ResetColor()
                    resultType = "HVSBs"
                    desc = "High value stale books in the " & theCity & " area.  These books can be sold for a profit, but only if the seller (who hasn't been successful selling them at the current asking price) will come down on the price a bit."
                    title = theCity & " Valuable Stale Books"
                    outfile = theCity & " Valuable Stale Books.xml"
                ElseIf b.IsHVOBO() Then
                    Console.ForegroundColor = ConsoleColor.Magenta
                    Console.Write("VALUABLE 'OR BEST OFFER' BOOK!")
                    Console.ResetColor()
                    resultType = "HVOBOs"
                    desc = "High value 'or best offer' books in the " & theCity & " area.  These books can be sold for a profit, but only if the seller (who indicated that they'd consider the best offer) will come down on the price a bit."
                    title = theCity & " Valuable 'Or Best Offer' Books"
                    outfile = theCity & " Valuable Or Best Offer Books.xml"
                Else
                    Console.ForegroundColor = ConsoleColor.DarkRed
                    Console.WriteLine("Trashed book found in: " & post.URL)
                    resultType = "trash"
                    Console.ResetColor()
                    proceed = False
                End If
            Else
                Console.ForegroundColor = ConsoleColor.DarkRed
                Console.WriteLine("Unparseable post or book:  " & post.URL)
                Debug.Print("Unparseable post or book:  " & post.URL)
                resultType = "unparseable"
                Console.ResetColor()
                proceed = False
            End If

            If proceed Then

                If Not AlreadyInRSSFeed(b.ID, resultType, Sftp, SftpDirectory, City, SftpURL, "book") Then
                    Console.ForegroundColor = ConsoleColor.Yellow
                    Console.WriteLine(" NOT YET IN THE FEEDS- ADDING!")
                    Console.ResetColor()

                    If Not FeedAlreadyExists("book", resultType, Sftp, SftpDirectory, City) Then
                        rssFeed = New RSSFeed(title, wwwRoot & "showfeed.php?feed=" & replaceSpacesWithTwenty(Path.GetFileName(outfile)), desc, resultType, outfile)
                    Else
                        rssFeed = New RSSFeed(wwwRoot & "leads/" & replaceSpacesWithTwenty(outfile))
                    End If


                    Dim postLink As String = "https://href.li/?" & post.URL 'arbitext:postLink 
                    Dim resultURL As String = wwwRoot & "showitem.php?item=" & replaceSpacesWithTwenty(Path.GetFileName(rssFeed.FileName)) & "|" & b.ID
                    Dim theDesc As String = getDesc(resultType, post.City, b.AskingPrice, b.Profit, b.BuybackAmount)
                    WriteRSSItem(rssFeed.Document, b.Title, resultURL, post.UpdateDate, theDesc,
                                 b.ID, postLink, post.Title, post.City, b.Title, b.IsbnFromPost, b.AskingPrice,
                                 b.BookscouterSiteLink, b.BuybackAmount, b.Profit, b.ProfitPercentage, b.isOBO, post.Image, b.ImageURL)
                    PushUpdatedXML(rssFeed, Sftp)
                Else
                    Console.ForegroundColor = ConsoleColor.DarkYellow
                    Console.WriteLine("  Already in the feeds, no need to add")
                    Console.ResetColor()
                End If
            End If 'proceed check
        Next b

    End Sub

End Class
