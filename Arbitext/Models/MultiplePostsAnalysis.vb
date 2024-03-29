﻿Imports Arbitext.ExcelHelpers
Imports ArbitextClassLibrary
Imports ArbitextClassLibrary.CraigslistHelpers
Imports System.Threading
Imports ArbitextClassLibrary.Globals

Public Class MultiplePostsAnalysis
    Dim _checkedPostsNotBooks As List(Of Post)      'list of posts checked before populating the post object with the accompanying books
    Dim _checkedPostsAndBooks As List(Of Post)      'list of posts checked with all of the books included

#Region "constructors"

    Sub New()
        TldUrl = "http://" & City & ".craigslist.org"
        _checkedPostsNotBooks = New List(Of Post)
        _checkedPostsAndBooks = New List(Of Post)
        ThisAddIn.t1 = New Thread(AddressOf allQuerySearch)
        With ThisAddIn.t1
            If .IsAlive() Then .Abort()
            .IsBackground = True
            .Priority = ThreadPriority.BelowNormal
            .SetApartmentState(ApartmentState.STA)
            .Start()
        End With
    End Sub

#End Region


#Region "Methods"

    Private Sub allQuerySearch()
        ThisAddIn.Proceed = True
        ThisAddIn.AppExcel.ScreenUpdating = True
        Ribbon1.tpnAuto.showLblRecordSafe("")
        Ribbon1.tpnAuto.UpdateLblNumberSafe("Starting...")
        If Not doesWSExist("Color Legend") Then BuildWSColorLegend.buildWSColorLegend()
        If ThisAddIn.Proceed Then
            Dim searchURL As String = ""


            searchURL = TldUrl & "/search/sss?query=isbn"
            oneQuerySearch(searchURL)

            If ThisAddIn.Proceed Then
                searchURL = TldUrl & "/search/sss?query=textbook"
                oneQuerySearch(searchURL)
            End If

            If ThisAddIn.Proceed Then
                searchURL = TldUrl & "/search/bka?query=college"
                oneQuerySearch(searchURL)
            End If

            If ThisAddIn.Proceed Then
                searchURL = TldUrl & "/search/bka?query=university"
                oneQuerySearch(searchURL)
            End If

            If ThisAddIn.Proceed Then
                searchURL = TldUrl & "/search/bka?query=text"
                oneQuerySearch(searchURL)
            End If

            If ThisAddIn.Proceed Then
                searchURL = TldUrl & "/search/bka?query=978"
                oneQuerySearch(searchURL)
            End If

            Ribbon1.tpnAuto.hideLblRecordSafe("")
            Ribbon1.tpnAuto.UpdateLblStatusSafe("Done!")
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
        Dim wc As New Net.WebClient
        Dim resultURL As String : resultURL = ""                     'URL of search result page
        Dim updatedSearchURL As String : updatedSearchURL = ""       'search result URL... this gets iterated through pagination
        Dim startPos As Integer : startPos = 1 'this is the start position of the search in the search results and will be incremented to find different results
        Dim searchPage As String : searchPage = ""                   'HTML of whole search result page

        'set or reset global variables
        searchPage = wc.DownloadString(searchURL)
        If Not searchPage Like "*Nothing found for that search*" Then
            Do While InStr(startPos, searchPage, ResultHook) > 0
                Ribbon1.tpnAuto.UpdateLblNumberSafe("On Result Number " & _checkedPostsNotBooks.Count + 1)
                Ribbon1.tpnAuto.UpdateLblStatusSafe("Cursory examination of post " & getLinkFromCLSearchResults(searchPage, startPos) & ", will ignore if out of town")

                'if not a nearby result
                If Not getLinkFromCLSearchResults(searchPage, startPos) Like "*http*" And Not getLinkFromCLSearchResults(searchPage, startPos) Like "*.org*" Then 'this prevents it from showing "nearby results"
                    Ribbon1.tpnAuto.UpdateLblStatusSafe("Learning more about post " & vbCrLf & getLinkFromCLSearchResults(searchPage, startPos))

                    postNotBooks = New Post(TldUrl & getLinkFromCLSearchResults(searchPage, startPos), False)

                    'ThisAddIn.AppExcel.StatusBar = "Currently on result number " & _checkedPosts.Count & " (old, in trash, or otherwise skipped: " & SearchSession.SkippedResultCount & ") (keepers: " & SearchSession.KeeperCount & ") (maybes: " & SearchSession.NegCount & ") [MULTIPOSTS: " & SearchSession.MultiCount & "]"
                    Ribbon1.tpnAuto.UpdateLblStatusSafe("Considering writing result to workbook")

                    'make sure it wasn't aleady checked this session
                    Dim match As Boolean = False
                    For Each p In _checkedPostsNotBooks
                        If p.Equals(postNotBooks) Then
                            match = True
                            Exit For
                        End If
                    Next

                    Ribbon1.tpnAuto.UpdateLblCountSafe("- Posts:  Partially checked: " & _checkedPostsNotBooks.Count & vbCrLf &
                                                    "- Posts:  Fully checked:  " & _checkedPostsAndBooks.Count & vbCrLf & vbCrLf &
                                                    "- Posts:  Unparseable:  " & _checkedPostsNotBooks.Where(Function(x) x.IsParsable = False).Count & vbCrLf &
                                                    "- Posts:  Parseable: " & _checkedPostsAndBooks.Where(Function(x) x.IsParsable = True).Count & vbCrLf & vbCrLf &
                                                    "- Books:  Winners: " & _checkedPostsAndBooks.SelectMany(Function(x) x.Books).Where(Function(y) y.IsWinner = True).Count & vbCrLf &
                                                    "- Books:  Maybes: " & _checkedPostsAndBooks.SelectMany(Function(x) x.Books).Where(Function(y) y.IsMaybe = True).Count & vbCrLf &
                                                    "- Books:  HVSBs: " & _checkedPostsAndBooks.SelectMany(Function(x) x.Books).Where(Function(y) y.IsHVSB = True).Count & vbCrLf &
                                                    "- Books:  HVOBOs: " & _checkedPostsAndBooks.SelectMany(Function(x) x.Books).Where(Function(y) y.IsHVOBO = True).Count & vbCrLf &
                                                    "- Books:  Trash: " & _checkedPostsAndBooks.SelectMany(Function(x) x.Books).Where(Function(y) y.IsTrash = True).Count)

                    If Not match _
                    AndAlso Not postNotBooks.IsMagazinePost Then
                        If postNotBooks.IsParsable Then
                            postAndBooks = New Post(postNotBooks)
                            Ribbon1.tpnAuto.UpdateLblStatusSafe("About to write result to workbook")
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
                            postNotBooks.IsParsable = False
                        End If
                    Else
                        postNotBooks.IsParsable = False
                    End If
                    _checkedPostsNotBooks.Add(postNotBooks)
                End If

                'do the pagination
                startPos = InStr(startPos, searchPage, ResultHook) + 100
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
            MsgBox("Nothing found for the search:" & vbCrLf & vbCrLf & searchURL, vbInformation, Title)
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
            Ribbon1.tpnAuto.UpdateLblStatusSafe("Making sure result wasn't already written")
            Ribbon1.tpnAuto.UpdateLblStatusSafe("Querying BookScouter about Book")
                b.GetDataFromBookscouter()
                Ribbon1.tpnAuto.UpdateLblStatusSafe("Writing Book Info")
            If post.IsParsable AndAlso b.IsParsable Then
                If b.IsWinner() Then
                    If Not doesWSExist("Winners") Then BuildWSResults.buildResultWS("Winners")
                    destSheet = "Winners"
                ElseIf b.IsMaybe() Then
                    If Not doesWSExist("Maybes") Then BuildWSResults.buildResultWS("Maybes")
                    destSheet = "Maybes"
                ElseIf b.IsHVSB() Then
                    If Not doesWSExist("HVSBs") Then BuildWSResults.buildResultWS("HVSBs")
                    destSheet = "HVSBs"
                ElseIf b.IsHVOBO() Then
                    If Not doesWSExist("HVOBOs") Then BuildWSResults.buildResultWS("HVOBOs")
                    destSheet = "HVOBOs"
                ElseIf b.IsTrash() Then
                    If Not doesWSExist("Trash") Then BuildWSResults.buildResultWS("Trash") Else unFilterTrash()
                    destSheet = "Trash"
                Else
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
                Ribbon1.tpnAuto.UpdateLblStatusSafe("Writing search result to " & destSheet)
                Ribbon1.tpnAuto.showLblCountsSafe("")

                .activate
                Dim r As Integer = lastUsedRow() + 1
                thinInnerBorder(.Range("a" & r & ":n" & r))

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
                    .range("l" & r).value2 = "'" & b.ID
                    .range("m" & r).value2 = b.ImageURL
                    .range("n" & r).value2 = post.Image
                    If b.aLaCarte() Then .Rows(r).font.colorindex = 3
                    If b.isOBO() Then .Rows(r).font.bold = True
                    If b.isWeirdEdition() Then .Rows(r).font.colorindex = 46
                    If b.isPDF() Then .Rows(r).font.colorindex = 53
                    If b.IsWinner() Then HandleWinner(post, b, r)
                    If b.IsMaybe() Then HandleMaybe(post, b, r)
                End If
            End With
        Next b

    End Sub

    Sub HandleWinner(post As Post, book As Book, r As Integer)
        SoundHelpers.PlayWAV("money")
        ThisAddIn.AppExcel.Range("a" & r & ":n" & r).Interior.ColorIndex = 35
    End Sub

    Sub HandleMaybe(post As Post, book As Book, r As Integer)
        ThisAddIn.AppExcel.Range("a" & r & ":n" & r).Interior.ColorIndex = 6
    End Sub

#End Region

End Class
