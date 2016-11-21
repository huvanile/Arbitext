Imports Arbitext.ExcelHelpers
Imports Arbitext.ArbitextHelpers

Public Class SinglePostAnalysis
    Sub New()
        ThisAddIn.TldUrl = "http://" & ThisAddIn.City & ".craigslist.org"
        Dim tmpURL As String = InputBox("What's the post URL?", ThisAddIn.Title).Trim
        If Not tmpURL = "" Then
            Dim post As Post = New Post(tmpURL)
            If post.IsParsable Then
                grabPHPicIfNeeded()
                If post.isMulti Then
                    doAMultiManualCheck(post)
                    MsgBox("Done!", vbOK, ThisAddIn.Title)
                Else
                    SingleManualCheck(post)
                End If
            Else
                MsgBox("Post is unparsable, unable to find books in post, or hit a 404 page", vbOK, ThisAddIn.Title)
            End If
        Else
            'hit cancel on inputbox, do nothing
        End If
    End Sub

#Region "Multipost checking"

    Private Function bookStartRow(theBookCount) As Short
        Select Case theBookCount + 1
            Case 1, 2
                Return 10
            Case 3, 4
                Return 25
            Case 5, 6
                Return 40
            Case 7, 8
                Return 55
            Case 9, 10
                Return 70
            Case 11, 12
                Return 85
            Case 13, 14
                Return 100
            Case 15, 16
                Return 115
            Case 17, 18
                Return 130
            Case 19, 20
                Return 145
        End Select
    End Function

    Private Sub doAMultiManualCheck(post As Post)
        Dim tmpPic As String : tmpPic = Environ("Temp") & "\cover.jpg"
        Dim phPic As String : phPic = Environ("Temp") & "\placeholder.jpg"
        Dim theCol As String = ""
        Dim wc As New Net.WebClient 'used to download covers
        Dim highestValue As Decimal = 0
        BuildWSMultipostManualCheck.BuildWSMultipostManualCheck()
        With ThisAddIn.AppExcel.Sheets("Multipost Manual Checks")
            .Range("c3").Value2 = post.URL
            .range("c3").hyperlinks.add(anchor:= .Range("c3"), Address:=post.URL, TextToDisplay:=post.URL)
            .Range("c4") = post.PostDate
            .Range("c5") = post.Title
            .Range("c6") = post.UpdateDate
            .Range("c7") = post.City

            If post.Books.Count = 0 Then
                .range("c10").value = "Post had no parsable books :("
            Else
                For b As Short = 0 To post.Books.Count - 1
                    Dim theRow As Long = bookStartRow(b)
                    If (b + 1) Mod 2 = 0 Then theCol = "g" Else theCol = "C"
                    BuildWSMultipostManualCheck.bookTile(.range(theCol & theRow).offset(-1, -1), b + 1)
                    .range(theCol & theRow).offset(-1, -1).value2 = post.Books(b).Title
                    .range(theCol & theRow).value2 = post.Books(b).Isbn13
                    .range(theCol & theRow).hyperlinks.add(anchor:= .range(theCol & theRow), Address:=post.Books(b).BookscouterSiteLink, TextToDisplay:=post.Books(b).Isbn13)
                    .range(theCol & theRow).offset(1, 0).value2 = post.Books(b).AskingPrice
                    .range(theCol & theRow).offset(2, 0).value2 = post.Books(b).BuybackAmount
                    .range(theCol & theRow).offset(3, 0).value2 = post.Books(b).Profit
                    .range(theCol & theRow).offset(4, 0).value2 = post.Books(b).ProfitPercentage
                    .range(theCol & theRow).offset(5, 0).value2 = post.Books(b).MinAskingPriceForDesiredProfit
                    .range(theCol & theRow).Offset(6, 0).Value2 = post.Books(b).aLaCarte()
                    .range(theCol & theRow).Offset(7, 0).Value2 = post.Books(b).isWeirdEdition()
                    .range(theCol & theRow).Offset(8, 0).Value2 = post.Books(b).isPDF()
                    .range(theCol & theRow).Offset(9, 0).Value2 = post.Books(b).isOBO()
                    If post.Books(b).IsWinner() Then
                        .range(theCol & theRow).offset(10, 0).value2 = "YES!!!"
                    ElseIf post.Books(b).IsMaybe() Then
                        .range(theCol & theRow).offset(10, 0).value2 = "No, but it's a MAYBE though"
                    Else
                        .range(theCol & theRow).offset(10, 0).value2 = "No :("
                    End If
                    If post.Books(b).BuybackAmount > highestValue Then highestValue = post.Books(b).BuybackAmount
                    .range(theCol & theRow).Offset(11, 0).Value2 = post.Books(b).SaleDescInPost
                    On Error Resume Next
                    System.IO.File.Delete(tmpPic)
                    wc.DownloadFile(post.Books(b).ImageURL, tmpPic)
                    If Not IO.File.Exists(tmpPic) Then tmpPic = phPic
                    Dim tmpLeft As Integer = Math.Round(.Columns(.range(theCol & theRow).Offset(0, 1).Column).Left, 0) + 10
                    Dim tmpTop As Integer = Math.Round(.Rows(.range(theCol & theRow).Offset(-1, 0).Row).Top, 0) + 10
                    .Shapes.AddPicture(fileName:=tmpPic, LinkToFile:=0, SaveWithDocument:=-1, Left:=tmpLeft, Top:=tmpTop, Width:=192, Height:=240)
                    On Error GoTo 0
                Next b
            End If
        End With
        wc = Nothing
    End Sub

    Private Function parsedBookCount() As Integer
        Dim tmp As Int16 = 0
        With ThisAddIn.AppExcel
            If .Range("c10").Value2 <> "" Then tmp = tmp + 1
            If .Range("g10").Value2 <> "" Then tmp = tmp + 1
            If .Range("c25").Value2 <> "" Then tmp = tmp + 1
            If .Range("g25").Value2 <> "" Then tmp = tmp + 1
            If .Range("c40").Value2 <> "" Then tmp = tmp + 1
            If .Range("g40").Value2 <> "" Then tmp = tmp + 1
            If .Range("c55").Value2 <> "" Then tmp = tmp + 1
            If .Range("g55").Value2 <> "" Then tmp = tmp + 1
            If .Range("c70").Value2 <> "" Then tmp = tmp + 1
            If .Range("g70").Value2 <> "" Then tmp = tmp + 1
            If .Range("c85").Value2 <> "" Then tmp = tmp + 1
            If .Range("g85").Value2 <> "" Then tmp = tmp + 1
            If .Range("c100").Value2 <> "" Then tmp = tmp + 1
            If .Range("g100").Value2 <> "" Then tmp = tmp + 1
            If .Range("c115").Value2 <> "" Then tmp = tmp + 1
            If .Range("g115").Value2 <> "" Then tmp = tmp + 1
            If .Range("c130").Value2 <> "" Then tmp = tmp + 1
            If .Range("g130").Value2 <> "" Then tmp = tmp + 1
            If .Range("c145").Value2 <> "" Then tmp = tmp + 1
            If .Range("g145").Value2 <> "" Then tmp = tmp + 1
            If .Range("c160").Value2 <> "" Then tmp = tmp + 1
            If .Range("g160").Value2 <> "" Then tmp = tmp + 1
            If .Range("c175").Value2 <> "" Then tmp = tmp + 1
            If .Range("g175").Value2 <> "" Then tmp = tmp + 1
            If .Range("c190").Value2 <> "" Then tmp = tmp + 1
            If .Range("g190").Value2 <> "" Then tmp = tmp + 1
            If .Range("c205").Value2 <> "" Then tmp = tmp + 1
            If .Range("g205").Value2 <> "" Then tmp = tmp + 1
            If .Range("c220").Value2 <> "" Then tmp = tmp + 1
            If .Range("g220").Value2 <> "" Then tmp = tmp + 1
        End With
        Return tmp
    End Function

#End Region

#Region "Single check"

    Private Sub SingleManualCheck(post As Post)
        Dim tmpPic As String : tmpPic = Environ("Temp") & "\cover.jpg"
        Dim phPic As String : phPic = Environ("Temp") & "\placeholder.jpg"
        Dim ret As Boolean          'return value for downloading book cover to file
        Dim wc As New Net.WebClient 'used to download covers
        If Not doesWSExist("Single Check") Then DeleteWS("Single Check")
        BuildWSSingleCheck.BuildWSSingleCheck()
        With ThisAddIn.AppExcel.Sheets("Single Check")

            .Range("f1").value2 = post.PostDate
            .Range("f2").value2 = post.UpdateDate
            .range("C3").value2 = post.URL
            .range("c3").hyperlinks.add(anchor:= .Range("c3"), Address:=post.URL, TextToDisplay:=post.URL)
            .Range("c4").value2 = post.City
            .Range("c5").value2 = post.Title
            .Range("c6").value2 = post.Books(0).Isbn13
            .Range("c6").Hyperlinks.Add(anchor:= .Range("c6"), Address:=post.Books(0).BookscouterSiteLink, TextToDisplay:="'" & post.Books(0).Isbn13)
            .Range("c7").value2 = post.AskingPrice
            .Range("C8") = post.Books(0).BuybackAmount
            .Range("C8").Hyperlinks.Add(anchor:= .Range("C8"), Address:=post.Books(0).BuybackLink, TextToDisplay:="'" & post.Books(0).BuybackAmount)
            .Range("C10").value2 = post.Books(0).Profit
            .Range("C11").value2 = post.Books(0).ProfitPercentage
            .Range("C12").value2 = post.Books(0).MinAskingPriceForDesiredProfit
            .Range("c15").Value2 = post.isMulti
            .Range("c16") = post.Books(0).aLaCarte()
            .Range("c17") = post.Books(0).isWeirdEdition()
            .Range("c18") = post.Books(0).isPDF()
            .Range("c19") = post.Books(0).isOBO()
            .Range("c20").value2 = post.Books(0).IsWinner()
            .Range("b23").value2 = post.Body

            'download amazon book pic if possible, otherwise just use placeholder pic
            On Error Resume Next
            System.IO.File.Delete(tmpPic)
            wc.DownloadFile(post.Books(0).ImageURL, tmpPic)
            On Error GoTo 0
            If Not IO.File.Exists(tmpPic) Then tmpPic = phPic
            Dim tmpLeft As Integer = 657
            Dim tmptop As Integer = 56
            .Shapes.AddPicture(fileName:=tmpPic, LinkToFile:=0, SaveWithDocument:=-1, Left:=tmpLeft, Top:=tmptop, Width:=192, Height:=240)

            'download craigslist book pic if possible, otherwise just use placeholder pic
            On Error Resume Next
            System.IO.File.Delete(tmpPic)
            wc.DownloadFile(post.Image, tmpPic)
            On Error GoTo 0
            If Not IO.File.Exists(tmpPic) Then tmpPic = phPic
            tmpLeft = 657
            tmptop = 339
            .Shapes.AddPicture(fileName:=tmpPic, LinkToFile:=0, SaveWithDocument:=-1, Left:=tmpLeft, Top:=tmptop, Width:=192, Height:=240)

        End With
        wc = Nothing
    End Sub

    Private Sub resetSingleCheckWS()
        With ThisAddIn.AppExcel.Sheets("Single Check")
            .Range("c4:c20") = ""
            .Range("e4:e5") = ""
            .Range("b23") = ""
            .Range("c10") = ""
            .Range("c11") = ""
            .Range("c12") = ""
            deleteAllPics("Single Check")
        End With
    End Sub

#End Region


End Class
