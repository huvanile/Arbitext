Imports Arbitext.ExcelHelpers
Imports Arbitext.ArbitextHelpers
Imports Arbitext.CraigslistHelpers
Imports Arbitext.SoundHelpers

Public Class SingleManualCheck
    Public Shared Sub SingleManualCheck()
        Dim ret As Boolean      'return value for downloading book cover to file
        Dim post As Post
        Dim wc As New Net.WebClient
        resetSingleCheckWS()
        grabPHPicIfNeeded()
        IO.File.Delete(Environ("Temp") & "\cover.jpg")

        With ThisAddIn.AppExcel

            'parse this page
            If .Sheets("Single Check").Range("c3").Value = "" Then
                MsgBox("Put a craigslist book sale posting URL in cell C3.  Flipping idiots!", vbCritical, ThisAddIn.Title)
            Else
                post = New Post(.Sheets("Single Check").Range("c3").Value)

                If Not post.Title Like "*Page Not Found*" Then
                    .Sheets("Single Check").Range("c15").value2 = post.isMulti
                    If Not post.isMulti Then

                        .Sheets("Single Check").Range("f1").value2 = post.PostDate
                        .Sheets("Single Check").Range("c5").value2 = post.Title

                        'do flag checks
                        .Sheets("Single Check").Range("c16") = post.Book.aLaCarte
                        .Sheets("Single Check").Range("c17") = post.Book.isWeirdEdition
                        .Sheets("Single Check").Range("c18") = post.Book.isPDF
                        .Sheets("Single Check").Range("c19") = post.Book.isOBO

                        'write data
                        .Sheets("Single Check").Range("c6").value2 = post.Book.Isbn13
                        .Sheets("Single Check").Range("c6").Hyperlinks.Add(anchor:= .Sheets("Single Check").Range("c6"), Address:=post.Book.BookscouterSiteLink, TextToDisplay:="'" & post.Book.Isbn13)
                        .Sheets("Single Check").Range("c7").value2 = post.AskingPrice
                        .Sheets("Single Check").Range("f2").value2 = post.UpdateDate
                        .Sheets("Single Check").Range("c4").value2 = post.City
                        .Sheets("Single Check").Range("b23").value2 = post.Body
                        .Sheets("Single Check").Range("C8") = post.Book.BuybackAmount
                        .Sheets("Single Check").Range("C8").Hyperlinks.Add(anchor:= .Sheets("Single Check").Range("C8"), Address:=post.Book.BuybackLink, TextToDisplay:="'" & post.Book.BuybackAmount)
                        .Sheets("Single Check").Range("C10").value2 = post.Book.Profit
                        .Sheets("Single Check").Range("C11").value2 = post.Book.ProfitPercentage
                        .Sheets("Single Check").Range("C12").value2 = post.Book.MinAskingPriceForDesiredProfit

                        If post.Book.ImageURL <> "" And post.Book.ImageURL <> "(unknown)" Then
                            wc.DownloadFile(post.Book.ImageURL, Environ("Temp") & "\cover.jpg")
                            If IO.File.Exists(Environ("Temp") & "\cover.jpg") Then
                                .Sheets("Single Check").Shapes.AddPicture(fileName:=Environ("Temp") & "\cover.jpg", LinkToFile:=0, SaveWithDocument:=-1, Left:=657, Top:=56, Width:=192, Height:=240)
                            End If
                        Else
                            '0 = msoFalse, -1 = msoTrue
                            .Sheets("Single Check").Shapes.AddPicture(fileName:=Environ("Temp") & "\placeholder.jpg", LinkToFile:=0, SaveWithDocument:=-1, Left:=657, Top:=56, Width:=192, Height:=240)
                        End If

                        On Error Resume Next
                        System.IO.File.Delete(Environ("Temp") & "\cover.jpg")
                        On Error GoTo 0

                        If post.Image <> "" Then
                            wc.DownloadFile(post.Image, Environ("Temp") & "\cover.jpg")
                            If IO.File.Exists(Environ("Temp") & "\cover.jpg") Then
                                '0 = msoFalse, -1 = msoTrue
                                .Sheets("Single Check").Shapes.AddPicture(fileName:=Environ("Temp") & "\cover.jpg", LinkToFile:=0, SaveWithDocument:=-1, Left:=657, Top:=339, Width:=192, Height:=240)
                            End If
                        Else
                            '0 = msoFalse, -1 = msoTrue
                            .Sheets("Single Check").Shapes.AddPicture(fileName:=Environ("Temp") & "\placeholder.jpg", LinkToFile:=0, SaveWithDocument:=-1, Left:=657, Top:=339, Width:=192, Height:=240)
                        End If
                        .Sheets("Single Check").Range("c20").value2 = post.Book.IsWinner

                    Else
                        Select Case MsgBox("This looks like a multipost. Should we analyze it using multipost manual checker?", vbYesNoCancel, ThisAddIn.Title)
                            Case vbYes
                                If Not doesWSExist("Multipost Manual Checks") Then
                                    BuildWSMultipostManualCheck.BuildWSMultipostManualCheck()
                                End If
                                With .Sheets("Multipost Manual Checks")
                                    .Range("c3").value2 = post.URL
                                    .Activate
                                End With
                                MultipostManualCheck.MultipostManualCheck()
                        End Select
                    End If
                Else '404 error
                    MsgBox("Page not found 404 error", vbInformation, ThisAddIn.Title)
                End If
            End If
        End With
        wc = Nothing
        post = Nothing
    End Sub

    Private Shared Sub resetSingleCheckWS()
        With ThisAddIn.AppExcel
            .Sheets("Single Check").Range("c4:c20") = ""
            .Sheets("Single Check").Range("e4:e5") = ""
            .Sheets("Single Check").Range("b23") = ""
            .Sheets("Single Check").Range("c10") = ""
            .Sheets("Single Check").Range("c11") = ""
            .Sheets("Single Check").Range("c12") = ""
            deleteAllPics("Single Check")
        End With
    End Sub



End Class
