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

Public Class SingleManualCheck
    Public Shared Sub SingleManualCheck()
        Dim ret As Boolean      'return value for downloading book cover to file
        Dim wc As New System.Net.WebClient
        resetSingleCheckWS()
        grabPHPicIfNeeded()
        IO.File.Delete(Environ("Temp") & "\cover.jpg")

        With ThisAddIn.AppExcel

            'parse this page
            If .Sheets("Single Check").Range("c3").Value = "" Then
                MsgBox("Put a craigslist book sale posting URL in cell C3.  Flipping idiots!", vbCritical, ThisAddIn.Title)
            Else
                ThisAddIn.PostURL = .Sheets("Single Check").Range("c3").Value
                learnAboutPost(ThisAddIn.PostURL)

                If Not ThisAddIn.PostTitle Like "*Page Not Found*" Then
                    If Not isMulti(ThisAddIn.PostBody) And Not isMulti(ThisAddIn.PostTitle) Then
                        .Sheets("Single Check").Range("c15").value2 = "no"
                        .Sheets("Single Check").Range("f1").value2 = ThisAddIn.PostDate
                        .Sheets("Single Check").Range("c5").value2 = ThisAddIn.PostTitle
                        doFlagChecks(ThisAddIn.PostBody)
                        doFlagChecks(ThisAddIn.PostTitle)

                        'write data
                        .Sheets("Single Check").Range("c6").value2 = getISBN(ThisAddIn.PostBody)
                        .Sheets("Single Check").Range("c6").Hyperlinks.Add(anchor:= .Sheets("Single Check").Range("c6"), Address:="https://bookscouter.com/prices.php?isbn=" & ThisAddIn.PostISBN & "&all", TextToDisplay:="'" & .Sheets("Single Check").Range("c6").Value)
                        .Sheets("Single Check").Range("c7").value2 = ThisAddIn.PostAskingPrice
                        .Sheets("Single Check").Range("f2").value2 = ThisAddIn.PostUpdateDate
                        .Sheets("Single Check").Range("c4").value2 = ThisAddIn.PostCity
                        .Sheets("Single Check").Range("b23").value2 = ThisAddIn.PostBody
                        If Not .Sheets("Single Check").Range("c6").value2 Like "(*" Then
                            askBSAboutBook(ThisAddIn.PostISBN)
                            .Sheets("Single Check").Range("C8") = ThisAddIn.PostSellingPrice
                            .Sheets("Single Check").Range("C8").Hyperlinks.Add(anchor:= .Sheets("Single Check").Range("C8"), Address:=ThisAddIn.BsBuybackLink, TextToDisplay:="'" & .Sheets("Single Check").Range("C8").Value)
                            If ThisAddIn.ImageURL <> "" And ThisAddIn.ImageURL <> "(unknown)" Then
                                wc.DownloadFile(ThisAddIn.ImageURL, Environ("Temp") & "\cover.jpg")
                                If IO.File.Exists(Environ("Temp") & "\cover.jpg") Then
                                    .Sheets("Single Check").Shapes.AddPicture(fileName:=Environ("Temp") & "\cover.jpg", LinkToFile:=0, SaveWithDocument:=-1, Left:=657, Top:=56, Width:=192, Height:=240)
                                End If
                            Else
                                '0 = msoFalse, -1 = msoTrue
                                .Sheets("Single Check").Shapes.AddPicture(fileName:=Environ("Temp") & "\placeholder.jpg", LinkToFile:=0, SaveWithDocument:=-1, Left:=657, Top:=56, Width:=192, Height:=240)
                            End If
                        Else
                            .Sheets("Single Check").Range("c8") = "Not checked because unknown ISBN"
                        End If
                        On Error Resume Next
                        System.IO.File.Delete(Environ("Temp") & "\cover.jpg")
                        On Error GoTo 0
                        If ThisAddIn.PostCLImg <> "" Then
                            wc.DownloadFile(ThisAddIn.PostCLImg, Environ("Temp") & "\cover.jpg")
                            If IO.File.Exists(Environ("Temp") & "\cover.jpg") Then
                                '0 = msoFalse, -1 = msoTrue
                                .Sheets("Single Check").Shapes.AddPicture(fileName:=Environ("Temp") & "\cover.jpg", LinkToFile:=0, SaveWithDocument:=-1, Left:=657, Top:=339, Width:=192, Height:=240)
                            End If
                        Else
                            '0 = msoFalse, -1 = msoTrue
                            .Sheets("Single Check").Shapes.AddPicture(fileName:=Environ("Temp") & "\placeholder.jpg", LinkToFile:=0, SaveWithDocument:=-1, Left:=657, Top:=339, Width:=192, Height:=240)
                        End If

                        'WINNER!!!
                        If IsNumeric(.Sheets("Single Check").Range("c7").value2) _
                        And IsNumeric(.Sheets("Single Check").Range("c8").value2) _
                        And IsNumeric(.Sheets("Single Check").Range("c10").value2) _
                        And IsNumeric(.Sheets("Single Check").Range("c11").value2) Then
                            If .Sheets("Single Check").Range("c10").value2 >= ThisAddIn.MinTolerableProfit Then
                                PlayWAV("money")
                                .Sheets("Single Check").Range("c20").value2 = "YES"
                            Else
                                .Sheets("Single Check").Range("c20").value2 = "no"
                            End If
                        Else
                            .Sheets("Single Check").Range("c20").value2 = "no"
                        End If
                    Else
                        .Sheets("Single Check").Range("c15").value2 = "YES"
                        Select Case MsgBox("This looks like a multipost. Should we analyze it using multipost manual checker?", vbYesNoCancel, ThisAddIn.Title)
                            Case vbYes
                                If Not doesWSExist("Multipost Manual Checks") Then
                                    BuildWSMultipostManualCheck.BuildWSMultipostManualCheck()
                                End If
                                With .Sheets("Multipost Manual Checks")
                                    .Range("c3").value2 = ThisAddIn.PostURL
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
    End Sub

    Private Shared Sub resetSingleCheckWS()
        With ThisAddIn.AppExcel
            .Sheets("Single Check").Range("c4:c20") = ""
            .Sheets("Single Check").Range("e4:e5") = ""
            .Sheets("Single Check").Range("b23") = ""
            .Sheets("Single Check").Range("c10").FormulaLocal = "=IFERROR(C8-c7,0)"
            .Sheets("Single Check").Range("c11").FormulaLocal = "=IFERROR(C10/c7,0)"
            .Sheets("Single Check").Range("c12").FormulaLocal = "=IF(ROUND(C8-$I$1,0)>0,ROUND(C8-$I$1,0),0)"
            deleteAllPics("Single Check")
        End With
    End Sub

    Private Shared Sub doFlagChecks(ByVal theStr As String)
        With ThisAddIn.AppExcel
            'do flag checks
            If aLaCarte(theStr) Then
                .Sheets("Single Check").Range("c16") = "YES"
            Else
                .Sheets("Single Check").Range("c16") = "no"
            End If
            If isWeirdEdition(theStr) Then
                .Sheets("Single Check").Range("c17") = "YES"
            Else
                .Sheets("Single Check").Range("c17") = "no"
            End If
            If isPDF(theStr) Then
                .Sheets("Single Check").Range("c18") = "YES"
            Else
                .Sheets("Single Check").Range("c18") = "no"
            End If
            If isOBO(theStr) Then
                .Sheets("Single Check").Range("c19") = "YES"
            Else
                .Sheets("Single Check").Range("c19") = "no"
            End If
        End With
    End Sub

End Class
