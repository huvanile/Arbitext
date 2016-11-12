Imports Arbitext.RegistryHelpers
Imports Arbitext.ExcelHelpers
Imports Arbitext.ArbitextHelpers
Imports Arbitext.PushbulletHelpers
Imports Arbitext.BookscouterHelpers
Imports Arbitext.CraigslistHelpers
Imports Arbitext.SoundHelpers
Imports Arbitext.StringHelpers
Imports Microsoft.Office.Interop.Excel

Public Class MultipostManualCheck
    Public Shared Sub MultipostManualCheck()
        With ThisAddIn.AppExcel
            .Sheets("Multipost Manual Checks").Range("c4:c" & lastUsedRow("Multipost Manual Checks")) = ""
            .Sheets("Multipost Manual Checks").Range("g8:g" & lastUsedRow("Multipost Manual Checks")) = ""
            .Sheets("Multipost Manual Checks").Range("h7").Value = 0

            deleteAllPics()
            grabPHPicIfNeeded()
            learnAboutPost(.Sheets("Multipost Manual Checks").Range("c3").Value2)

            If bookCount(ThisAddIn.PostBody) = 0 Then
                MsgBox("No ISBN's present in post.  Unable to parse.", vbInformation, ThisAddIn.Title)
            Else
                If Not ThisAddIn.PostTitle Like "*Page Not Found*" Then
                    .Sheets("Multipost Manual Checks").Range("c4") = ThisAddIn.PostDate
                    .Sheets("Multipost Manual Checks").Range("c5") = ThisAddIn.PostTitle
                    .Sheets("Multipost Manual Checks").Range("c6") = ThisAddIn.PostUpdateDate
                    .Sheets("Multipost Manual Checks").Range("c7") = ThisAddIn.PostCity
                    doAMultiManualCheck("2br")
                    If parsedBookCount() = 0 Or .Range("g10").Value2 = "" Then doAMultiManualCheck("1br")
                    If parsedBookCount() = 0 Or .Range("g10").Value2 = "" Then doAMultiManualCheck("p")

                    If parsedBookCount() = 0 Then
                        MsgBox("No viable ISBN's found.  Unable to parse.", vbInformation, ThisAddIn.Title)
                    Else
                        MsgBox("Done", vbInformation, ThisAddIn.Title)
                    End If

                Else
                    MsgBox("Page not found 404 error", vbInformation, ThisAddIn.Title)
                End If
            End If
        End With
    End Sub

    Private Shared Sub doAMultiManualCheck(checkMethod As String)
        Dim i As Integer        'book count
        Dim s As Long           'splitholder start position, increments as books are parsed
        Dim t As Integer         'to loop through the splitholder <BR> results
        Dim tmpPostBody As String
        Dim splitholder() As String
        Dim tmpPic As String : tmpPic = Environ("Temp") & "\cover.jpg"
        Dim phPic As String : phPic = Environ("Temp") & "\placeholder.jpg"
        Dim theCol As String = ""
        Dim theRow As Long = 0
        Dim highestValue As Integer = 0
        s = 0

        With ThisAddIn.AppExcel

            For i = 1 To bookCount(ThisAddIn.PostBody)

                Diagnostics.Debug.Print("book " & i) '''''''''''''''''''''''''''
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
                    Diagnostics.Debug.Print(splitholder(t)) '''''''''''''''''''''''''''
                    If Not getISBN(splitholder(t)) Like "(*" Then 'match!

                        'determine result row
                        Select Case i
                            Case 1, 2
                                theRow = 10
                            Case 3, 4
                                theRow = 25
                            Case 5, 6
                                theRow = 40
                            Case 7, 8
                                theRow = 55
                            Case 9, 10
                                theRow = 70
                            Case 11, 12
                                theRow = 85
                            Case 13, 14
                                theRow = 100
                            Case 15, 16
                                theRow = 115
                            Case 17, 18
                                theRow = 130
                            Case 19, 20
                                theRow = 145
                        End Select

                        'determine result column
                        If i Mod 2 = 0 Then thecol = "g" Else thecol = "C"

                        'select result row
                        .Range(thecol & theRow).Select()

                        .ActiveCell.Value = getISBN(splitholder(t)) 'isbn
                        .ActiveCell.Hyperlinks.Add(Anchor:=ThisAddIn.AppExcel.ActiveCell, Address:="https://bookscouter.com/prices.php?isbn=" & ThisAddIn.AppExcel.ActiveCell.Value2 & "&all", TextToDisplay:="'" & ThisAddIn.AppExcel.ActiveCell.Value2)
                        .ActiveCell.Offset(1, 0).Value2 = getAskingPrice(splitholder(t)) 'asking price

                        IO.File.Delete(tmpPic)

                        'don't hit bookscouter if there's already a sell price present (maybe from a prior pass)
                        If .ActiveCell.Offset(2, 0).Value2 = "" Then
                            askBSAboutBook(.ActiveCell.Value)
                            .ActiveCell.Offset(2, 0).Value2 = ThisAddIn.PostSellingPrice
                            If ThisAddIn.PostSellingPrice > highestValue Then highestValue = ThisAddIn.PostSellingPrice
                            ThisAddIn.AppExcel.ActiveCell.Offset(2, 0).Hyperlinks.Add(Anchor:= .ActiveCell.Offset(2, 0), Address:=ThisAddIn.BsBuybackLink, TextToDisplay:="'" & .ActiveCell.Offset(2, 0).Value)
                            If ThisAddIn.ImageURL <> "None" And ThisAddIn.ImageURL <> "" Then
                                'If Download_File(ThisAddIn.ImageURL, tmpPic) Then
                                '    .ActiveSheet.Shapes.AddPicture(fileName:=tmpPic, Left:= .Round(.Columns(.ActiveCell.Offset(0, 1).Column).Left, 0) + 10, Top:= .Round(.Rows(.ActiveCell.Offset(-1, 0).Row).Top, 0) + 10, Width:=192, Height:=240)
                                'Else
                                '    .ActiveSheet.Shapes.AddPicture(fileName:=phPic, Left:= .Round(.Columns(.ActiveCell.Offset(0, 1).Column).Left, 0) + 10, Top:= .Round(.Rows(.ActiveCell.Offset(-1, 0).Row).Top, 0) + 10, Width:=192, Height:=240)
                                'End If
                            Else
                                .ActiveSheet.Shapes.AddPicture(fileName:=phPic, Left:= .Round(.Columns(.ActiveCell.Offset(0, 1).Column).Left, 0) + 10, Top:= .Round(.Rows(.ActiveCell.Offset(-1, 0).Row).Top, 0) + 10, Width:=192, Height:=240)
                            End If
                        End If

                        .ActiveCell.Offset(3, 0).FormulaLocal = "=IFERROR(" & .ActiveCell.Offset(2, 0).Address & "-" & .ActiveCell.Offset(1, 0).Address & ",0)"
                        .ActiveCell.Offset(4, 0).FormulaLocal = "=IFERROR(" & .ActiveCell.Offset(3, 0).Address & "/" & .ActiveCell.Offset(1, 0).Address & ",0)"
                        .ActiveCell.Offset(5, 0).FormulaLocal = "=IFERROR(IF(ROUND(" & .ActiveCell.Offset(2, 0).Address & "-$f$1,0)>0,ROUND(" & .ActiveCell.Offset(2, 0).Address & "-$f$1,0),0),""-"")"
                        If .ActiveCell.Offset(3, 0).Value2 > .Range("h7").Value2 Then .Range("h7").Value2 = .ActiveCell.Offset(3, 0).Value2

                        'do flag checks
                        If aLaCarte(splitholder(t)) Then
                            .ActiveCell.Offset(6, 0).Value2 = "YES"
                        End If
                        If isWeirdEdition(splitholder(t)) Then
                            .ActiveCell.Offset(7, 0).Value2 = "YES"
                        End If
                        If isPDF(splitholder(t)) Then
                            .ActiveCell.Offset(8, 0).Value2 = "YES"
                        End If
                        If isOBO(splitholder(t)) Then
                            .ActiveCell.Offset(9, 0).Value2 = "YES"
                        End If

                        'WINNER!!!
                        If IsNumeric(.ActiveCell.Offset(1, 0)) _
                        And IsNumeric(.ActiveCell.Offset(2, 0)) _
                        And IsNumeric(.ActiveCell.Offset(3, 0)) _
                        And IsNumeric(.ActiveCell.Offset(4, 0)) Then
                            If .ActiveCell.Offset(3, 0).Value2 >= ThisAddIn.MinTolerableProfit Then
                                If ThisAddIn.OnWinnersOK Then
                                    sendPushbulletNote(ThisAddIn.Title, "Winner Found!")
                                End If
                                PlayWAV("money")
                                .ActiveCell.Offset(10, 0).Value2 = "YES"
                            End If
                        End If

                        .ActiveCell.Offset(11, 0).Value2 = clean(splitholder(t), False, False, False, False, False, False) 'post body

                        s = t + 1
                        Exit For
                    End If
                Next t
            Next i

        End With
    End Sub

    Private Shared Function parsedBookCount() As Integer
        parsedBookCount = 0
        With ThisAddIn.AppExcel
            If .Range("c10").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g10").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c25").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g25").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c40").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g40").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c55").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g55").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c70").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g70").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c85").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g85").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c100").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g100").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c115").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g115").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c130").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g130").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c145").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g145").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c160").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g160").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c175").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g175").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c190").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g190").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c205").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g205").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("c220").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
            If .Range("g220").Value2 <> "" Then parsedBookCount = parsedBookCount + 1
        End With
    End Function

End Class
