Imports Arbitext.ExcelHelpers
Imports Arbitext.ArbitextHelpers
Imports Arbitext.PushbulletHelpers
Imports Arbitext.SoundHelpers

Public Class MultipostManualCheck
    Public Shared Sub MultipostManualCheck()
        Dim post As Post
        With ThisAddIn.AppExcel
            .Sheets("Multipost Manual Checks").Range("c4:c" & lastUsedRow("Multipost Manual Checks")) = ""
            .Sheets("Multipost Manual Checks").Range("g8:g" & lastUsedRow("Multipost Manual Checks")) = ""
            .Sheets("Multipost Manual Checks").Range("h7").Value = 0

            deleteAllPics()
            grabPHPicIfNeeded()
            post = New Post(.Sheets("Multipost Manual Checks").Range("c3").Value2)

            If post.bookCount = 0 Then
                MsgBox("No ISBN's present in post.  Unable to parse.", vbInformation, ThisAddIn.Title)
            Else
                If Not post.Title Like "*Page Not Found*" Then
                    .Sheets("Multipost Manual Checks").Range("c4") = post.PostDate
                    .Sheets("Multipost Manual Checks").Range("c5") = post.Title
                    .Sheets("Multipost Manual Checks").Range("c6") = post.UpdateDate
                    .Sheets("Multipost Manual Checks").Range("c7") = post.City

                    doAMultiManualCheck(post)

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

    Private Shared Sub doAMultiManualCheck(post As Post)
        Dim tmpPic As String : tmpPic = Environ("Temp") & "\cover.jpg"
        Dim phPic As String : phPic = Environ("Temp") & "\placeholder.jpg"
        Dim theCol As String = ""
        Dim theRow As Long = 0
        Dim highestValue As Integer = 0
        Dim book As Book : book = Nothing

        With ThisAddIn.AppExcel

            For b As Short = 0 To post.Books.Count
                'determine result row
                Select Case b + 1
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
                If (b + 1) Mod 2 = 0 Then theCol = "g" Else theCol = "C"

                'select result row
                .Range(theCol & theRow).Select()

                'write result
                .ActiveCell.Value = post.Books(b) 'isbn
                .ActiveCell.Hyperlinks.Add(Anchor:=ThisAddIn.AppExcel.ActiveCell, Address:=post.Books(b).BookscouterSiteLink, TextToDisplay:="'" & ThisAddIn.AppExcel.ActiveCell.Value2)
                .ActiveCell.Offset(1, 0).Value2 = post.Books(b).AskingPrice 'asking price
                IO.File.Delete(tmpPic)
                .ActiveCell.Offset(2, 0).Value2 = post.Books(b).BuybackAmount
                If post.Books(b).BuybackAmount > highestValue Then highestValue = post.Books(b).BuybackAmount
                ThisAddIn.AppExcel.ActiveCell.Offset(2, 0).Hyperlinks.Add(Anchor:= .ActiveCell.Offset(2, 0), Address:=post.Books(b).BuybackLink, TextToDisplay:="'" & .ActiveCell.Offset(2, 0).Value)

                'image
                If post.Books(b).ImageURL <> "None" And post.Books(b).ImageURL <> "" Then
                    'If Download_File(ThisAddIn.ImageURL, tmpPic) Then
                    '    .ActiveSheet.Shapes.AddPicture(fileName:=tmpPic, Left:= .Round(.Columns(.ActiveCell.Offset(0, 1).Column).Left, 0) + 10, Top:= .Round(.Rows(.ActiveCell.Offset(-1, 0).Row).Top, 0) + 10, Width:=192, Height:=240)
                    'Else
                    '    .ActiveSheet.Shapes.AddPicture(fileName:=phPic, Left:= .Round(.Columns(.ActiveCell.Offset(0, 1).Column).Left, 0) + 10, Top:= .Round(.Rows(.ActiveCell.Offset(-1, 0).Row).Top, 0) + 10, Width:=192, Height:=240)
                    'End If
                Else
                    .ActiveSheet.Shapes.AddPicture(fileName:=phPic, Left:= .Round(.Columns(.ActiveCell.Offset(0, 1).Column).Left, 0) + 10, Top:= .Round(.Rows(.ActiveCell.Offset(-1, 0).Row).Top, 0) + 10, Width:=192, Height:=240)
                End If

                'profit calcs
                .ActiveCell.Offset(3, 0).FormulaLocal = "=IFERROR(" & .ActiveCell.Offset(2, 0).Address & "-" & .ActiveCell.Offset(1, 0).Address & ",0)"
                .ActiveCell.Offset(4, 0).FormulaLocal = "=IFERROR(" & .ActiveCell.Offset(3, 0).Address & "/" & .ActiveCell.Offset(1, 0).Address & ",0)"
                .ActiveCell.Offset(5, 0).FormulaLocal = "=IFERROR(IF(ROUND(" & .ActiveCell.Offset(2, 0).Address & "-$f$1,0)>0,ROUND(" & .ActiveCell.Offset(2, 0).Address & "-$f$1,0),0),""-"")"
                If .ActiveCell.Offset(3, 0).Value2 > .Range("h7").Value2 Then .Range("h7").Value2 = .ActiveCell.Offset(3, 0).Value2

                'flags
                .ActiveCell.Offset(6, 0).Value2 = post.Books(b).aLaCarte
                .ActiveCell.Offset(7, 0).Value2 = post.Books(b).isWeirdEdition
                .ActiveCell.Offset(8, 0).Value2 = post.Books(b).isPDF
                .ActiveCell.Offset(9, 0).Value2 = post.Books(b).isOBO

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

                .ActiveCell.Offset(11, 0).Value2 = post.Books(b).SaleDescInPost
            Next b
        End With
    End Sub

    Private Shared Function parsedBookCount() As Integer
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

End Class
