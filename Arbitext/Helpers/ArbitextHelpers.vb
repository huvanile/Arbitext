Imports Arbitext.ExcelHelpers
Imports Microsoft.Office.Interop.Excel
Imports Arbitext.StringHelpers
Imports Arbitext.FileHelpers
Imports System.Windows.Forms

Public Class ArbitextHelpers

    Public Shared Sub CreateXMLIfDesired()
        If isAnyWBOpen() Then
            If MsgBox("Would you like to output winners and maybes to an XML file?", vbYesNoCancel, ThisAddIn.Title) = vbYes Then
                Dim saveAsFolder As String
                Dim dialog As New FolderBrowserDialog()
                dialog.RootFolder = Environment.SpecialFolder.Desktop
                dialog.SelectedPath = "C:\"
                dialog.Description = "Select a location to save the .XML file"
                If dialog.ShowDialog() = DialogResult.OK Then
                    saveAsFolder = dialog.SelectedPath
                    Dim rssfeed As New RSSFeed()
                    rssfeed.CreateChannel("Aribtext", "", "Profitable book deals", Now, "en-US")
                    rssfeed.PopulateFeed()
                    WriteToFile(TrailingSlash(saveAsFolder) & ThisAddIn.City & ".xml", rssfeed.ToString)
                    MsgBox("Done!", vbInformation, ThisAddIn.Title)
                End If
            End If
        Else
            MsgBox("This can only be performed from an open results workbook", vbCritical, ThisAddIn.Title)
        End If
    End Sub

    Public Shared Function bookCountFromString(theStr As String) As Integer
        Dim tmpCount As Integer = 0
        Dim tmp1 As Integer
        Dim tmp2 As Integer
        Dim tmp3 As Integer
        tmp1 = UBound(Split(LCase(theStr), "isbn"))
        tmp2 = UBound(Split(LCase(theStr), "978"))
        tmp3 = UBound(Split(LCase(theStr), "$"))
        tmpCount = tmp1
        If tmp2 > tmpCount Then tmpCount = tmp2
        If tmp3 > tmpCount Then tmpCount = tmp3
        Return tmpCount
    End Function

    Public Shared Function randomUID() As String
        randomUID = ""
        Dim rgch As String
        rgch = "abcdefgh"
        rgch = rgch & "0123456789"
        Do Until Len(randomUID) = 16
            randomUID = randomUID & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
        Loop
    End Function

    Public Shared Sub deleteBlankAutomatedResults()
        With ThisAddIn.AppExcel
            .ScreenUpdating = False
            Dim c As Long
            For c = lastUsedRow("Automated Checks") To 5 Step -1
                If .Sheets("Automated Checks").Range("b" & c).Value = "" Then .Sheets("Automated Checks").Rows(c).Delete
            Next c
            .ScreenUpdating = True
        End With
    End Sub

    Public Shared Sub grabPHPicIfNeeded()
        If Not System.IO.File.Exists(Environ("Temp") & "\placeholder.jpg") Then
            Dim wc As New System.Net.WebClient
            wc.DownloadFile("http://datadump.thelightningpath.com/images/bookimages/PlaceholderBook.png", Environ("Temp") & "\placeholder.jpg")
            wc = Nothing
        End If
    End Sub

    Public Shared Sub rowTitles(theRange As Excel.Range)
        With theRange
            .HorizontalAlignment = XlHAlign.xlHAlignRight
            .Interior.ColorIndex = 15
        End With
        thinInnerBorder(theRange)
    End Sub

    Public Shared Sub rowValues(theRange As Excel.Range)
        theRange.Interior.ColorIndex = 6
        thinInnerBorder(theRange)
    End Sub

    Public Shared Sub reCategorizeRow(theRow As Integer, theDest As String)
        Dim callingWS As String : callingWS = ThisAddIn.AppExcel.ActiveSheet.Name
        Dim matchR As Integer
        If Not doesWSExist(theDest) Then
            MsgBox("The '" & theDest & "' sheet must be present in order to run this tool.", vbCritical, ThisAddIn.Title)
            ThisAddIn.Proceed = False
        End If
        With ThisAddIn.AppExcel
            .ScreenUpdating = False
            matchR = lastUsedRow(theDest) + 1
            .Rows(theRow).Select
            .Selection.Cut
            .Sheets(theDest).Activate
            .Rows(matchR).Select
            .ActiveSheet.Paste
            .Sheets(callingWS).Activate
            .Rows(theRow).Delete
            .ScreenUpdating = True
        End With
    End Sub

    Public Shared Sub categorizeRow(theRow As Integer, Optional theDest As String = "Trash")
        Dim tR As Integer
        Dim matchR As Integer
        Dim postTitle As String : postTitle = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("e" & theRow).Value 'title
        Dim postCity As String : postCity = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("l" & theRow).Value
        Dim postISBN As String : postISBN = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("f" & theRow).Value 'isbn
        If Not doesWSExist(theDest) Then
            MsgBox("The '" & theDest & "' sheet must be present in order to run this tool.", vbCritical, ThisAddIn.Title)
        Else
            With ThisAddIn.AppExcel.Sheets(theDest)
                If canFind(postTitle, theDest) Then
                    'match was already in trash, but showed up in results again because they changed the price and reposted it
                    'handle this situation by over-writing old row
                    matchR = ThisAddIn.AppExcel.Range(canFind(postTitle, theDest, , True, False)).Row
                    If .Range("d" & .Range(canFind(postTitle, theDest, , True, False)).Row).Value = postISBN _
                And .Range("j" & .Range(canFind(postTitle, theDest, , True, False)).Row).Value = postCity Then
                        .Range("a" & matchR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("c" & theRow).Value 'date
                        .Range("b" & matchR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("d" & theRow).Value 'url
                        .Hyperlinks.Add(anchor:= .Range("b" & matchR), Address:= .Range("b" & matchR).Value, TextToDisplay:= .Range("b" & matchR).Value)
                        .Range("c" & matchR) = postTitle
                        .Hyperlinks.Add(anchor:= .Range("c" & matchR), Address:="http://www.amazon.com/s/ref=nb_sb_noss?url=search-alias%3Daps&field-keywords=" & replacePlusWithSpace(postTitle), TextToDisplay:= .Range("c" & matchR).Value)
                        .Range("d" & matchR) = postISBN 'isbn
                        '                If Not .Range("d" & matchR) Like "*(*" Then
                        '                    .Hyperlinks.Add anchor:=.Range("d" & matchR), Address:="https://bookscouter.com/prices.php?isbn=" & postISBN & "&all", TextToDisplay:=.Range("d" & matchR).Value
                        '                End If
                        .Range("e" & matchR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("g" & theRow).Value 'price
                        .Range("f" & matchR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("h" & theRow).Value 'online buy price
                        .Range("g" & matchR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("i" & theRow).Value 'profit
                        .Range("h" & matchR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("j" & theRow).Value 'profit margin
                        .Range("i" & matchR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("k" & theRow).Value 'min. asking price for desired profit margin
                        .Range("j" & matchR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("l" & theRow).Value
                        .Range("k" & matchR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("m" & theRow).Value 'date updated
                        .Range("l" & matchR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("n" & theRow).Value 'price delta

                        'apply flags
                        If ThisAddIn.AppExcel.Sheets("Automated Checks").Range("b" & theRow).Font.Bold = True Then .Rows(matchR).Font.Bold = True
                        .Rows(matchR).Font.ColorIndex = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("b" & theRow).Font.ColorIndex
                        .Range("a" & matchR & ":l" & matchR).Interior.Color = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("b" & theRow).Interior.Color

                        ThisAddIn.AppExcel.Sheets("Automated Checks").Rows(theRow).Delete
                    Else
                        tR = lastUsedRow(theDest) + 1
                        thinInnerBorder(ThisAddIn.AppExcel.Sheets(theDest).Range("A" & tR & ":l" & tR))
                        .Range("a" & tR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("c" & theRow).Value 'date
                        .Range("b" & tR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("d" & theRow).Value 'url
                        .Hyperlinks.Add(anchor:= .Range("b" & tR), Address:= .Range("b" & tR).Value, TextToDisplay:= .Range("b" & tR).Value)
                        .Range("c" & tR) = postTitle
                        .Hyperlinks.Add(anchor:= .Range("c" & tR), Address:="http://www.amazon.com/s/ref=nb_sb_noss?url=search-alias%3Daps&field-keywords=" & replacePlusWithSpace(postTitle), TextToDisplay:= .Range("c" & tR).Value)
                        .Range("d" & tR) = postISBN 'isbn
                        If Not .Range("d" & tR) Like "*(*" Then
                            .Hyperlinks.Add(anchor:= .Range("d" & tR), Address:="https://bookscouter.com/prices.php?isbn=" & postISBN & "&all", TextToDisplay:="'" & .Range("d" & tR).Value)
                        End If
                        .Range("e" & tR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("g" & theRow).Value 'price
                        .Range("f" & tR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("h" & theRow).Value 'online buy price
                        .Range("g" & tR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("i" & theRow).Value 'profit
                        .Range("h" & tR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("j" & theRow).Value 'profit margin
                        .Range("i" & tR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("k" & theRow).Value 'min. asking price for desired profit margin
                        .Range("j" & tR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("l" & theRow).Value
                        .Range("k" & tR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("m" & theRow).Value 'date updated
                        .Range("l" & tR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("n" & theRow).Value 'price delta

                        'apply flags
                        If ThisAddIn.AppExcel.Sheets("Automated Checks").Range("b" & theRow).Font.Bold = True Then .Rows(tR).Font.Bold = True
                        .Rows(tR).Font.ColorIndex = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("b" & theRow).Font.ColorIndex
                        .Range("a" & tR & ":l" & tR).Interior.Color = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("b" & theRow).Interior.Color

                        ThisAddIn.AppExcel.Sheets("Automated Checks").Rows(theRow).Delete
                    End If
                Else
                    tR = lastUsedRow(theDest) + 1
                    thinInnerBorder(ThisAddIn.AppExcel.Sheets(theDest).Range("A" & tR & ":l" & tR))
                    .Range("a" & tR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("c" & theRow).Value 'date
                    .Range("b" & tR) = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("d" & theRow).Value 'url
                    .Hyperlinks.Add(anchor:= .Range("b" & tR), Address:= .Range("b" & tR).Value, TextToDisplay:= .Range("b" & tR).Value)
                    .Range("c" & tR) = postTitle
                    .Hyperlinks.Add(anchor:= .Range("c" & tR), Address:="http://www.amazon.com/s/ref=nb_sb_noss?url=search-alias%3Daps&field-keywords=" & replacePlusWithSpace(postTitle), TextToDisplay:= .Range("c" & tR).Value)
                    .Range("d" & tR) = postISBN 'isbn
                    If Not .Range("d" & tR).value2 Like "*(*" Then
                        .Hyperlinks.Add(anchor:= .Range("d" & tR), Address:="https://bookscouter.com/prices.php?isbn=" & postISBN & "&all", TextToDisplay:="'" & .Range("d" & tR).Value)
                    End If
                    .Range("e" & tR).value2 = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("g" & theRow).Value 'price
                    .Range("f" & tR).value2 = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("h" & theRow).Value 'online buy price
                    .Range("g" & tR).value2 = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("i" & theRow).Value 'profit
                    .Range("h" & tR).value2 = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("j" & theRow).Value 'profit margin
                    .Range("i" & tR).value2 = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("k" & theRow).Value 'min. asking price for desired profit margin
                    .Range("j" & tR).value2 = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("l" & theRow).Value
                    .Range("k" & tR).value2 = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("m" & theRow).Value 'date updated
                    .Range("l" & tR).value2 = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("n" & theRow).Value 'price delta

                    'apply flags
                    If ThisAddIn.AppExcel.Sheets("Automated Checks").Range("b" & theRow).Font.Bold = True Then .Rows(tR).Font.Bold = True
                    .Rows(tR).Font.ColorIndex = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("b" & theRow).Font.ColorIndex
                    .Range("a" & tR & ":l" & tR).Interior.Color = ThisAddIn.AppExcel.Sheets("Automated Checks").Range("b" & theRow).Interior.Color

                    ThisAddIn.AppExcel.Sheets("Automated Checks").Rows(theRow).Delete
                End If
            End With
        End If
    End Sub

    Public Shared Sub trashWSCheck()
        If Not doesWSExist("Trash") Then
            MsgBox("The Trash worksheet must be present before this tool can be run.", vbCritical, ThisAddIn.Title)
            ThisAddIn.Proceed = False
        End If
    End Sub

    Public Shared Sub singleCheckPageCheck()
        If Not ThisAddIn.AppExcel.ActiveSheet.Name = "Single Check" Then
            MsgBox("Must be on the 'Single Check' page to run this tool", vbCritical, ThisAddIn.Title)
            ThisAddIn.Proceed = False
        End If
    End Sub

    Public Shared Sub automatedChecksPageCheck()
        If Not ThisAddIn.AppExcel.ActiveSheet.Name = "Automated Checks" Then
            MsgBox("Must be on the 'Automated Checks' page to run this tool", vbCritical, ThisAddIn.Title)
            ThisAddIn.Proceed = False
        End If
    End Sub

    Public Shared Sub unFilterTrash()
        On Error Resume Next
        ThisAddIn.AppExcel.Worksheets("Trash").AutoFilter.Sort.SortFields.Clear
        ThisAddIn.AppExcel.Worksheets("Trash").ShowAllData
        On Error GoTo 0
    End Sub

    Public Shared Function bookAlreadyPresent(thePostTitle As String, theAskingPrice As Decimal, theCity As String) As Boolean
        bookAlreadyPresent = False
        Dim categories As New List(Of String)
        categories.Add("Trash")
        categories.Add("Maybes")
        categories.Add("Keepers")
        For Each c As String In categories
            With ThisAddIn.AppExcel.Sheets(c)
                If canFind(thePostTitle, c,, False, False) Then
                    Dim foundrow As Short = .range(canFind(thePostTitle, c,, True, False)).row
                    If .Range("c" & foundrow).Value2.trim = thePostTitle _
                    And .Range("f" & foundrow).Value2 = theAskingPrice _
                    And .Range("k" & foundrow).Value2.trim = theCity Then
                        Return True
                    End If
                End If
            End With
        Next
    End Function

End Class
