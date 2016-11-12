Imports Arbitext.RegistryHelpers
Imports Arbitext.ExcelHelpers
Imports Microsoft.Office.Interop.Excel
Imports Arbitext.ArbitextHelpers

Public Class BuildWSMultipostManualCheck
    Public Shared Sub BuildWSMultipostManualCheck()
        createWS("Multipost Manual Checks")
        standardPageTitle("Multipost Checks")
        With ThisAddIn.AppExcel
            .Columns("a").ColumnWidth = 2
            .Columns("b").ColumnWidth = 34
            .Columns("c").ColumnWidth = 52
            .Columns("d").ColumnWidth = 40
            .Columns("e").ColumnWidth = 2
            .Columns("f").ColumnWidth = 34
            .Columns("g").ColumnWidth = 51
            .Columns("h").ColumnWidth = 40
            .Range("b3").Value2 = "Craigslist Post URL"
            .Range("b4").Value2 = "Post Date and Time"
            .Range("b5").Value2 = "Post Title"
            .Range("b6").Value2 = "Post Updated"
            .Columns("h").Style = "Currency"
            .Range("b7").Value2 = "Post City"
            With .Range("b3:b7")
                .Font.Bold = True
                .HorizontalAlignment = XlHAlign.xlHAlignRight
                .VerticalAlignment = XlVAlign.xlVAlignTop
            End With
            With .Range("g2:g7")
                .Font.Bold = True
                .HorizontalAlignment = XlHAlign.xlHAlignRight
                .VerticalAlignment = XlVAlign.xlVAlignTop
            End With
            .Range("g2").Value2 = "Total cost"
            .Range("h2").FormulaLocal = "=SUM(C11,G11,C26,G26,C41,G41,C56,G56,C71,G71,C86,G86,C101,G101,C116,G116,C131,G131)"
            .Range("g3").Value2 = "Total value per Bookscouter.com"
            .Range("h3").FormulaLocal = "=SUM(C12,G12,C27,G27,C42,G42,C57,G57,C72,G72,C87,G87,C102,G102,C117,G117,C132,G132)"
            .Range("g4").Value2 = "Total profit"
            .Range("h4").FormulaLocal = "=h3-h2"
            .Range("g6").Value2 = "Combined min asking price for desired profit"
            .Range("h6").FormulaLocal = "=h3-" & ThisAddIn.MinTolerableProfit
            .Range("g7").Value2 = "Highest profit single buy in this post" : .Range("h7").Value2 = 0
            rowValues(.Range("c3:d7"))
            rowValues(.Range("h2:h4"))
            rowValues(.Range("h6:h7"))
            .Range("c3:d3").MergeCells = True
            .Range("c4:d4").MergeCells = True
            .Range("c5:d5").MergeCells = True
            .Range("c6:d6").MergeCells = True
            .Range("c7:d7").MergeCells = True
            .ActiveWindow.DisplayGridlines = False
            .Rows("8:8").Select
            .ActiveWindow.FreezePanes = True
            .Range("g2:g7").HorizontalAlignment = XlHAlign.xlHAlignRight
            bookTile(.Range("b9"), 1)
            bookTile(.Range("f9"), 2)
            bookTile(.Range("b24"), 3)
            bookTile(.Range("f24"), 4)
            bookTile(.Range("b39"), 5)
            bookTile(.Range("f39"), 6)
            bookTile(.Range("b54"), 7)
            bookTile(.Range("f54"), 8)
            bookTile(.Range("b69"), 9)
            bookTile(.Range("f69"), 10)
            bookTile(.Range("b84"), 11)
            bookTile(.Range("f84"), 12)
            bookTile(.Range("b99"), 13)
            bookTile(.Range("f99"), 14)
            bookTile(.Range("b114"), 15)
            bookTile(.Range("f114"), 16)
            bookTile(.Range("b129"), 17)
            bookTile(.Range("f129"), 18)
            .ActiveWindow.Zoom = 80
        End With
    End Sub

    Private Shared Sub bookTile(topLeft As Excel.Range, theID As Integer)
        With topLeft
            .Value = "Book " & theID
            .Font.ColorIndex = 2
            .Font.Bold = True
            .Font.Size = 16
        End With
        ThisAddIn.AppExcel.Range(topLeft.Offset(1, 0).Address & ":" & topLeft.Offset(12, 0).Address).HorizontalAlignment = XlHAlign.xlHAlignRight
        topLeft.Offset(1, 0).Value2 = "ISBN"
        topLeft.Offset(2, 0).Value2 = "Asking Price"
        topLeft.Offset(3, 0).Value2 = "Selling Price on Bookscouter.com"
        topLeft.Offset(4, 0).Value2 = "Profit"
        topLeft.Offset(5, 0).Value2 = "Profit Margin"
        topLeft.Offset(6, 0).Value2 = "Min. Asking Price for desired profit"
        topLeft.Offset(7, 0).Value2 = "Binder of paper?"
        topLeft.Offset(8, 0).Value2 = "Weird edition?"
        topLeft.Offset(9, 0).Value2 = "ebook or PDF?"
        topLeft.Offset(10, 0).Value2 = "Or best offer?"
        topLeft.Offset(11, 0).Value2 = "WINNER?"
        topLeft.Offset(12, 0).Value2 = "String"
        topLeft.Offset(12, 1).WrapText = True
        ThisAddIn.AppExcel.Rows(topLeft.Offset(12, 0).Row).RowHeight = 60
        topLeft.Offset(1, 1).NumberFormat = "0"
        topLeft.Offset(2, 1).Style = "Currency"
        topLeft.Offset(3, 1).Style = "Currency"
        topLeft.Offset(4, 1).Style = "Currency"
        topLeft.Offset(5, 1).Style = "Percent"
        topLeft.Offset(6, 1).Style = "Currency"
        With ThisAddIn.AppExcel.Range(topLeft.Offset(1, 0).Address & ":" & topLeft.Offset(12, 0).Address)
            .HorizontalAlignment = XlHAlign.xlHAlignRight
            .VerticalAlignment = XlVAlign.xlVAlignTop
        End With
        With ThisAddIn.AppExcel.Range(topLeft.Address & ":" & topLeft.Offset(13, 2).Address)
            .Interior.ColorIndex = 15
            .VerticalAlignment = XlVAlign.xlVAlignTop
            .BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium)
        End With
        With topLeft.Offset(13, 2)
            .Font.Bold = True
            .Value = "TEXTBOOK COVER per AMAZON.COM"
            .HorizontalAlignment = XlHAlign.xlHAlignCenter
        End With
        rowValues(ThisAddIn.AppExcel.Range(topLeft.Offset(1, 1).Address & ":" & topLeft.Offset(12, 1).Address))
    End Sub

End Class
