Imports Arbitext.ExcelHelpers
Imports Arbitext.ArbitextHelpers
Imports Microsoft.Office.Interop.Excel

Public Class BuildWSSingleCheck
    Public Shared Sub BuildWSSingleCheck()
        'DeleteWS("Single Check")
        createWS("Single Check")
        standardPageTitle("Analyze Single Posting")
        With ThisAddIn.AppExcel
            .ActiveWindow.DisplayGridlines = False
            .Columns("a").ColumnWidth = 2
            .Columns("b").ColumnWidth = 40
            .Columns("c").ColumnWidth = 73
            .Columns("d").ColumnWidth = 3
            .Columns("e").ColumnWidth = 24
            .Columns("f").ColumnWidth = 20
            .Columns("g").ColumnWidth = 9
            .Columns("h").ColumnWidth = 4
            .Columns("i").ColumnWidth = 8
            .Columns("j").ColumnWidth = 8

            .Range("b3").Value2 = "Craigslist Post URL"
            .Range("b4").Value2 = "Post City"
            .Range("b5").Value2 = "Post Title"
            .Range("b6").Value2 = "ISBN"
            .Range("b7").Value2 = "Asking Price"
            .Range("b8").Value2 = "Selling Price on Bookscouter.com"
            .Range("b9").Value2 = ""
            .Range("b10").Value2 = "Profit"
            .Range("b11").Value2 = "Profit Margin"
            .Range("b12").Value2 = "Min. Asking Price for desired profit"
            .Range("b13").Value2 = ""
            .Range("b15").Value2 = "Multipost?"
            .Range("b16").Value2 = "Binder, loose-leaf, or student value edition?"
            .Range("b17").Value2 = "Weird edition?"
            .Range("b18").Value2 = "ebook or PDF?"
            .Range("b19").Value2 = "Or best offer?"
            .Range("b20").Value2 = "WINNER?"
            .Range("b21").Value2 = ""

            rowTitles(.Range("b3:b8")) : rowValues(.Range("c3:c8"))
            rowTitles(.Range("b10:b12")) : rowValues(.Range("c10:c12"))
            rowTitles(.Range("b15:b20")) : rowValues(.Range("c15:c20"))

            .Range("e1").Value2 = "Date Posted"
            .Range("e2").Value2 = "Date Updated"
            rowTitles(.Range("e1:e2"))
            rowValues(.Range("f1:f2"))

            With .Range("b22:c22")
                .MergeCells = True
                .Value = "Text body of craigslist post"
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
                .Interior.ColorIndex = 15
            End With
            thinInnerBorder(.Range("b22:c22"))

            With .Range("b23:c38")
                .MergeCells = True
                .VerticalAlignment = XlVAlign.xlVAlignTop
                .Interior.ColorIndex = 6
            End With
            thinInnerBorder(.Range("b23:c38"))

            With .Range("e21:f21")
                .MergeCells = True
                .Value = "TEXTBOOK COVER per AMAZON.COM"
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
                .Interior.ColorIndex = 15
            End With
            thinInnerBorder(.Range("e21:f21"))
            thinOuterBorders(.Range("e4:f20"))
            .Range("e4:f20").Interior.ColorIndex = 6

            With .Range("e40:f40")
                .MergeCells = True
                .Value = "CRAIGLIST AD PHOTO"
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
                .Interior.ColorIndex = 15
            End With
            thinInnerBorder(.Range("e40:f40"))
            thinOuterBorders(.Range("e23:f39"))
            .Range("e23:f39").Interior.ColorIndex = 6

            With .Range("b14:c14")
                .Font.Bold = True
                .Font.Underline = True
                .MergeCells = True
                .Value = "FLAGS"
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
            End With

            .Range("c6").NumberFormat = "0"
            .Range("C8").Style = "Currency"
            .Range("C7").Style = "Currency"
            .Range("C10").Style = "Currency"
            .Range("C11").Style = "Percent"
            .Range("C12").Style = "Currency"
            With .Range("C15:C20")
                .FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlEqual, Formula1:="=""YES""")
                .FormatConditions(.FormatConditions.Count).SetFirstPriority
                With .FormatConditions(1).Font
                    .Bold = True
                    .Italic = False
                    .Strikethrough = False
                    .TintAndShade = 0
                End With
                .FormatConditions(1).StopIfTrue = False
                .FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlEqual, Formula1:="=""no""")
                .FormatConditions(.FormatConditions.Count).SetFirstPriority
                With .FormatConditions(1).Font
                    .ThemeColor = XlThemeColor.xlThemeColorDark1
                    .TintAndShade = -0.249946592608417
                End With
                .FormatConditions(1).StopIfTrue = False
            End With
        End With
    End Sub

End Class
