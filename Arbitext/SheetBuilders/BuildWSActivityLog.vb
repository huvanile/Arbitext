Imports Arbitext.ExcelHelpers
Imports Microsoft.Office.Interop.Excel

Public Class BuildWSActivityLog
    Public Shared Sub buildWSActivityLog()
        createWS("Activity Log")
        With ThisAddIn.AppExcel
            .Columns("a").ColumnWidth = 2
            .Columns("b").ColumnWidth = 20
            .Columns("c").ColumnWidth = 13
            .Columns("d").ColumnWidth = 20
            .Columns("e").ColumnWidth = 11
            .Columns("f").ColumnWidth = 12
            .Columns("g").ColumnWidth = 13
            .Columns("h").ColumnWidth = 13
            .Columns("i").ColumnWidth = 25
            .Columns("j").ColumnWidth = 12
            .Columns("k").ColumnWidth = 12
            .Columns("l").ColumnWidth = 16
            .Columns("m").ColumnWidth = 14
            .Columns("n").ColumnWidth = 14
            .Columns("o").ColumnWidth = 33
            .Columns("p").ColumnWidth = 9
            .Columns("q").ColumnWidth = 11
            .Columns("r").ColumnWidth = 14
            .Columns("s").ColumnWidth = 20
            .Columns("t").ColumnWidth = 45
            standardPageTitle("Activity Log")
            standardColumnTitles(.Range("b5:t5"))
            .Range("b5").Value2 = "Trade-in" & Chr(10) & "Placed Date"
            .Range("c5").Value2 = "$ Paid by Buyer"
            .Range("d5").Value2 = "Book Count"
            .Range("e5").Value2 = "Book" & Chr(10) & "Cost"
            .Range("f5").Value2 = "Other" & Chr(10) & "Costs"
            .Range("g5").Value2 = "Profit"
            .Range("h5").Value2 = "Profit Margin"
            .Range("i5").Value2 = "Status"
            .Range("j5").Value2 = "Drive Time (Hrs)"
            .Range("k5").Value2 = "Search Time (Hrs)"
            .Range("l5").Value2 = "Profit per" & Chr(10) & "Hour"
            .Range("m5").Value2 = "Round-trip Miles"
            .Range("n5").Value2 = "Buyer"
            .Range("o5").Value2 = "Source"
            .Range("p5").Value2 = "Shipped Via"
            .Range("q5").Value2 = "Date Shipped"
            .Range("r5").Value2 = "Buyer Order #"
            .Range("s5").Value2 = "Shipment Tracking #"
            .Range("t5").Value2 = "Book(s) Sold"
            .Rows("6").Select
            .ActiveWindow.FreezePanes = True
            .ActiveWindow.DisplayGridlines = False
            .Range("A1").Activate()
            .Rows("2:3").Font.Italic = True
            .Rows("2:3").WrapText = False
            .Range("b3").Value2 = "Total Paid by Buyers:" : .Range("b3").HorizontalAlignment = XlHAlign.xlHAlignRight
            .Range("c3").FormulaLocal = "=SUM(C6:C10000)" : thinInnerBorder(.Range("C3"))
            .Range("D2").Value2 = "Avg. Profit per Book:" : .Range("d2:d3").HorizontalAlignment = XlHAlign.xlHAlignRight
            .Range("d3").Value2 = "Average Book Cost:"
            .Range("e2").FormulaLocal = "=SUM(G6:G10000)/SUM(D6:D10000)" : thinInnerBorder(.Range("e2:e3"))
            .Range("e3").FormulaLocal = "=SUM(E6:E10000)/SUM(D6:D10000)" : thinInnerBorder(.Range("g3"))
            .Range("f3").Value2 = "Profit:" : .Range("f3").HorizontalAlignment = XlHAlign.xlHAlignRight
            .Range("G3").FormulaLocal = "=SUM(G6:G1000)" : thinInnerBorder(.Range("g3"))
            .Range("k3").Value2 = "Average profit per hour:" : .Range("k3").HorizontalAlignment = XlHAlign.xlHAlignRight
            .Range("l3").FormulaLocal = "=SUM(L6:L10000)/SUM(J6:K10000)" : thinInnerBorder(.Range("l3"))
            thinInnerBorder(.Range("b6:t6"))
            .Range("g6").FormulaLocal = "=C6-E6-F6"
            .Range("h6").FormulaLocal = "=G6/SUM(E6:F6)"
            .Range("l6").FormulaLocal = "=IF(SUM(J6:K6)=0,G6,G6/SUM(J6:K6))"
        End With
    End Sub

End Class
