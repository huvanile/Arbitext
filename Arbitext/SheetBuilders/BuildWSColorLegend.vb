Imports Arbitext.ExcelHelpers
Imports Microsoft.Office.Interop.Excel

Public Class BuildWSColorLegend
    Public Shared Sub buildWSColorLegend()
        createWS("Color Legend")
        With ThisAddIn.AppExcel
            .Columns("b").ColumnWidth = 20
            .Columns("C").ColumnWidth = 30
            standardPageTitle("Color Legend")
            .ActiveWindow.DisplayGridlines = False
            thinInnerBorder(.Range("b4:c10"))

            .Range("b3").Value2 = "Color"
            .Range("c3").Value2 = "Meaning"
            .Rows("3").Font.Bold = True

            .Range("b4").Value2 = "red text"
            .Range("c4").Value2 = "loose leaf / binder"
            .Rows("4").Font.ColorIndex = 3

            .Range("b5").Value2 = "orange text"
            .Range("c5").Value2 = "international or custom edition"
            .Rows("5").Font.ColorIndex = 46

            .Range("b6").Value2 = "'teal background"
            .Range("c6").Value2 = "'multipost"
            .Range("b6:c6").Interior.ColorIndex = 20

            .Range("b7").Value2 = "brown text"
            .Range("c7").Value2 = "ebook"
            .Rows("7").Font.ColorIndex = 30

            .Range("b8").Value2 = "'bold text"
            .Range("c8").Value2 = "'or best offer / OBO"
            .Rows("8").Font.Bold = True

            .Range("b9").Value2 = "Green cells"
            .Range("c9").Value2 = "Possible KEEPER"
            .Range("b9:c9").Interior.ColorIndex = 35

            .Range("b10").Value2 = "Yellow cells"
            .Range("c10").Value2 = "Keeper if negotiated"
            .Range("b10:c10").Interior.Color = 65535
        End With
    End Sub
End Class
