Imports Arbitext.ExcelHelpers
Imports Microsoft.Office.Interop.Excel

Public Class BuildWSAutomatedChecks
    Public Shared Sub BuildWSAutomatedChecks()
        createWS("Automated Checks")
        With ThisAddIn.AppExcel
            .Columns("a").ColumnWidth = 2
            .Columns("b").ColumnWidth = 6
            .Columns("c").ColumnWidth = 18
            .Columns("d").ColumnWidth = 51
            .Columns("e").ColumnWidth = 70
            .Columns("f").ColumnWidth = 25
            .Columns("g").ColumnWidth = 13
            .Columns("h").ColumnWidth = 13
            .Columns("i").ColumnWidth = 18
            .Columns("j").ColumnWidth = 12
            .Columns("k").ColumnWidth = 24
            .Columns("l").ColumnWidth = 15
            .Columns("m").ColumnWidth = 18
            .Columns("n").ColumnWidth = 17
            standardPageTitle("Automated Checks")
            standardColumnTitles(.Range("b3:n3"))
            .Range("b3").Value2 = "No."
            .Range("c3").Value2 = "Date Posted"
            .Range("d3").Value2 = "Craigslist Post URL"
            .Range("e3").Value2 = "Post Title" & Chr(10) & "(link to Amazon buyback search)"
            .Range("f3").Value2 = "ISBN" & Chr(10) & "(link to Bookscouter)"
            .Range("g3").Value2 = "Asking" & Chr(10) & "Price"
            .Range("h3").Value2 = "Online selling price"
            .Range("i3").Value2 = "Profit"
            .Range("j3").Value2 = "Profit" & Chr(10) & "Margin"
            .Range("k3").Value2 = "Asking Price for desired min. profit (rounded)"
            .Range("l3").Value2 = "City"
            .Range("m3").Value2 = "Date Updated"
            .Range("n3").Value2 = "Delta b/w my price and theirs"
            .Range("G:G,H:H,K:K").Style = "Currency"
            .Columns("j:J").Style = "Percent"
            .Columns("I:I").Style = "Currency"
            .Columns("N:N").Style = "Percent"
            .Columns("F:F").NumberFormat = "0"
            .Rows("4").Select
            .ActiveWindow.FreezePanes = True
            .ActiveWindow.DisplayGridlines = False
            .Range("A1").Activate()
        End With
    End Sub

End Class
