Imports Arbitext.ExcelHelpers

Public Class BuildWSResults
    Public Shared Sub buildResultWS(wsName As String)
        DeleteWS(wsName)
        createWS(wsName)
        With ThisAddIn.AppExcel
            .Range("a3").Value2 = "Date Posted"
            .Range("b3").Value2 = "Date Updated"
            .Range("c3").Value2 = "Craigslist Post"
            .Range("d3").Value2 = "Book Title" & Chr(10) & "(Links to Amazon Search)"
            .Range("e3").Value2 = "ISBN"
            .Range("f3").Value2 = "Asking" & Chr(10) & "Price"
            .Range("g3").Value2 = "Buyback Price"
            .Range("h3").Value2 = "Profit"
            .Range("i3").Value2 = "Profit Margin"
            .Range("j3").Value2 = "Min. Asking Price for desired profit margin"
            .Range("k3").Value2 = "City"
            standardColumnTitles(.Range("a3:k3"))
            .Columns("a").ColumnWidth = 20
            .Columns("b").ColumnWidth = 20
            .Columns("c").ColumnWidth = 50
            .Columns("d").ColumnWidth = 50
            .Columns("e").ColumnWidth = 20
            .Columns("f").ColumnWidth = 10
            .Columns("g").ColumnWidth = 10
            .Columns("h").ColumnWidth = 10
            .Columns("i").ColumnWidth = 10
            .Columns("j").ColumnWidth = 20
            .Columns("k").ColumnWidth = 15
            .Columns("i").Style = "Percent"
            .Columns("f:h").Style = "Currency"
            .Columns("j").Style = "Currency"
            .Rows("4").Select
            On Error Resume Next 'dunno why this errors sometimes...
            .ActiveWindow.FreezePanes = True
            On Error GoTo 0
            standardPageTitle(wsName)
            .Range("A1").Activate()
            .Rows("3").RowHeight = 30
        End With
    End Sub
End Class
