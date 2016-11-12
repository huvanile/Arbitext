Imports Arbitext.ExcelHelpers

Public Class BuildWSKeepers
    Public Shared Sub BuildWSKeepers()
        createWS("Keepers")
        With ThisAddIn.AppExcel
            .Range("a3").Value2 = "Date Posted"
            .Range("b3").Value2 = "Craigslist Post URL"
            .Range("c3").Value2 = "Post Title"
            .Range("d3").Value2 = "ISBN"
            .Range("e3").Value2 = "Asking" & Chr(10) & "Price"
            .Range("f3").Value2 = "Online selling price"
            .Range("g3").Value2 = "Profit"
            .Range("h3").Value2 = "Profit Margin"
            .Range("i3").Value2 = "Min. Asking Price for desired profit margin"
            .Range("j3").Value2 = "City"
            .Range("k3").Value2 = "Date" & Chr(10) & "updated"
            .Range("l3").Value2 = "Delta b/w my price and theirs"
            standardColumnTitles(.Range("a3:l3"))
            .Columns("a").ColumnWidth = 20
            .Columns("b").ColumnWidth = 50
            .Columns("c").ColumnWidth = 70
            .Columns("d").ColumnWidth = 20
            .Columns("e").ColumnWidth = 10
            .Columns("f").ColumnWidth = 15
            .Columns("g").ColumnWidth = 10
            .Columns("h").ColumnWidth = 11
            .Columns("i").ColumnWidth = 23
            .Columns("j").ColumnWidth = 10
            .Columns("k").ColumnWidth = 12
            .Columns("l").ColumnWidth = 18
            .Rows("4").Select
            .ActiveWindow.FreezePanes = True
            standardPageTitle("Keepers")
            .Range("A1").Activate()
            .Rows("3").RowHeight = 30
        End With
    End Sub

End Class
