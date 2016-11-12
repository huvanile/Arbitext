Public Class BuildWSMaybes
    Public Sub buildWSMaybes()
        createWS "Maybes"
        Range("a3") = "Date Posted"
        Range("b3") = "Craigslist Post URL"
        Range("c3") = "Post Title"
        Range("d3") = "ISBN"
        Range("e3") = "Asking" & Chr(10) & "Price"
        Range("f3") = "Online selling price"
        Range("g3") = "Profit"
        Range("h3") = "Profit Margin"
        Range("i3") = "Min. Asking Price for desired profit margin"
        Range("j3") = "City"
        Range("k3") = "Date" & Chr(10) & "updated"
        Range("l3") = "Delta b/w my price and theirs"
        standardColumnTitles Range("a3:l3")
        Columns("a").ColumnWidth = 20
        Columns("b").ColumnWidth = 50
        Columns("c").ColumnWidth = 70
        Columns("d").ColumnWidth = 20
        Columns("e").ColumnWidth = 10
        Columns("f").ColumnWidth = 15
        Columns("g").ColumnWidth = 10
        Columns("h").ColumnWidth = 11
        Columns("i").ColumnWidth = 23
        Columns("j").ColumnWidth = 10
        Columns("k").ColumnWidth = 12
        Columns("l").ColumnWidth = 18
        Rows("4").Select
        ActiveWindow.FreezePanes = True
        standardPageTitle "Trash Bin"
        Range("a3:l3").AutoFilter
        Range("A1").Activate
        Rows("3").RowHeight = 30
    End Sub


End Class
