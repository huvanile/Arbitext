Imports Arbitext.StringHelpers

Public Class ResultPage

    Sub New(b As Book, p As Post, resultType As String, saveAsFolder As String)
        Dim tmp As New StringBuilder
        Dim outfile As String = TrailingSlash(saveAsFolder) & b.ID & ".php"
        tmp.AppendLine(header())
        tmp.AppendLine("<h2>" & Strings.Left(resultType, Len(resultType) - 1) & " Found in " & p.City & "</h2>")
        tmp.AppendLine("<hr/>")
        tmp.AppendLine("<h3>Book Information</h3>")
        tmp.AppendLine("<p>Book Title: " & b.Title & "</p>")
        tmp.AppendLine("<p>ISBN: " & b.Isbn13 & "</p>")
        tmp.AppendLine("<p><a href='https://href.li/?" & b.AmazonSearchURL & "'</a>Amazon Link</a></p>")
        tmp.AppendLine("<img id='bookimage' src='" & b.ImageURL & "' alt='Amazon Image'/>")
        tmp.AppendLine("<hr/>")
        tmp.AppendLine("<h3>Post Information</h3>")
        tmp.AppendLine("<p>Post Title: <a href='https://href.li/?" & p.URL & "'>" & p.Title & "</a></p>")
        tmp.AppendLine("<p>Post Last Updated: " & p.UpdateDate & "</p>")
        tmp.AppendLine("<p>Asking Price for Book: $" & b.AskingPrice & "</p>")
        tmp.AppendLine("<p>Sale Description: " & b.SaleDescInPost & "</p>")
        If Not p.Image = "" Then tmp.AppendLine("<img id='postimage' src='" & p.Image & "' alt='Post Image'/>")
        tmp.AppendLine("<hr/>")
        tmp.AppendLine("<h3>Buyback Information</h3>")
        tmp.AppendLine("<p>Best Online Buyback Price: <a href='" & b.BookscouterSiteLink & "'>$" & b.BuybackAmount & "</a></p>")
        tmp.AppendLine("<p>Potential Profit: $" & b.Profit & "</p>")
        tmp.AppendLine("<p>Profit Margin: " & b.ProfitPercentage & "</p>")
        tmp.AppendLine(footer)
        'WriteToFile(outfile, tmp.ToString)
    End Sub

    Sub New(r As Short, ws As String, saveAsFolder As String)
        Dim tmp As New StringBuilder
        With ThisAddIn.AppExcel.Sheets(ws)
            Dim datePosted As String = .range("a" & r).value2
            Dim dateUpdated As String = .range("b" & r).value2
            Dim postTitle As String = .range("c" & r).value2
            Dim postLink As String = .range("c" & r).hyperlinks(1).address
            Dim city As String = .range("k" & r).value2
            Dim bookTitle As String = .range("d" & r).value2
            Dim amazonLink As String = .range("d" & r).hyperlinks(1).address
            Dim isbn As String = .range("e" & r).value2
            Dim askingPrice As Decimal = .range("f" & r).value2
            Dim bsLink As String = .rangE("e" & r).hyperlinks(1).address
            Dim buybackPrice As Decimal = .range("g" & r).value2
            Dim profit As Decimal = .range("h" & r).value2
            Dim profitMargin As Decimal = .range("i" & r).value2
            Dim id As String = .range("l" & r).value2
            Dim bookImageURL As String = .range("m" & r).hyperlinks(1).address
            Dim postImageURL As String = .range("n" & r).hyperlinks(1).address
            Dim outfile As String = TrailingSlash(saveAsFolder) & id & ".php"
            tmp.AppendLine(header())
            tmp.AppendLine("<h2>" & Strings.Left(ws, Len(ws) - 1) & " Found in " & city & "</h2>")
            tmp.AppendLine("<hr/>")
            tmp.AppendLine("<h3>Book Information</h3>")
            tmp.AppendLine("<p>Book Title: " & bookTitle & "</p>")
            tmp.AppendLine("<p>ISBN: " & isbn & "</p>")
            tmp.AppendLine("<p>Amazon Link: <a href='" & amazonLink & "'</a>" & amazonLink & "</a></p>")
            tmp.AppendLine("<hr/>")
            tmp.AppendLine("<h3>Post Information</h3>")
            tmp.AppendLine("<p>Post Title: <a href='https://href.li/?" & postLink & "'>" & postTitle & "</a></p>")
            tmp.AppendLine("<p>Post Last Updated: " & dateUpdated & "</p>")
            tmp.AppendLine("<p>Asking Price for Book: $" & askingPrice & "</p>")
            tmp.AppendLine("<hr/>")
            tmp.AppendLine("<h3>Buyback Information</h3>")
            tmp.AppendLine("<p>Best Online Buyback Price: <a href='" & bsLink & "'>$" & buybackPrice & "</a></p>")
            tmp.AppendLine("<p>Potential Profit: $" & profit & "</p>")
            tmp.AppendLine("<p>Profit Margin: " & profitMargin & "</p>")
            tmp.AppendLine(footer)
            'WriteToFile(outfile, tmp.ToString)
        End With
    End Sub

    Private Function header() As String
        Return "<?php include_once(""http://huvanile.com/includes/header.php""); ?>"
    End Function

    Private Function footer() As String
        Return "<?php include_once(""http://huvanile.com/includes/footer.php""); ?>"
    End Function
End Class
