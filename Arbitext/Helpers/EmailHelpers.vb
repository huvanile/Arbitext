Imports Arbitext.RegistryHelpers

Public Class EmailHelpers
    Public Shared Sub sendSilentNotification(emailBodyMessage As String, emailSubject As String)
        If ThisAddIn.EmailsOK Then
            Dim iMsg As Object
            Dim iConf As Object
            Dim strBody As String
            Dim Flds As Object
            iMsg = CreateObject("CDO.Message")
            iConf = CreateObject("CDO.Configuration")
            'iConf.Load -1    ' CDO Source Defaults
            Flds = iConf.Fields
            With Flds
                .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
                .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = ThisAddIn.EmailAddress
                .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = ThisAddIn.EmailPassword
                .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
                .Update
            End With
            With iMsg
                .Configuration = iConf
                .To = ThisAddIn.EmailAddress
                .CC = ""
                .BCC = ""
                .from = ThisAddIn.EmailAddress
                .Subject = emailSubject
                .HTMLBody = emailBodyMessage
                On Error Resume Next
                .send
                On Error GoTo 0
            End With
        End If
    End Sub

    Public Shared Function emailBodyString(post As Post, book As Book) As StringBuilder
        Dim message As StringBuilder = New StringBuilder
        If book.IsMaybe() Then
            message.AppendLine("<h2 style='color:orange; text-align:left'>..:: Textbook lead with negotiation potential found! ::..</h2>")
        ElseIf book.IsWinner() Then
            message.AppendLine("<h2 style='color:green; text-align:left'>..:: Definite Textbook Lead Found! ::..</h2>")
        Else
            message.AppendLine("<h2 style='text-align:left'>..:: Textbook Lead ::..</h2>")
        End If
        message.AppendLine("<hr/>")

        message.AppendLine("<h3 style='text-decoration: underline;'>Craigslist Post Details</h3>")
        message.AppendLine("<p><b>Post Title:</b>  " & post.Title & "</p>")
        message.AppendLine("<p><b>Post URL:</b>https: //href.li/?" & post.URL & "</p>")
        message.AppendLine("<p><b>Date Posted:</b>  " & post.PostDate & "</p>")
        message.AppendLine("<p><b>Date Last Updated:</b>  " & post.UpdateDate & "</p>")
        message.AppendLine("<p><b>City:</b>  " & post.City & "</p>")
        message.AppendLine("<hr/>")

        message.AppendLine("<h3 style='text-decoration: underline;'>Book Details</h3>")
        message.AppendLine("<p><b>Title:</b>" & book.Title & "</p>")
        message.AppendLine("<p><b>Author:</b>" & book.Author & "</p>")
        message.AppendLine("<p><b>ISBN13:</b>  <a href=""" & book.BookscouterSiteLink & """>" & book.Isbn13 & "</a></p>")
        message.AppendLine("<hr/>")

        message.AppendLine("<h3 style='text-decoration: underline;'>Financial Details</h3>")
        message.AppendLine("<p><b>Asking Price:</b>  $" & book.AskingPrice & "</p>")
        message.AppendLine("<p><b>Online Buyback Price:</b>  $" & book.BuybackAmount & " (from <a href=""" & book.BuybackLink & """>" & book.BuybackSite & "</a>)</p>")
        message.AppendLine("<p><b>Profit:</b>  $" & book.Profit & "</p>")
        message.AppendLine("<p><b>Profit Margin:</b>  " & 100 * book.ProfitPercentage & "%</p>")
        message.AppendLine("<p><b>Asking price I'd need to get minimum desired profit (rounded):</b>  $" & book.MinAskingPriceForDesiredProfit & "</p>")
        message.AppendLine("<p><b>Delta between my minimum price and their asking price:</b>  " & 100 * book.PriceDelta & "%</p>")
        message.AppendLine("<hr/>")

        message.AppendLine("<h3 style='text-decoration: underline;'>Flags</h3>")
        If book.isPDF() Then
            message.AppendLine("<p style='color:red;><b>eBook Flag? </b>Yes</p>")
        Else
            message.AppendLine("<p style='color:green;><b>eBook Flag? </b>No</p>")
        End If
        If book.isWeirdEdition() Then
            message.AppendLine("<p style='color:red;><b>Weird Edition Flag? </b>Yes</p>")
        Else
            message.AppendLine("<p style='color:green;><b>Weird Edition Flag? </b>No</p>")
        End If
        If book.aLaCarte() Then
            message.AppendLine("<p style='color:red;><b>A La Carte Edition Flag? </b>Yes</p>")
        Else
            message.AppendLine("<p style='color:green;><b>A La Carte Edition Flag? </b>No</p>")
        End If
        If book.isOBO() Then
            message.AppendLine("<p style='color:GREEN;><b>""Or Best Offer"" Flag? </b>Yes!</p>")
        Else
            message.AppendLine("<p><b>""Or Best Offer"" Flag? </b>No</p>")
        End If
        message.AppendLine("<hr/>")

        message.AppendLine("<h3 style='text-decoration: underline;'>Craigslist Description</h3>")
        message.AppendLine(book.SaleDescInPost)
        message.AppendLine("<hr/>")
        message.AppendLine("<p><img src='http://replygif.net/i/159.gif'/></p>")

        Return message
    End Function

End Class
