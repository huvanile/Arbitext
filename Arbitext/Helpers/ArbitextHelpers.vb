Public Class ArbitextHelpers

    Public Shared Sub grabPHPicIfNeeded()
        If Not System.IO.File.Exists(Environ("Temp") & "\placeholder.jpg") Then
            Dim wc As New System.Net.WebClient
            wc.DownloadFile("http://datadump.thelightningpath.com/images/bookimages/PlaceholderBook.png", Environ("Temp") & "\placeholder.jpg")
            wc = Nothing
        End If
    End Sub

End Class
