Imports System.Diagnostics
Imports System.Xml

Public Class BookscouterHelpers

    Public Shared Function BSParser(Optional thatISBN As String = "") As String
        If thatISBN = "" Then thatISBN = ThisAddIn.PostISBN
        Dim wc As New Net.WebClient
        Dim bXML() As Byte = wc.DownloadData("http://api.bookscouter.com/prices.php?isbn=" & thatISBN & "&uid=" & randomUID())
        Dim sXML As String = New UTF8Encoding().GetString(bXML)
        Dim doc As XmlDocument = New XmlDocument
        Dim nodes As XmlNodeList
        doc.LoadXml(sXML)
        wc = Nothing
        bXML = Nothing
        Dim tmp As String = ""

        If doc.ChildNodes.Count > 2 Then 'its 2 if its a bad isbn, and 0 if not response at all

            'get title
            nodes = doc.GetElementsByTagName("title")
            tmp = tmp & "Title=" & nodes(0).InnerText.Trim() & vbCrLf

            'get image
            nodes = doc.GetElementsByTagName("image")
            tmp = tmp & "Image=" & nodes(0).InnerText.Trim() & vbCrLf

            'get isbn10
            nodes = doc.GetElementsByTagName("isbn10")
            tmp = tmp & "ISBN10=" & nodes(0).InnerText.Trim() & vbCrLf

            'get isbn13
            nodes = doc.GetElementsByTagName("isbn13")
            tmp = tmp & "ISBN13=" & nodes(0).InnerText.Trim() & vbCrLf

            'get author
            nodes = doc.GetElementsByTagName("author")
            tmp = tmp & "Author=" & nodes(0).InnerText.Trim() & vbCrLf

            'get best buyback price
            nodes = doc.GetElementsByTagName("amount")
            tmp = tmp & "Best Buyback Price=" & nodes(0).InnerText.Trim() & vbCrLf

            'get best buyback vendor
            nodes = doc.GetElementsByTagName("vendor")
            tmp = tmp & "Best Buyback Vendor=" & nodes(0).InnerText.Trim() & vbCrLf

            'get best buyback link
            nodes = doc.GetElementsByTagName("link")
            tmp = tmp & "Best Buyback Link=" & nodes(0).InnerText.Trim() & vbCrLf

        Else
            tmp = tmp & "Title=(unknown)" & vbCrLf
            tmp = tmp & "Image=(unknown)" & vbCrLf
            tmp = tmp & "ISBN10=(unknown)" & vbCrLf
            tmp = tmp & "ISBN13=(unknown)" & vbCrLf
            tmp = tmp & "Author=(unknown)" & vbCrLf
            tmp = tmp & "Best Buyback Price=(unknown)" & vbCrLf
            tmp = tmp & "Best Buyback Vendor=(unknown)" & vbCrLf
            tmp = tmp & "Best Buyback Link=(unknown)" & vbCrLf
        End If
        nodes = Nothing
        doc = Nothing
        Return tmp
    End Function

    Public Shared Sub askBSAboutBook(Optional thatISBN As String = "")
        If thatISBN = "" Then thatISBN = ThisAddIn.PostISBN

        'reset other result-specific variables
        ThisAddIn.ImageURL = ""
        ThisAddIn.PostSellingPrice = ""
        ThisAddIn.BsBuybackLink = ""
        ThisAddIn.BsTitle = ""

        Dim resp As String = BSParser("http://api.bookscouter.com/prices.php?isbn=" & thatISBN & "&uid=" & randomUID())
        If resp.Length = 0 Then
            MsgBox("No response from BS python script! Could be a bad URL.", vbCritical, ThisAddIn.Title)
        Else
            Dim splitholder1
            splitholder1 = Split(resp, vbCrLf)

            'assign variables from string returned by the python script
            For g = LBound(splitholder1) To UBound(splitholder1)
                If splitholder1(g) Like "Image=*" And Len(splitholder1(g)) > Len("Image=") Then ThisAddIn.ImageURL = Trim(Right(splitholder1(g), Len(splitholder1(g)) - Len("Image=")))
                If splitholder1(g) Like "Best Buyback Price=*" Then ThisAddIn.PostSellingPrice = Trim(Right(splitholder1(g), Len(splitholder1(g)) - Len("Best Buyback Price=")))
                If splitholder1(g) Like "Title=*" Then ThisAddIn.BsTitle = Trim(Right(splitholder1(g), Len(splitholder1(g)) - Len("Title=")))
                If splitholder1(g) Like "Best Buyback Link=*" Then ThisAddIn.BsBuybackLink = Trim(Right(splitholder1(g), Len(splitholder1(g)) - Len("Best Buyback Link=")))
            Next g

            If ThisAddIn.ImageURL <> "" And ThisAddIn.ImageURL <> "(unknown)" Then ThisAddIn.ImageURL = Replace(ThisAddIn.ImageURL, "._SL75_", "")

            'error checking
            If ThisAddIn.PostSellingPrice = "" Then
                Debug.Print("---------------------------")
                'Debug.Print("Something went wrong when finding best buyback price using the bookscouter parser.  The command that failed was:" & vbCrLf & vbCrLf & theCommand, vbCritical, ThisAddIn.Title)
                'Debug.Print(theCommand)
                Debug.Print("---------------------------")
                ThisAddIn.PostSellingPrice = "(unknown)"
                ThisAddIn.BsTitle = "(unknown)"
                ThisAddIn.BsBuybackLink = "http://api.bookscouter.com/prices.php?isbn=" & thatISBN & "&uid=" & randomUID() & ""
                ThisAddIn.ImageURL = "(unknown)"
            End If
        End If
    End Sub

    Public Shared Function randomUID() As String
        Dim rgch As String
        rgch = "abcdefgh"
        rgch = rgch & "0123456789"
        Do Until Len(randomUID) = 16
            randomUID = randomUID & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
        Loop
    End Function
End Class
