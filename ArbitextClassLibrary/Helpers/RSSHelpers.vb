Imports System.IO
Imports System.Xml
Imports Renci.SshNet
Imports System.Net
Imports System.Text

Public Class RSSHelpers

    ''' <summary>
    ''' This loops through all entries in the feed and kills any that have been removed from craigslist
    ''' </summary>
    Public Shared Sub RemoveDeadLeads(sftp As SftpClient)
        Dim files As IEnumerable(Of Sftp.SftpFile) = sftp.ListDirectory(Globals.SftpDirectory)
        For Each file As Sftp.SftpFile In files
            If file.Name Like "*.xml*" Then
                Try
                    Dim url As String = Globals.wwwRoot & "leads/" & StringHelpers.replaceSpacesWithTwenty(file.Name)
                    Dim MyRssRequest As WebRequest = WebRequest.Create(url)
                    Dim MyRssResponse As WebResponse = MyRssRequest.GetResponse()
                    Dim MyRssStream As Stream = MyRssResponse.GetResponseStream() 'errors when feed doesnt yet exist
                    Dim doc = New XmlDocument()
                    doc.Load(MyRssStream)
                    MyRssRequest = Nothing
                    MyRssResponse = Nothing
                    MyRssStream = Nothing

                    Dim channels As XmlNodeList = doc.GetElementsByTagName("channel")
                    If channels.Count = 0 Then Throw New ArgumentException("Please create a channel first by calling CreateChannel")
                    Dim mainchannel = channels(0)
                    Dim x As XmlNodeList = doc.GetElementsByTagName("item")
                    For i As Short = (doc.GetElementsByTagName("item").Count - 1) To 0 Step -1
                        Debug.Print("Checking item " & i & " in feed: " & file.Name)
                        If x(i).ChildNodes.Count > 0 Then
                            Dim thelink As String = x(i).ChildNodes(4).InnerText
                            Using wc As New WebClient
                                Dim value As String = wc.DownloadString(thelink)
                                If LCase(value) Like "*posting has been deleted*" Then
                                    mainchannel.RemoveChild(x(i))
                                    mainchannel.ChildNodes(4).InnerText = FormatDateTime(Now(), vbLongDate)
                                End If
                            End Using
                        End If
                    Next i
                    x = Nothing
                    mainchannel = Nothing
                    channels = Nothing

                    PushUpdatedXML(doc, sftp, file.Name)
                Catch ex As Exception
                    Debug.Print("error in removing old leads")
                End Try
            End If
        Next file
    End Sub

    Public Shared Function FeedAlreadyExists(category As String, type As String, sftp As SftpClient, sftpDirectory As String, city As String)
        Dim files As IEnumerable(Of Sftp.SftpFile) = sftp.ListDirectory(sftpDirectory)
        If LCase(type) Like "*hvs*" Then type = "stale"
        If LCase(type) Like "*obo*" Then type = "best"
        For Each file As Sftp.SftpFile In files
            If file.Name Like "*.xml*" _
            AndAlso LCase(file.Name) Like "*" & LCase(city) & "*" _
            AndAlso LCase(file.Name) Like "*" & LCase(category) & "*" _
            AndAlso LCase(file.Name) Like "*" & LCase(type) & "*" Then
                files = Nothing
                Return True
            End If
        Next
        files = Nothing
        Return False
    End Function

    Public Shared Function AlreadyInRSSFeed(id As String, type As String, sftp As SftpClient, sftpDirectory As String, city As String, sftpURL As String, category As String)
        If LCase(type) Like "*hvs*" Then
            type = "stale"
        End If
        If LCase(type) Like "*obo*" Then
            type = "best"
        End If
        Dim files As IEnumerable(Of Sftp.SftpFile) = sftp.ListDirectory(sftpDirectory)
        For Each file As Sftp.SftpFile In files
            If file.Name Like "*.xml*" _
            AndAlso LCase(file.Name) Like "*" & LCase(city) & "*" _
            AndAlso LCase(file.Name) Like "*" & LCase(category) & "*" _
            AndAlso LCase(file.Name) Like "*" & LCase(type) & "*" Then
                Using wc As New WebClient
                    Dim value As String = wc.DownloadString("http://" & sftpURL & "/leads/" & file.Name)
                    If value.Contains(id) Then
                        files = Nothing
                        Return True
                    Else
                        files = Nothing
                        Return False
                    End If
                End Using
            End If
        Next
        files = Nothing
        Return False
    End Function

    Public Shared Function getDesc(resultType As String, postcity As String, askingprice As Decimal, profit As Decimal, buybackPrice As Decimal) As String
        Dim tmp As String = "Someone"
        If postcity.Trim.Length <= 3 And LCase(postcity.Trim) <> "google map" Then tmp = tmp & " in " & StrConv(postcity, VbStrConv.ProperCase).Trim
        tmp = tmp & " is asking " & FormatCurrency(askingprice, 2, TriState.False) &
                    " for this item which sells online for " & FormatCurrency(buybackPrice, 2, TriState.False) & ".  "
        Select Case resultType
            Case "HVOBOs"
                Return tmp & "BUT, they said they'd take the best offer. Could be profitable if negotiated."
            Case "HVSBs"
                Return tmp & "It's been over 2 weeks and they haven't sold it.  They might take a lowball offer."
            Case "Maybes", "Winners"
                Return tmp & "That's a potential profit of " & FormatCurrency(profit, 2, TriState.False) & " (maybe more if negotiated)!"
        End Select
        Return ""
    End Function

    Public Shared Sub PushUpdatedXML(rss As RSSFeed, sftp As SftpClient)
        Dim xmlStream As New MemoryStream
        rss.Document.Save(xmlStream)
        xmlStream.Position = 0
        sftp.BufferSize = 4 * 1024
        sftp.UploadFile(xmlStream, Path.GetFileName(rss.FileName))
    End Sub

    Public Shared Sub PushUpdatedXML(doc As XmlDocument, sftp As SftpClient, url As String)
        Dim xmlStream As New MemoryStream
        doc.Save(xmlStream)
        xmlStream.Position = 0
        sftp.BufferSize = 4 * 1024
        sftp.UploadFile(xmlStream, Path.GetFileName(url))
    End Sub

    Public Shared Sub WriteRSSItem(ByRef document As XmlDocument,
        searchTerm As String,
        link As String,
        pubDate As DateTime,
        Description As String,
        Guid As String,
        postLink As String,
        postTitle As String,
        postCity As String,
        askingPrice As Decimal,
        tpgUrl As String,
        median As Decimal,
        mean As Decimal,
        profit As Decimal,
        profitMargin As Decimal,
        isOBO As Boolean,
        Optional postImageURL As String = Globals.wwwRoot & "/img/phoneph.png")

        'First check we haven't already created a channel, as there should only be one in the feed
        Dim channels = document.GetElementsByTagName("channel")
        If channels.Count = 0 Then Throw New ArgumentException("Please create a channel first by calling CreateChannel")

        Dim mainchannel = channels(0)
        mainchannel.ChildNodes(4).InnerText = FormatDateTime(Now(), vbLongDate)

        'Create an item
        Dim thisitem = document.CreateElement("item")
        thisitem.AppendChild(CreateTextElement("title", postTitle, document))
        thisitem.AppendChild(CreateTextElement("link", link, document))
        thisitem.AppendChild(CreateTextElement("guid", Guid, {New KeyValuePair(Of String, String)("isPermaLink", "false")}, document))
        thisitem.AppendChild(CreateTextElement("pubDate", FormatDateTime(pubDate, vbLongDate), document))
        thisitem.AppendChild(addCustomNode("postLink", postLink, document))
        thisitem.AppendChild(addCustomNode("postTitle", postTitle, document))
        thisitem.AppendChild(addCustomNode("postCity", postCity, document))
        thisitem.AppendChild(addCustomNode("askingPrice", askingPrice, document))
        thisitem.AppendChild(addCustomNode("buybackLink", tpgUrl, document))
        thisitem.AppendChild(addCustomNode("median", median, document))
        thisitem.AppendChild(addCustomNode("mean", mean, document))
        thisitem.AppendChild(addCustomNode("profit", profit, document))
        thisitem.AppendChild(addCustomNode("profitMargin", profitMargin, document))
        thisitem.AppendChild(addCustomNode("postImage", postImageURL, document))
        thisitem.AppendChild(addCustomNode("itemImage", "/img/phoneph.png", document))
        thisitem.AppendChild(CreateTextElement("description", Description, document))

        'Append the element
        mainchannel.AppendChild(thisitem)
    End Sub

    Public Shared Sub WriteRSSItem(ByRef document As XmlDocument,
        Title As String,
        Link As String,
        pubDate As DateTime,
        Description As String,
        Guid As String,
        postLink As String,
        postTitle As String,
        postCity As String,
        bookTitle As String,
        isbn As String,
        askingPrice As Decimal,
        bsLink As String,
        buybackPrice As Decimal,
        profit As Decimal,
        profitMargin As Decimal,
        isOBO As Boolean,
        Optional postImageURL As String = Globals.wwwRoot & "/img/phoneph.png",
        Optional amazonItemImageURL As String = Globals.wwwRoot & "/img/phoneph.png")


        'First check we haven't already created a chnanel, as there should only be one in the feed
        Dim channels = document.GetElementsByTagName("channel")
        If channels.Count = 0 Then Throw New ArgumentException("Please create a channel first by calling CreateChannel")

        Dim mainchannel = channels(0)

        'Create an item
        Dim thisitem = document.CreateElement("item")
        thisitem.AppendChild(CreateTextElement("title", Title, document))
        thisitem.AppendChild(CreateTextElement("link", Link, document))
        thisitem.AppendChild(CreateTextElement("guid", Guid, {New KeyValuePair(Of String, String)("isPermaLink", "false")}, document))
        thisitem.AppendChild(CreateTextElement("pubDate", FormatDateTime(pubDate, vbLongDate), document))
        thisitem.AppendChild(addCustomNode("postLink", postLink, document))
        thisitem.AppendChild(addCustomNode("postTitle", postTitle, document))
        thisitem.AppendChild(addCustomNode("postCity", postCity, document))
        thisitem.AppendChild(addCustomNode("bookTitle", bookTitle, document))
        thisitem.AppendChild(addCustomNode("isbn", isbn, document))
        thisitem.AppendChild(addCustomNode("isOBO", isOBO, document))
        thisitem.AppendChild(addCustomNode("askingPrice", askingPrice, document))
        thisitem.AppendChild(addCustomNode("buybackLink", bsLink, document))
        thisitem.AppendChild(addCustomNode("buybackPrice", buybackPrice, document))
        thisitem.AppendChild(addCustomNode("profit", profit, document))
        thisitem.AppendChild(addCustomNode("profitMargin", profitMargin, document))
        thisitem.AppendChild(addCustomNode("postImage", postImageURL, document))
        thisitem.AppendChild(addCustomNode("itemImage", amazonItemImageURL, document))
        thisitem.AppendChild(CreateTextElement("description", Description, document))

        'Append the element
        mainchannel.AppendChild(thisitem)
    End Sub

    Private Shared Function addCustomNode(theName As String, theValue As String, ByRef document As XmlDocument) As XmlNode
        Dim tmpNode = document.CreateNode(XmlNodeType.Element, "arbitext", theName, document.GetElementsByTagName("rss")(0).GetNamespaceOfPrefix("arbitext"))
        tmpNode.InnerText = theValue
        Return tmpNode
    End Function

    ''' <summary>
    ''' A custom string writer which outputs in UTF8 rather than UTF16
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class UTF8StringWriter
        Inherits StringWriter
        Public Overrides ReadOnly Property Encoding As Encoding
            Get
                Return System.Text.Encoding.UTF8
            End Get
        End Property
        Public Sub New(stringBuilder As StringBuilder)
            MyBase.New(stringBuilder)
        End Sub
        Public Sub New(stringBuilder As StringBuilder, format As System.IFormatProvider)
            MyBase.New(stringBuilder, format)
        End Sub
    End Class

    ''' <summary>
    ''' Creates a simple text element with the given name and inner text value
    ''' </summary>
    ''' <param name="Name">The name of the element</param>
    ''' <param name="Value">The value of it</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CreateTextElement(Name As String, Value As String, ByRef document As XmlDocument) As XmlElement
        CreateTextElement = document.CreateElement(Name)
        CreateTextElement.InnerText = Value
    End Function

    ''' <summary>
    ''' Creates a simple text element with the given name and inner text value, whilst also settings any passed attributes
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <param name="Value"></param>
    ''' <param name="Attributes"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function CreateTextElement(Name As String, Value As String, Attributes() As KeyValuePair(Of String, String), ByRef document As XmlDocument) As XmlElement
        Dim element = CreateTextElement(Name, Value, document)
        For Each item In Attributes
            element.SetAttribute(item.Key, item.Value)
        Next
        Return element
    End Function

End Class
