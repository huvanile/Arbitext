Imports Arbitext.ExcelHelpers
Imports System.IO
Imports System.Xml
Imports Arbitext.StringHelpers

''' <summary>
''' Creates a simple RSS feed for blogs/articles etc
''' http://www.codeproject.com/Articles/516625/CreatingplusanplusRSSplus-plusfeedpluswithplus-N
''' </summary>
''' <remarks></remarks>
Public Class RSSFeed

    'RSS channel-level attributes
    Private _fileName As String
    Private _document As XmlDocument
    Private _title As String
    Private _link As String
    Private _description As String
    Private _resultType As String

#Region "Constructors"

    ''' <summary>
    ''' Default Constructor
    ''' </summary>
    ''' <param name="title">RSS Feed Title (e.g. Arbitext: Dallas Maybes)</param>
    ''' <param name="link">RSS Feed Link</param>
    ''' <param name="description">RSS Feed Description</param>
    ''' <param name="resultType">Winners, Maybes, or HVSBs</param>
    ''' <param name="fileName">Full path and filename of new XML file on the local machine</param>
    Public Sub New(title As String, link As String, description As String, resultType As String, fileName As String)
        _document = New XmlDocument
        _title = title
        _link = link
        _description = description
        _resultType = resultType
        _fileName = fileName
        Document.AppendChild(Document.CreateNode(XmlNodeType.XmlDeclaration, Nothing, Nothing))
        Dim rootelement = Document.CreateElement("rss")
        rootelement.SetAttribute("version", "2.0")
        Document.AppendChild(rootelement)
        AddNamespace("arbitext", ThisAddIn.wwwRoot & "#")
        CreateChannel()
    End Sub

#End Region

#Region "Readonly Properties"

    ''' <summary>
    ''' Returns the XML document
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private ReadOnly Property Document As XmlDocument
        Get
            Return _document
        End Get
    End Property

    ''' <summary>
    ''' Returns a UTF8 string representation of the XML document
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shadows ReadOnly Property ToString As String
        Get
            Dim b = New StringBuilder
            Using w = New UTF8StringWriter(b)
                Document.Save(w)
            End Using
            Return b.ToString
        End Get
    End Property

#End Region

#Region "Customized Stuff"

    Public Sub PopulateFeedFromSheet()
        If isAnyWBOpen() Then
            If doesWSExist(_resultType) Then
                For r As Short = 4 To lastUsedRow(_resultType)
                    With ThisAddIn.AppExcel.Sheets(_resultType)
                        Dim dateUpdated As String = .range("b" & r).value2 'pubDate 
                        Dim postTitle As String = .range("c" & r).value2 'arbitext:postTitle
                        Dim postLink As String = "https://href.li/?" & .range("c" & r).hyperlinks(1).address 'arbitext:postLink 
                        Dim postCity As String = .range("k" & r).value2 'arbitext:postCity 
                        Dim bookTitle As String = .range("d" & r).value2 'arbitext:bookTitle
                        Dim isbn As String = .range("e" & r).value2 'arbitext:isbn
                        Dim askingPrice As Decimal = .range("f" & r).value2 'arbitext:askingprice
                        Dim bsLink As String = .range("e" & r).hyperlinks(1).address 'arbitext:buybackLink
                        Dim buybackPrice As Decimal = .range("g" & r).value2 'arbitext:buybackPrice
                        Dim profit As Decimal = .range("h" & r).value2 'arbitext:profit
                        Dim profitMargin As Decimal = .range("i" & r).value2 'arbitext:profitMargin
                        Dim id As String = .range("l" & r).value2 'GUID 
                        Dim resultURL As String = ThisAddIn.wwwRoot & "showitem.php?item=" & replaceSpacesWithTwenty(Path.GetFileName(_fileName)) & "|" & id
                        Dim theDesc As String = getDesc(postCity, askingPrice, profit, buybackPrice)
                        Dim amazonBookImage As String = .range("m" & r).value2 'arbitext:bookImage 
                        Dim postImage As String = .range("n" & r).value2 'arbitext:postImage 
                        WriteRSSItem(bookTitle, resultURL, dateUpdated, theDesc, id, postLink, postTitle, postCity, bookTitle, isbn, askingPrice, bsLink, buybackPrice, profit, profitMargin, postImage, amazonBookImage)
                    End With
                Next
            End If
        Else
            MsgBox("A workbook must be open in order to populate the XML file", vbCritical, ThisAddIn.Title)
        End If
    End Sub

    Private Function getDesc(postcity As String, askingprice As Decimal, profit As Decimal, buybackPrice As Decimal) As String
        Select Case _resultType
            Case "HVSBs"
                Return "Someone in " & StrConv(postcity, VbStrConv.ProperCase) &
                    " is asking " & FormatCurrency(askingprice, 2, TriState.False) &
                    " for this book which sells online for " & FormatCurrency(buybackPrice, 2, TriState.False) &
                    ". It's been over 2 weeks and they haven't sold it. Could be profitable if negotiated."
            Case "Maybes", "Winners"
                Return "Someone In " & StrConv(postcity, VbStrConv.ProperCase) &
                    " Is asking " & FormatCurrency(askingprice, 2, TriState.False) &
                    " For this book which sells online For " & FormatCurrency(buybackPrice, 2, TriState.False) &
                    ". That's a potential profit of " & FormatCurrency(profit, 2, TriState.False) & " (maybe more if negotiated)!"
        End Select
    End Function
    Public Sub WriteRSSItem(Title As String,
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
                        Optional postImageURL As String = ThisAddIn.wwwRoot & "/img/PlaceholderBook.png",
                        Optional amazonBookImageURL As String = ThisAddIn.wwwRoot & "/img/PlaceholderBook.png",
                        Optional Content As String = Nothing) 'content could be used to show book.saleDescInPost

        'First check we haven't already created a chnanel, as there should only be one in the feed
        Dim channels = Document.GetElementsByTagName("channel")
        If channels.Count = 0 Then Throw New ArgumentException("Please create a channel first by calling CreateChannel")

        Dim mainchannel = channels(0)

        'Create an item
        Dim thisitem = Document.CreateElement("item")
        thisitem.AppendChild(CreateTextElement("title", Title))
        thisitem.AppendChild(CreateTextElement("link", Link))
        thisitem.AppendChild(CreateTextElement("guid", Guid, {New KeyValuePair(Of String, String)("isPermaLink", "false")}))
        thisitem.AppendChild(CreateTextElement("pubDate", FormatDateTime(pubDate, vbLongDate)))
        thisitem.AppendChild(addCustomNode("postLink", postLink))
        thisitem.AppendChild(addCustomNode("postTitle", postTitle))
        thisitem.AppendChild(addCustomNode("postCity", postCity))
        thisitem.AppendChild(addCustomNode("bookTitle", bookTitle))
        thisitem.AppendChild(addCustomNode("isbn", isbn))
        thisitem.AppendChild(addCustomNode("askingPrice", askingPrice))
        thisitem.AppendChild(addCustomNode("buybackLink", bsLink))
        thisitem.AppendChild(addCustomNode("buybackPrice", buybackPrice))
        thisitem.AppendChild(addCustomNode("profit", profit))
        thisitem.AppendChild(addCustomNode("profitMargin", profitMargin))
        thisitem.AppendChild(addCustomNode("postImage", postImageURL))
        thisitem.AppendChild(addCustomNode("bookImage", amazonBookImageURL))
        thisitem.AppendChild(CreateTextElement("description", Description))

        'Write the content node
        If Not Content Is Nothing Then
            Dim contentNode = Document.CreateNode(XmlNodeType.Element, "content", "encoded", Document.GetElementsByTagName("rss")(0).GetNamespaceOfPrefix("content"))
            contentNode.InnerText = Content
            thisitem.AppendChild(contentNode)
        End If

        'Append the element
        mainchannel.AppendChild(thisitem)
    End Sub

    Public Sub WriteXMLFile()
        If File.Exists(_fileName) Then Kill(_fileName)
        File.WriteAllText(_fileName, Me.ToString)
    End Sub

    Private Function addCustomNode(theName As String, theValue As String) As XmlNode
        Dim tmpNode = Document.CreateNode(XmlNodeType.Element, "arbitext", theName, Document.GetElementsByTagName("rss")(0).GetNamespaceOfPrefix("arbitext"))
        tmpNode.InnerText = theValue
        Return tmpNode
    End Function

#End Region

#Region "Helper Functions"

    ''' <summary>
    ''' Creates the channel in the RSS feed
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateChannel()

        'First check we haven't already created a chnanel, as there should only be one in the feed
        Dim channels = Document.GetElementsByTagName("channel")
        If channels.Count > 0 Then
            Throw New ArgumentException("Channel " & channels(0).Name & " has already been created")
        End If

        Dim mainchannel = channels(0)

        'Create the channel element
        mainchannel = Document.CreateElement("channel")
        mainchannel.AppendChild(CreateTextElement("title", _title))
        mainchannel.AppendChild(CreateTextElement("link", _link))
        mainchannel.AppendChild(CreateTextElement("description", _description))
        mainchannel.AppendChild(CreateTextElement("language", "en-US"))
        mainchannel.AppendChild(CreateTextElement("lastBuildDate", FormatDateTime(Now(), vbLongDate)))
        Document.DocumentElement.AppendChild(mainchannel)

    End Sub

    ''' <summary>
    ''' Adds a prefix and url as a namespace to the rss root element
    ''' </summary>
    ''' <param name="prefix">The prefix for the namespace</param>
    ''' <param name="url">The reference url</param>
    ''' <remarks></remarks>
    Public Sub AddNamespace(prefix As String, url As String)
        Document.DocumentElement.SetAttribute("xmlns:" & prefix, url)
    End Sub

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
    Private Function CreateTextElement(Name As String, Value As String) As XmlElement
        CreateTextElement = Document.CreateElement(Name)
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
    Private Function CreateTextElement(Name As String, Value As String, Attributes() As KeyValuePair(Of String, String)) As XmlElement
        Dim element = CreateTextElement(Name, Value)
        For Each item In Attributes
            element.SetAttribute(item.Key, item.Value)
        Next
        Return element
    End Function

#End Region

End Class
