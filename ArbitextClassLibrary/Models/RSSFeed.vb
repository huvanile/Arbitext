Imports System.IO
Imports System.Xml
Imports System.Text
Imports ArbitextClassLibrary.RSSHelpers
Imports System.Net

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
        Document.AppendChild(_document.CreateNode(XmlNodeType.XmlDeclaration, Nothing, Nothing))
        Dim rootelement = _document.CreateElement("rss")
        rootelement.SetAttribute("version", "2.0")
        _document.AppendChild(rootelement)
        AddNamespace("arbitext", globals.wwwRoot & "#")
        CreateChannel()
    End Sub

    Public Sub New(url As String)
        Try
            Dim MyRssRequest As WebRequest = WebRequest.Create(url)
            Dim MyRssResponse As WebResponse = MyRssRequest.GetResponse()
            Dim MyRssStream As Stream = MyRssResponse.GetResponseStream() 'errors when feed doesnt yet exist
            _document = New XmlDocument()
            _document.Load(MyRssStream)
            MyRssRequest = Nothing
            MyRssResponse = Nothing
            MyRssStream = Nothing
        Catch
            Throw New Exception
        End Try
        _fileName = Path.GetFileName(New Uri(url).LocalPath)
        _description = _document.GetElementsByTagName("description")(0).InnerText.ToString
        _title = _document.GetElementsByTagName("title")(0).InnerText.ToString
        _link = _document.GetElementsByTagName("link")(0).InnerText.ToString
        Dim channels = _document.GetElementsByTagName("channel")
        If channels.Count = 0 Then
            CreateChannel()
        End If
    End Sub

#End Region

    Public ReadOnly Property FileName As String
        Get
            Return _fileName
        End Get
    End Property

    ''' <summary>
    ''' Returns the XML document
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Document As XmlDocument
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
                _document.Save(w)
            End Using
            Return b.ToString
        End Get
    End Property

#Region "Customized Stuff"

    Public Sub WriteXMLFile()
        If Not Directory.Exists(Path.GetDirectoryName(_fileName)) Then MkDir(Path.GetDirectoryName(_fileName))
        If File.Exists(_fileName) Then Kill(_fileName)
        File.WriteAllText(_fileName, Me.ToString)
    End Sub



#End Region

    ''' <summary>
    ''' Creates the channel in the RSS feed
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateChannel()

        'First check we haven't already created a chnanel, as there should only be one in the feed
        Dim channels = _document.GetElementsByTagName("channel")
        If channels.Count > 0 Then
            Throw New ArgumentException("Channel " & channels(0).Name & " has already been created")
        End If

        Dim mainchannel = channels(0)

        'Create the channel element
        mainchannel = _document.CreateElement("channel")
        mainchannel.AppendChild(CreateTextElement("title", _title, _document))
        mainchannel.AppendChild(CreateTextElement("link", _link, _document))
        mainchannel.AppendChild(CreateTextElement("description", _description, _document))
        mainchannel.AppendChild(CreateTextElement("language", "en-US", _document))
        mainchannel.AppendChild(CreateTextElement("lastBuildDate", FormatDateTime(Now(), vbLongDate), _document))
        _document.DocumentElement.AppendChild(mainchannel)

    End Sub

    ''' <summary>
    ''' Adds a prefix and url as a namespace to the rss root element
    ''' </summary>
    ''' <param name="prefix">The prefix for the namespace</param>
    ''' <param name="url">The reference url</param>
    ''' <remarks></remarks>
    Public Sub AddNamespace(prefix As String, url As String)
        _document.DocumentElement.SetAttribute("xmlns:" & prefix, url)
    End Sub
End Class
