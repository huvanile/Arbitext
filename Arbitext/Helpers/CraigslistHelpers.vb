Imports Arbitext.StringHelpers
Imports mshtml

Public Class CraigslistHelpers

    Public Shared Function getLinkFromCLSearchResults(str As String, Optional startPosition As Integer = 1) 'str is search result page html
        Dim z As Long : z = 0
        Dim splitholder
        Dim m As String : m = ""
        z = Strings.InStr(startPosition, LCase(str), ThisAddIn.ResultHook) + Len(ThisAddIn.ResultHook) 'start at different locations "startPosition" to get different results
        m = Right(str, Len(str) - z)
        splitholder = Split(m, "class=""result-image ") 'boundary
        m = Trim(splitholder(0))
        splitholder = Split(m, "href=")
        m = Trim(splitholder(UBound(splitholder)))
        m = Right(m, Len(m) - 1)
        m = Left(m, Len(m) - 1)
        getLinkFromCLSearchResults = m
    End Function

    Public Shared Function getAskingPrice(str As String) As String 'str is posting page html
        Dim i As Integer : i = 0
        Dim m As String : m = ""

        'check for encoded dollar sign
        If InStr(1, str, "&#x0024;") <> 0 Then
            i = InStr(1, LCase(str), "&#x0024;") + 8
            m = clean(Mid(str, i, 4), False, True, True, True, True, False)
        End If

        If Not IsNumeric(m) Then
            'check for actual dollar sign
            If InStr(1, str, "$") <> 0 Then
                i = InStr(1, LCase(str), "$") + 1
                m = clean(Mid(str, i, 6), False, True, True, True, True, False)
            End If
        End If
        On Error GoTo oops
        If m <> "" Then m = CInt(m)
        On Error GoTo 0
        If IsNumeric(m) Then getAskingPrice = m Else getAskingPrice = "(unknown)"
        Exit Function
oops:
        getAskingPrice = "(unknown)"
    End Function

    Public Shared Function CLPostParser(Optional thePostURL As String = "") As String
        If thePostURL = "" Then thePostURL = ThisAddIn.PostURL Else ThisAddIn.PostURL = thePostURL
        Dim tmp As String = ""
        Dim wc As New Net.WebClient
        Dim bHTML() As Byte = wc.DownloadData(thePostURL)
        Dim sHTML As String = New UTF8Encoding().GetString(bHTML)
        Dim doc As IHTMLDocument2 = New HTMLDocument
        doc.clear()
        doc.write(sHTML)
        Dim allElements As IHTMLElementCollection = doc.all
        Dim element As IHTMLElement

        'find title
        tmp = tmp & "Title=" & doc.title & vbCrLf

        'find price
        Dim allSpans As IHTMLElementCollection = allElements.tags("span")
        For Each element In allSpans
            If element.className = "price" Then
                tmp = tmp & "Price=" & Trim(element.innerText) & vbCrLf
                Exit For
            End If
        Next
        allSpans = Nothing

        'find location
        Dim allSmalls As IHTMLElementCollection = allElements.tags("small")
        For Each element In allSmalls
            Dim tmpLoc As String = Trim(element.innerText)
            tmpLoc = Replace(tmpLoc, ")", "")
            tmpLoc = Replace(tmpLoc, "(", "")
            tmpLoc = Replace(tmpLoc, "\n", "")
            tmpLoc = Replace(tmpLoc, "\r", "")
            tmp = tmp & "Location=" & tmpLoc & vbCrLf
            Exit For
        Next
        allSmalls = Nothing

        'find og:image
        Dim metaTag As HTMLMetaElement
        For Each element In allElements
            If element.tagName = "META" Then
                metaTag = element
                If LCase(metaTag.content) Like "*images.craigslist*" Then
                    tmp = tmp & "Image=" & (metaTag.content) & vbCrLf
                    Exit For
                End If
            End If
        Next
        metaTag = Nothing

        Dim allTimes As IHTMLElementCollection = allElements.tags("time")

        'find date posted
        For Each element In allTimes
            If LCase(element.parentElement.innerText) Like "*posted:*" Then
                tmp = tmp & "Date Posted=" & Trim(element.innerText) & vbCrLf
                Exit For
            End If
        Next

        'find date updated
        For Each element In allTimes
            If LCase(element.parentElement.innerText) Like "*updated:*" Then
                tmp = tmp & "Date Updated=" & Trim(element.innerText) & vbCrLf
                Exit For
            End If
        Next

        allTimes = Nothing

        'get posting body
        Dim allSections As IHTMLElementCollection = allElements.tags("section")
        For Each element In allSections
            If element.id = "postingbody" Then
                tmp = tmp & "Posting Body=" & Trim(element.innerText) & vbCrLf
                Exit For
            End If
        Next
        allSections = Nothing

        'clean up
        doc.close()
        element = Nothing
        doc = Nothing
        wc = Nothing
        bHTML = Nothing
        allElements = Nothing
        Return tmp
    End Function

    Public Shared Sub learnAboutPost(Optional thePostURL As String = "")
        If thePostURL = "" Then thePostURL = ThisAddIn.PostURL Else ThisAddIn.PostURL = thePostURL

        'reset other result-specific variables
        ThisAddIn.PostISBN = ""           'CL post ISBN
        ThisAddIn.ALaCarteFlag = False
        ThisAddIn.WeirdEditionFlag = False
        ThisAddIn.PdfFlag = False
        ThisAddIn.OboFlag = False
        ThisAddIn.PostAskingPrice = "?"
        ThisAddIn.PostDate = "?"
        ThisAddIn.PostUpdateDate = "-"
        ThisAddIn.PostTitle = "?"
        ThisAddIn.PostCity = "-"
        ThisAddIn.PostBody = ""
        ThisAddIn.PostCLImg = ""

        'parse the post using python
        Dim resp As String = CLPostParser(ThisAddIn.PostURL)
        Dim splitholder1
        Dim splitholder2

        If resp.Length = 0 Then
            MsgBox("No response from the CL python script! Could be a bad URL.", vbCritical, ThisAddIn.Title)
            'End
        ElseIf resp.Length = 10 Then
            Diagnostics.Debug.Print("Incomplete response from the CL python script: " & ThisAddIn.PostURL)
        Else

            splitholder1 = Split(resp, vbCrLf)

            'assign variables from string returned by the post parser
            For g = LBound(splitholder1) To UBound(splitholder1)
                If splitholder1(g) Like "Location=*" And Len(splitholder1(g)) > Len("Location=") Then ThisAddIn.PostCity = Trim(Right(splitholder1(g), Len(splitholder1(g)) - Len("Location=")))
                If splitholder1(g) Like "Title=*" Then ThisAddIn.PostTitle = Trim(Right(splitholder1(g), Len(splitholder1(g)) - Len("Title=")))
                If splitholder1(g) Like "Image=*" Then ThisAddIn.PostCLImg = Trim(Right(splitholder1(g), Len(splitholder1(g)) - Len("Image=")))
                If splitholder1(g) Like "Date Updated=*" Then ThisAddIn.PostUpdateDate = Trim(Right(splitholder1(g), Len(splitholder1(g)) - Len("Date Updated=")))
                If splitholder1(g) Like "Date Posted=*" Then ThisAddIn.PostDate = Trim(Right(splitholder1(g), Len(splitholder1(g)) - Len("Date Posted=")))
                If splitholder1(g) Like "Price=*" Then ThisAddIn.PostAskingPrice = Trim(Right(splitholder1(g), Len(splitholder1(g)) - Len("Price=")))
                If splitholder1(g) Like "Posting Body=*" Then
                    splitholder2 = Split(resp, "Posting Body=")
                    ThisAddIn.PostBody = Trim(splitholder2(1))
                    Exit For 'post body is always the last tag in the parser result, so we can safely exit this for here
                End If
            Next g
            ThisAddIn.PostISBN = getISBN(ThisAddIn.PostTitle)
            If ThisAddIn.PostISBN = "" Then ThisAddIn.PostISBN = getISBN(ThisAddIn.PostBody)
            If ThisAddIn.PostAskingPrice <> "" Then ThisAddIn.PostAskingPrice = getAskingPrice(ThisAddIn.PostAskingPrice)
            If ThisAddIn.PostBody = "" Then System.Diagnostics.Debug.Print(vbCrLf & "----------------" & vbCrLf & "Empty post body:  " & ThisAddIn.PostURL & vbCrLf & "----------------" & vbCrLf)
        End If
    End Sub

    Public Shared Function getISBN(str As String) 'str is posting page html
        Dim i As Integer : i = 0
        Dim splitholder
        Dim m As String : m = ""

        'parse out different variants of isbn in the ad

        If LCase(str) Like "*isbn in picture*" _
        Or LCase(str) Like "*isbn can be found in pic*" _
        Or LCase(str) Like "*pictures for isbn*" _
        Or LCase(str) Like "*isbn # included in pic*" _
        Or LCase(str) Like "*picture for isbn*" Then
            getISBN = "(not listed in ad)"
            Exit Function

        ElseIf LCase(str) Like "*isbn*" Then

            '--------- ISBN 13

            If LCase(str) Like "*13 digit isbn*" Then
                i = InStr(1, LCase(str), "13 digit isbn") + Len("13 digit isbn")

            ElseIf LCase(str) Like "*isbn-13: text*" Then
                i = InStr(1, LCase(str), "isbn-13: text") + Len("isbn-13: text")

            ElseIf LCase(str) Like "*isbn 13*" Then
                i = InStr(1, LCase(str), "isbn 13") + Len("isbn 13")

            ElseIf LCase(str) Like "*isbn (13)*" Then
                i = InStr(1, LCase(str), "isbn (13)") + Len("isbn (13)")

            ElseIf LCase(str) Like "*isbn - 13*" Then
                i = InStr(1, LCase(str), "isbn - 13") + Len("isbn - 13")

            ElseIf LCase(str) Like "*isbn -13*" Then
                i = InStr(1, LCase(str), "isbn -13") + Len("isbn -13")

            ElseIf LCase(str) Like "*isbn13*" Then
                i = InStr(1, LCase(str), "isbn13") + Len("isbn13")

            ElseIf LCase(str) Like "*isbn- 13*" Then
                i = InStr(1, LCase(str), "isbn- 13") + Len("isbn- 13")

            ElseIf LCase(str) Like "*isbn : 13*" Then
                i = InStr(1, LCase(str), "isbn : 13") + Len("isbn : 13")

            ElseIf LCase(str) Like "*isbn:13*" Then
                i = InStr(1, LCase(str), "isbn:13") + Len("isbn:13")

            ElseIf LCase(str) Like "*isbn: 13*" Then
                i = InStr(1, LCase(str), "isbn: 13") + Len("isbn: 13")

            ElseIf LCase(str) Like "*isbn-13*" Then
                i = InStr(1, LCase(str), "isbn-13") + Len("isbn-13")

                '--------- ISBN 13 w/ wildcards

            ElseIf LCase(str) Like "*isbn ? 13*" Then 'used ? as a sub for # to get it to work
                i = InStr(1, LCase(str), "isbn # 13") + Len("isbn # 13")

            ElseIf LCase(str) Like "*isbn?:13*" Then  'added 2/1/15  'used ? as a sub for # to get it to work
                i = InStr(1, LCase(str), "isbn#:13") + Len("isbn#:13")

            ElseIf LCase(str) Like "*isbn? 13*" Then 'added 2/1/15  'used ? as a sub for # to get it to work
                i = InStr(1, LCase(str), "isbn# 13") + Len("isbn# 13")

                '------- ISBN 10

            ElseIf LCase(str) Like "*10 digit isbn*" Then
                i = InStr(1, LCase(str), "10 digit isbn") + Len("10 digit isbn")

            ElseIf LCase(str) Like "*isbn-10: text*" Then
                i = InStr(1, LCase(str), "isbn-10: text") + Len("isbn-10: text")

            ElseIf LCase(str) Like "*isbn 10*" Then
                i = InStr(1, LCase(str), "isbn 10") + Len("isbn 10")

            ElseIf LCase(str) Like "*isbn (10)*" Then
                i = InStr(1, LCase(str), "isbn (10)") + Len("isbn (10)")

            ElseIf LCase(str) Like "*isbn - 10*" Then
                i = InStr(1, LCase(str), "isbn : 10") + Len("isbn - 10")

            ElseIf LCase(str) Like "*isbn : 10*" Then
                i = InStr(1, LCase(str), "isbn : 10") + Len("isbn : 10")

            ElseIf LCase(str) Like "*isbn-10*" Then
                i = InStr(1, LCase(str), "isbn-10") + Len("isbn-10")

            ElseIf LCase(str) Like "*isbn:10*" Then
                i = InStr(1, LCase(str), "isbn:10") + Len("isbn:10")

            ElseIf LCase(str) Like "*isbn: 10*" Then
                i = InStr(1, LCase(str), "isbn: 10") + Len("isbn: 10")

            ElseIf LCase(str) Like "*isbn- 10*" Then
                i = InStr(1, LCase(str), "isbn- 10") + Len("isbn- 10")

            ElseIf LCase(str) Like "*isbn -10*" Then
                i = InStr(1, LCase(str), "isbn -10") + Len("isbn -10")

            ElseIf LCase(str) Like "*isbn10*" Then
                i = InStr(1, LCase(str), "isbn10") + Len("isbn10")

                '------- ISBN 10 w/ wildcards

            ElseIf LCase(str) Like "*isbn?:10*" Then  'added 2/1/15  'used ? as a sub for # to get it to work
                i = InStr(1, LCase(str), "isbn#:10") + Len("isbn#:10")

            ElseIf LCase(str) Like "*isbn? 10*" Then 'added 2/1/15 'used ? as a sub for # to get it to work
                i = InStr(1, LCase(str), "isbn# 10") + Len("isbn# 10")

            ElseIf LCase(str) Like "*isbn ? 10*" Then 'used ? as a sub for # to get it to work
                i = InStr(1, LCase(str), "isbn # 10") + Len("isbn # 10")

                '---- TEXT

            ElseIf LCase(str) Like "*isbn/*" Then 'added 6/23/15
                i = InStr(1, LCase(str), "isbn/") + Len("isbn/")

            ElseIf LCase(str) Like "*/isbn/*" Then 'added 2/5/15
                i = InStr(1, LCase(str), "/isbn/") + Len("/isbn/")

            ElseIf LCase(str) Like "*isbn:  lecture book *" Then
                i = InStr(1, LCase(str), "isbn:  lecture book") + Len("isbn:  lecture book ")

            ElseIf LCase(str) Like "*isbn is *" Then
                i = InStr(1, LCase(str), "isbn is ") + Len("isbn is ")

            ElseIf LCase(str) Like "*isbn ? is *" Then 'used ? as a sub for # to get it to work
                i = InStr(1, LCase(str), "isbn # is ") + Len("isbn # is ")

            ElseIf LCase(str) Like "*isbn number is *" Then
                i = InStr(1, LCase(str), "isbn number is ") + Len("isbn number is ")

            ElseIf LCase(str) Like "*isbn for this book is *" Then
                i = InStr(1, LCase(str), "isbn for this book is ") + Len("isbn for this book is ")

            ElseIf LCase(str) Like "*isbn:*" Then
                i = InStr(1, LCase(str), "isbn:") + Len("isbn:")

            ElseIf LCase(str) Like "*isbn*" Then
                i = InStr(1, LCase(str), "isbn") + Len("isbn")

            End If

            m = clean(Mid(str, i, 24), True, True, True, True, True, True, True) '2/1/14 changed from 22 '24 is an arbitrary lentgh of the isbn space including hyphens and whatnot

        ElseIf str Like "* 978*" Then
            i = InStr(1, str, "978")
            m = clean(Mid(str, i, 24), True, True, True, True, True, True, True)

        ElseIf Len(str) < 300 And str Like "*978*" Then 'added 6/25/15 for multipost processing
            i = InStr(1, str, "978")
            m = clean(Mid(str, i, 24), True, True, True, True, True, True, True)
        Else
            'still need to write something to get ISBN from post title if present. example:  http://dallas.craigslist.org/dal/bks/4856530989.html
            getISBN = "(not listed in ad)"
            Exit Function
        End If

        'just to get any pesky lingering spaces
        On Error Resume Next
        If Not IsNumeric(Right(m, 1)) _
    And Right(m, 1) <> "X" Then
            m = Left(m, Len(m) - 1)
        End If
        On Error GoTo 0
        m = clean(m, True, True, True, True, True, True, True)

        'check resulting string composition
        If (Len(m) = 10 _
    Or Len(m) = 13) _
    And (IsNumeric(m) Or Right(m, 1) = "X") Then
            getISBN = m
        Else
            getISBN = "(unknown)"
            System.Diagnostics.Debug.Print("weird ISBN:  " & m & " | post URL: " & ThisAddIn.PostURL)
        End If

    End Function

    Public Shared Function bookCount(ByVal webpage As String) As Integer
        Dim tmp1 As Integer
        Dim tmp2 As Integer
        Dim tmp3 As Integer
        tmp1 = UBound(Split(LCase(webpage), "isbn"))
        tmp2 = UBound(Split(LCase(webpage), "978"))
        tmp3 = UBound(Split(LCase(webpage), "$"))
        bookCount = tmp1
        If tmp2 > bookCount Then bookCount = tmp2
        If tmp3 > bookCount Then bookCount = tmp3
    End Function

    Public Shared Function isMulti(ByVal theStr As String) As Boolean
        If LCase(theStr) Like "*books*" Or LCase(theStr) Like "*texts*" Then
            isMulti = True 'multipost
        Else
            isMulti = False
        End If
    End Function

    Public Shared Function isPDF(theStr As String) As Boolean
        isPDF = False
        If LCase(theStr) Like "*ebook*" _
        Or LCase(theStr) Like "*e-book*" _
        Or LCase(theStr) Like "*e-text version*" _
        Or LCase(theStr) Like "* e book*" _
        Or LCase(theStr) Like "*pdf*" _
        Or LCase(theStr) Like "*electronic version*" Then
            isPDF = True
        End If
    End Function

    Public Shared Function isOBO(theStr As String) As Boolean
        isOBO = False
        If LCase(theStr) Like "* obo*" _
        Or LCase(theStr) Like "*best offer*" _
        Or LCase(theStr) Like "*make an offer*" _
        Or LCase(theStr) Like "*make me offer*" _
        Or LCase(theStr) Like "*me an offer*" _
        Or LCase(theStr) Like "*better offer*" _
        Or LCase(theStr) Like "*negotiable*" Then
            isOBO = True
        End If
    End Function

    Public Shared Function isWeirdEdition(theStr As String) As Boolean
        isWeirdEdition = False
        If LCase(theStr) Like "*international*" _
        Or LCase(theStr) Like "*custom edition*" _
        Or LCase(theStr) Like "*teacher's edition*" _
        Or LCase(theStr) Like "*teachers edition*" _
        Or LCase(theStr) Like "*teachers' edition*" _
        Or LCase(theStr) Like "*teacher edition*" _
        Or LCase(theStr) Like "*instructor edition*" _
        Or LCase(theStr) Like "*instructor's edition*" _
        Or LCase(theStr) Like "*instructors' edition*" _
        Or LCase(theStr) Like "*college edition*" Then
            isWeirdEdition = True
        End If
    End Function

    Public Shared Function aLaCarte(theStr As String) As Boolean
        aLaCarte = False
        If LCase(theStr) Like "*binder*" _
        Or LCase(theStr) Like "*paper cop*" _
        Or LCase(theStr) Like "*shrink wrap*" _
        Or LCase(theStr) Like "*loosseleaf*" _
        Or LCase(theStr) Like "*loose leaf*" _
        Or LCase(theStr) Like "*looseleaf*" _
        Or LCase(theStr) Like "*loose-leaf*" _
        Or LCase(theStr) Like "*value edition*" _
        Or LCase(theStr) Like "*student value*" _
        Or LCase(theStr) Like "*carte*" _
        Or LCase(theStr) Like "*3-hole punch*" Then
            aLaCarte = True
        End If
    End Function
End Class
