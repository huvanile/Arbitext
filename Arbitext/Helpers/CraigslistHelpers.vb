Imports Arbitext.StringHelpers

Public Class CraigslistHelpers

    Public Shared Function isMulti(str As String) As Boolean
        Dim match As Boolean = False
        If (LCase(str) Like "*books*" And Not LCase(str) Like "*bookstore*") _
        Or LCase(str) Like "*texts*" Then
            match = True
        End If
        Return match
    End Function

    Public Shared Function isPDF(Str As String) As Boolean
        Dim match As Boolean = False
        If LCase(Str) Like "*ebook*" _
        Or LCase(Str) Like "*e-book*" _
        Or LCase(Str) Like "*e-text version*" _
        Or LCase(Str) Like "* e book*" _
        Or LCase(Str) Like "*pdf*" _
        Or LCase(Str) Like "*electronic version*" Then
            match = True
        End If
        Return match
    End Function

    Public Shared Function isOBO(str As String) As Boolean
        Dim match As Boolean = False
        If LCase(str) Like "* obo*" _
        Or LCase(str) Like "*best offer*" _
        Or LCase(str) Like "*make an offer*" _
        Or LCase(str) Like "*make me offer*" _
        Or LCase(str) Like "*me an offer*" _
        Or LCase(str) Like "*better offer*" _
        Or LCase(str) Like "*negotiable*" Then
            match = True
        End If
        Return match
    End Function

    Public Shared Function isWeirdEdition(str As String) As Boolean
        Dim match As Boolean = False
        If LCase(str) Like "*international*" _
        Or LCase(str) Like "*custom edition*" _
        Or str Like "*CCCD*" _
        Or LCase(str) Like "*teacher's edition*" _
        Or LCase(str) Like "*teachers edition*" _
        Or LCase(str) Like "*teachers' edition*" _
        Or LCase(str) Like "*teacher edition*" _
        Or LCase(str) Like "*instructor edition*" _
        Or LCase(str) Like "*instructor's edition*" _
        Or LCase(str) Like "*instructors' edition*" _
        Or LCase(str) Like "*college edition*" Then
            match = True
        End If
        Return match
    End Function

    Public Shared Function aLaCarte(str As String) As Boolean
        Dim match As Boolean = False
        If LCase(str) Like "*binder*" _
        Or LCase(str) Like "*paper cop*" _
        Or LCase(str) Like "*shrink wrap*" _
        Or LCase(str) Like "*loosseleaf*" _
        Or LCase(str) Like "*loose leaf*" _
        Or LCase(str) Like "*looseleaf*" _
        Or LCase(str) Like "*loose-leaf*" _
        Or LCase(str) Like "*value edition*" _
        Or LCase(str) Like "*student value*" _
        Or LCase(str) Like "*carte*" _
        Or LCase(str) Like "*3-hole punch*" Then
            match = True
        End If
        Return match
    End Function

    Public Shared Function getLinkFromCLSearchResults(str As String, Optional startPosition As Integer = 1) As String 'str is search result page html
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
        If Left(m, 2) = "//" Then m = "http:" & m
        Return m
    End Function

    Public Shared Function getAskingPrice(str As String) As Decimal 'str is posting page html
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

            'check again for the actual dollar sign, further down the page
            If Not IsNumeric(m) Then
                If InStr(i + 3, str, "$") <> 0 Then
                    i = InStr(i + 3, LCase(str), "$") + 1
                    m = clean(Mid(str, i, 6), False, True, True, True, True, False)
                End If
            End If

        End If

        On Error GoTo oops
        If m <> "" Then m = CInt(m)
        On Error GoTo 0
        If IsNumeric(m) Then getAskingPrice = CInt(m) Else getAskingPrice = -1
        Exit Function
oops:
        getAskingPrice = -1
    End Function

    Public Shared Function getISBN(str As String, postURL As String) 'str is posting page html
        Dim i As Integer : i = 0
        Dim splitholder
        Dim m As String : m = ""

        'parse out different variants of isbn in the ad

        If LCase(str) Like "*isbn in picture*" _
        Or LCase(str) Like "*isbn can be found in pic*" _
        Or LCase(str) Like "*pictures for isbn*" _
        Or LCase(str) Like "*isbn # included in pic*" _
        Or LCase(str) Like "*picture for isbn*" Then
            Return "(not listed in ad)"

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
            Return "(not listed in ad)"

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
            Return m
        Else

            System.Diagnostics.Debug.Print("weird ISBN:  " & m & " | post URL: " & postURL)
            Return "(unknown)"
        End If

    End Function

End Class
