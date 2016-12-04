Public Class StringHelpers

    Public Shared Function randomUID() As String
        randomUID = ""
        Dim rgch As String
        rgch = "abcdefgh"
        rgch = rgch & "0123456789"
        Do Until Len(randomUID) = 16
            randomUID = randomUID & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
        Loop
    End Function

    Public Shared Function clean(str As String, deepSplits As Boolean, removeSpecials As Boolean, removeLetters As Boolean, Optional removeSingleSpaces As Boolean = True, Optional removeParens As Boolean = False, Optional removePeriods As Boolean = False, Optional isbnWork As Boolean = False)
        Dim o As Integer

        'splits on specific special characters
        If deepSplits Then
            str = splitOnValue(str, "isbn", isbnWork) 'added 2/1/15 to address the isbn %%%% isbn #### scenario
            str = splitOnValue(str, "ISBN", isbnWork) 'added 2/1/15 to address the isbn %%%% isbn #### scenario
            str = splitOnValue(str, "and", isbnWork) 'added 6/26/15
            str = splitOnValue(str, "AND", isbnWork) 'added 6/26/15
            str = splitOnValue(str, "And", isbnWork) 'added 2/1/15
            str = splitOnValue(str, "edition", isbnWork)  'added 6/26/15
            str = splitOnValue(str, "Edition", isbnWork)  'added 12/4/16
            str = splitOnValue(str, "EDITION", isbnWork) 'added 6/26/15
            str = splitOnValue(str, "or", isbnWork) 'added 2/1/15
            str = splitOnValue(str, "OR", isbnWork) 'added 12/4/16
            str = splitOnValue(str, "Or", isbnWork) 'added 12/4/16
            str = splitOnValue(str, "1st", isbnWork) 'added 12/4/16
            str = splitOnValue(str, "2nd", isbnWork) 'added 12/4/16
            str = splitOnValue(str, "3rd", isbnWork) 'added 12/4/16
            str = splitOnValue(str, "4th", isbnWork) 'added 12/4/16
            str = splitOnValue(str, "5th", isbnWork) 'added 12/4/16
            str = splitOnValue(str, "6th", isbnWork) 'added 12/4/16
            str = splitOnValue(str, "7th", isbnWork) 'added 12/4/16
            str = splitOnValue(str, "8th", isbnWork) 'added 12/4/16
            str = splitOnValue(str, "9th", isbnWork) 'added 12/4/16
            str = splitOnValue(str, "10th", isbnWork) 'added 12/4/16
            str = splitOnValue(str, "$", isbnWork)
            str = splitOnValue(str, "&", isbnWork) 'added 2/1/15
            str = splitOnValue(str, Chr(10), isbnWork)
            str = splitOnValue(str, Chr(13), isbnWork)
            str = splitOnValue(str, vbCr, isbnWork)
            str = splitOnValue(str, vbCrLf, isbnWork)
            str = splitOnValue(str, vbLf, isbnWork)
            str = splitOnValue(str, "/", isbnWork)
            str = splitOnValue(str, "(", isbnWork)
            str = splitOnValue(str, ")", isbnWork)
            str = splitOnValue(str, "|", isbnWork)
            str = splitOnValue(str, "x", isbnWork) 'believe it or not, this lets me get isbn's from the post title
        End If

        'remove special characters
        If removeSpecials Then
            For o = 1 To 31
                str = Replace(str, Chr(o), "") 'https://support.office.com/en-ie/article/insert-ASCII-or-Unicode-Latinbased-symbols-and-characters-d13f58d3-7bcb-44a7-a4d5-972ee12e50e0
            Next o
            str = Replace(str, vbTab, "")   'tab
            str = Replace(str, vbNewLine, "")
            str = Replace(str, "$", "")
            str = Replace(str, "<", "") 'reordered from bottom 6/25/15
            str = Replace(str, ">", "") 'reordered from bottom 6/25/15
            str = Replace(str, Chr(10), "")
            str = Replace(str, Chr(13), "")
            str = Replace(str, vbCr, "")
            str = Replace(str, vbCrLf, "")
            str = Replace(str, vbLf, "")
            str = Replace(str, "/", "")
            str = Replace(str, "\", "")
            str = Replace(str, "*", "")
            str = Replace(str, "  ", "")
            str = Replace(str, "   ", "")
            str = Replace(str, "-", "")
            str = Replace(str, ":", "")
            str = Replace(str, ";", "")
            str = Replace(str, ",", "")
            str = Replace(str, "'", "")
            str = Replace(str, "[", "")
            str = Replace(str, "]", "")
            str = Replace(str, "=", "") 'added 2/1/15
            str = Replace(str, """", "")
            'str = Replace(str, ".", "") 'removed 6/23/15
            str = Replace(str, "|", "")
            str = Replace(str, "#", "")
            'str = Replace(str, "&", "") 'removed 2/1/15
            str = Replace(str, "Â", "") 'added 6/25/15
        End If

        If removeSingleSpaces Then
            str = Replace(str, " ", "")
        End If

        If removeParens Then
            str = Replace(str, "(", "")
            str = Replace(str, ")", "")
        End If

        'remove all letters
        If removeLetters Then
            'remove A through W
            With CreateObject("vbscript.regexp")
                .Pattern = "[A-Wa-w]"
                .Global = True
                str = .Replace(str, "")
            End With

            'remove Y through Z (keep X intact because some ISBN's end with X)
            With CreateObject("vbscript.regexp")
                .Pattern = "[Y-Zy-z]"
                .Global = True
                str = .Replace(str, "")
            End With
        End If

        If removePeriods Then
            str = Replace(str, ".", "")
        End If
        str = Trim(str)
        If Left(str, 3) = "987" Then str = "978" & Right(str, Len(str) - 3) 'added 12/4/16 to catch when people swap the first few numbers of the isbn
        clean = Trim(str)
    End Function


    Public Shared Function splitOnValue(theStr As String, theVal As String, Optional isbnWork As Boolean = False)
        Dim sHolder

        Do While theStr Like "*" & theVal & "*"

            'starts with it
            If theStr Like theVal & "*" Then
                theStr = Left(theStr, Len(theStr) - Len(theVal))
            End If

            'in middle
            If theStr Like "*" & theVal & "*" Then
                sHolder = Split(theStr, theVal)
                For g = LBound(sHolder) To UBound(sHolder)
                    If g = 0 Then theStr = Trim(sHolder(0))
                    If isbnWork Then
                        If Len(sHolder(g)) = 10 _
                        Or Len(sHolder(g)) = 11 _
                        Or Len(sHolder(g)) = 13 Then
                            theStr = sHolder(g)
                            Exit For
                        ElseIf Len(clean(sHolder(g), False, True, True, True, True, True, False)) = 10 _
                        Or Len(clean(sHolder(g), False, True, True, True, True, True, False)) = 11 _
                        Or Len(clean(sHolder(g), False, True, True, True, True, True, False)) = 13 Then
                            theStr = sHolder(g)
                            Exit For
                        End If
                    Else
                        If Len(sHolder(g)) >= Len(theStr) And IsNumeric(sHolder(g)) Then theStr = sHolder(g)
                    End If

                Next g
            Else
                theStr = Trim(theStr)
            End If

            'ends with it
            If theStr Like "*" & theVal Then
                theStr = Right(theStr, Len(theStr) - Len(theVal))
            End If

        Loop

        splitOnValue = Trim(theStr)
    End Function

    Public Shared Function replacePlusWithSpace(theStr As String) As String
        Return Replace(theStr, " ", "+")
    End Function

    Public Shared Function replaceSpacesWithTwenty(theSTr As String) As String
        Return Replace(theSTr, " ", "%20")
    End Function

    Public Shared Function TrailingSlash(strFolder As String) As String
        If Len(strFolder) > 0 Then
            If Right(strFolder, 1) = "\" Then
                TrailingSlash = strFolder
            Else
                TrailingSlash = strFolder & "\"
            End If
        Else
            MsgBox("Error in trailingslash function!")
            TrailingSlash = ""
        End If
    End Function

End Class
