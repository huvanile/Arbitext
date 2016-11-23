Imports Renci.SshNet
Imports Arbitext.ExcelHelpers
Imports Microsoft.Office.Interop.Excel
Imports Arbitext.StringHelpers
Imports System.Windows.Forms
Imports System.IO

Public Class ArbitextHelpers

    Public Shared Sub CreateXMLIfDesired()
        If isAnyWBOpen() AndAlso (doesWSExist("HVSBs") Or doesWSExist("Winners") Or doesWSExist("Maybes")) Then
            Dim saveAsFolder As String
            Dim theCity As String = StrConv(ThisAddIn.City, VbStrConv.ProperCase)
            If MsgBox("Would you like to output XML files and results?", vbYesNoCancel, ThisAddIn.Title) = vbYes Then
                Dim dialog As New FolderBrowserDialog()
                dialog.RootFolder = Environment.SpecialFolder.Desktop
                dialog.SelectedPath = Environment.SpecialFolder.Desktop
                dialog.Description = "Select a local folder location to save the .XML files and results"
                If dialog.ShowDialog() = DialogResult.OK Then
                    saveAsFolder = dialog.SelectedPath
                    Dim rssFeed As RSSFeed
                    Dim desc As String = ""
                    Dim title As String = ""
                    Dim outfile As String = ""

                    'HANDLE WINNERS
                    If doesWSExist("Winners") Then

                        'build RSS feed
                        desc = "Profitable book deals (winners) in " & theCity & "!"
                        title = "Arbitext: " & theCity & " Winners"
                        outfile = TrailingSlash(saveAsFolder) & theCity & " Winners.xml"
                        rssFeed = New RSSFeed(title, ThisAddIn.wwwLeadsFolder & Path.GetFileName(outfile), desc, Now, "en-US", "Winners", outfile)

                        'build result PHP files
                        For r As Short = 4 To lastUsedRow("Winners")
                            With ThisAddIn.AppExcel.Sheets("Winners")
                                Dim resultPage As New ResultPage(r, "Winners", saveAsFolder)
                                resultPage = Nothing
                            End With
                        Next

                    End If

                    'HANDLE MAYBES
                    If doesWSExist("Maybes") Then

                        'build RSS feed
                        desc = "Potentially profitable book deals (maybes) in " & theCity & "!"
                        title = "Arbitext: " & theCity & " Maybes"
                        outfile = TrailingSlash(saveAsFolder) & theCity & " Maybes.xml"
                        rssFeed = New RSSFeed(title, ThisAddIn.wwwLeadsFolder & Path.GetFileName(outfile), desc, Now, "en-US", "Maybes", outfile)

                        'build result PHP files
                        For r As Short = 4 To lastUsedRow("Maybes")
                            With ThisAddIn.AppExcel.Sheets("Maybes")
                                Dim resultPage As New ResultPage(r, "Maybes", saveAsFolder)
                                resultPage = Nothing
                            End With
                        Next
                    End If

                    'HANDLE HVSBs
                    If doesWSExist("HVSBs") Then

                        'build RSS feed
                        desc = "High value stale books in " & theCity & ".  These books can be sold for a profit, but only if the seller (who hasn't been successful selling them at the current asking price) will come down on the price a bit."
                        title = "Arbitext: " & theCity & " High Value Stale Books"
                        outfile = TrailingSlash(saveAsFolder) & theCity & " High Value Stale Books.xml"
                        rssFeed = New RSSFeed(title, ThisAddIn.wwwLeadsFolder & Path.GetFileName(outfile), desc, Now, "en-US", "HVSBs", outfile)

                        'build result PHP files
                        For r As Short = 4 To lastUsedRow("HVSBs")
                            With ThisAddIn.AppExcel.Sheets("HVSBs")
                                Dim resultPage As New ResultPage(r, "HVSBs", saveAsFolder)
                                resultPage = Nothing
                            End With
                        Next

                    End If
                    rssFeed = Nothing

                    If MsgBox("File(s) created successfully." & vbCrLf & vbCrLf & "Would you also like to SFTP the XML files and results the site?", vbYesNoCancel, title) = vbYes Then
                        Using sftp As New SftpClient(ThisAddIn.SFTPUrl, ThisAddIn.SFTPUser, ThisAddIn.SFTPPass)
                            sftp.Connect()
                            sftp.ChangeDirectory(ThisAddIn.SFTPDirectory)
                            If sftp.IsConnected Then
                                For Each file As String In Directory.GetFiles(saveAsFolder)
                                    If file Like "*.xml" Or file Like "*.php*" Then
                                        Using filestream As New FileStream(file, FileMode.Open)
                                            sftp.BufferSize = 4 * 1024
                                            sftp.UploadFile(filestream, Path.GetFileName(file))
                                        End Using
                                    End If
                                Next
                            Else
                                MsgBox("SFTP Connection Error!", vbCritical, ThisAddIn.Title)
                            End If
                        End Using
                    End If

                    MsgBox("Done!", vbInformation, ThisAddIn.Title)
                End If
            End If
        Else
            MsgBox("This can only be performed from an open results workbook", vbCritical, ThisAddIn.Title)
        End If
    End Sub

    Public Shared Function bookCountFromString(theStr As String) As Integer
        Dim tmpCount As Integer = 0
        Dim tmp1 As Integer
        Dim tmp2 As Integer
        Dim tmp3 As Integer
        tmp1 = UBound(Split(LCase(theStr), "isbn"))
        tmp2 = UBound(Split(LCase(theStr), "978"))
        tmp3 = UBound(Split(LCase(theStr), "$"))
        tmpCount = tmp1
        If tmp2 > tmpCount Then tmpCount = tmp2
        If tmp3 > tmpCount Then tmpCount = tmp3
        Return tmpCount
    End Function

    Public Shared Function randomUID() As String
        randomUID = ""
        Dim rgch As String
        rgch = "abcdefgh"
        rgch = rgch & "0123456789"
        Do Until Len(randomUID) = 16
            randomUID = randomUID & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
        Loop
    End Function

    Public Shared Sub grabPHPicIfNeeded()
        If Not System.IO.File.Exists(Environ("Temp") & "\placeholder.jpg") Then
            Dim wc As New System.Net.WebClient
            wc.DownloadFile("http://datadump.thelightningpath.com/images/bookimages/PlaceholderBook.png", Environ("Temp") & "\placeholder.jpg")
            wc = Nothing
        End If
    End Sub

    Public Shared Sub rowTitles(theRange As Excel.Range)
        With theRange
            .HorizontalAlignment = XlHAlign.xlHAlignRight
            .Interior.ColorIndex = 15
        End With
        thinInnerBorder(theRange)
    End Sub

    Public Shared Sub rowValues(theRange As Excel.Range)
        theRange.Interior.ColorIndex = 6
        thinInnerBorder(theRange)
    End Sub

    Public Shared Sub unFilterTrash()
        On Error Resume Next
        ThisAddIn.AppExcel.Worksheets("Trash").AutoFilter.Sort.SortFields.Clear
        ThisAddIn.AppExcel.Worksheets("Trash").ShowAllData
        On Error GoTo 0
    End Sub

End Class
